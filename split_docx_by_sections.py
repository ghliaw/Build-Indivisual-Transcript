import sys
import os
import zipfile
import io
import re
from lxml import etree

# WordprocessingML 主要命名空間
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}
def _rm_all(node, xpath):
    for n in node.xpath(xpath, namespaces=NS):
        parent = n.getparent()
        if parent is not None:
            parent.remove(n)

def normalize_last_section(body):
    """
    清除所有內嵌 sectPr，只保留一個放在 <w:body> 尾端，並將 type 調為 continuous。
    同時移除末端常見的分頁符與 pageBreakBefore，避免多出空白頁。
    """
    # 1) 找出所有 sectPr（保留最後一個當模板）
    sect_nodes = body.xpath(".//w:sectPr", namespaces=NS)
    sect_last = None
    if sect_nodes:
        sect_last = etree.fromstring(etree.tostring(sect_nodes[-1]))

    # 2) 砍掉 body 裡所有現存的 sectPr（包含段落內的）
    _rm_all(body, ".//w:sectPr")

    # 3) 清掉最後幾個段落的換頁符與 pageBreakBefore
    paras = body.xpath("./w:p", namespaces=NS)
    tail = paras[-3:] if len(paras) >= 3 else paras
    for p in tail:
        # 移除 <w:br w:type="page"/> 和 <w:lastRenderedPageBreak/>
        _rm_all(p, ".//w:br[@w:type='page']")
        _rm_all(p, ".//w:lastRenderedPageBreak")
        # 移除 <w:pPr><w:pageBreakBefore/>
        pPr = p.find("w:pPr", NS)
        if pPr is not None:
            for pb in pPr.findall("w:pageBreakBefore", NS):
                pPr.remove(pb)

    # 4) 若沒有找到 sectPr，就造一個最基本的
    if sect_last is None:
        sect_last = etree.Element(f"{{{NS['w']}}}sectPr")

    # 5) 把 sectPr 的 type 改為 continuous（避免 next/odd/even 造成補頁）
    type_node = sect_last.find("w:type", NS)
    if type_node is None:
        type_node = etree.SubElement(sect_last, f"{{{NS['w']}}}type")
    type_node.set(f"{{{NS['w']}}}val", "continuous")

    # 6) 將唯一的 sectPr 放到 <w:body> 末端（不是放在段落之內）
    body.append(sect_last)

def text_of(elem):
    """取段落/表格內的可見文字（用於輸出檔名提示）。"""
    if elem is None:
        return ""
    # 抓所有 <w:t>
    texts = [t.text for t in elem.xpath(".//w:t", namespaces=NS) if t.text]
    return "".join(texts).strip()

def find_section_spans(body_children):
    """
    根據 <w:sectPr> 尋找每一節的 (start_idx, end_idx)（包含 end_idx）。
    分節屬性通常位於某段落 p 的 pPr/sectPr，或 body 尾端。
    """
    spans = []
    start = 0
    n = len(body_children)

    for i, child in enumerate(body_children):
        # 任何子孫含有 sectPr，就視為「本節的結尾」
        sect = child.xpath(".//w:sectPr", namespaces=NS)
        if sect:
            spans.append((start, i))
            start = i + 1

    # 極少數文件最後的 sectPr 可能在 <w:body> 尾端（非某段落內）
    # 若最後一節尚未收束，但 body 最後仍應該有 sectPr，可補上
    if start < n:
        # 嘗試檢查最後一個元素之外，body 結尾是否有 sectPr（保守處理：視為最後一節）
        spans.append((start, n - 1))

    return spans

def build_new_document_xml(template_xml_bytes, slice_elements):
    """
    用原始 document.xml 的外框（含命名空間與設定）建立新的 document.xml，
    但 <w:body> 只塞入本節的元素（保留最後一個元素中的 sectPr 當作此檔的分節設定）。
    """
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(template_xml_bytes, parser)

    body = root.find("w:body", NS)
    if body is None:
        raise RuntimeError("無法找到 <w:body>。這不是有效的 Word 文檔。")

    # 清空 body 內容
    for child in list(body):
        body.remove(child)

    # 填入本節的元素（深拷貝）
    for elem in slice_elements:
        body.append(elem)

    # 正規化分節設定
    normalize_last_section(body)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

def safe_title(s, maxlen=30):
    """將段首文字整理成適合檔名的短字串。"""
    s = s.strip()
    if not s:
        return "section"
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r'[\\/:*?"<>|]', "_", s)  # 不合法檔名字元
    return s[:maxlen].strip("_ ").strip()

def split_docx_by_sections(input_path, output_dir):
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"找不到檔案：{input_path}")
    os.makedirs(output_dir, exist_ok=True)

    # 讀入整個 .docx（ZIP）
    with open(input_path, "rb") as f:
        package_bytes = f.read()

    with zipfile.ZipFile(io.BytesIO(package_bytes)) as zin:
        # 取出 word/document.xml
        try:
            doc_xml = zin.read("word/document.xml")
        except KeyError:
            raise RuntimeError("此檔案不是有效的 .docx（缺少 word/document.xml）。")

        parser = etree.XMLParser(remove_blank_text=False)
        root = etree.fromstring(doc_xml, parser)
        body = root.find("w:body", NS)
        if body is None:
            raise RuntimeError("無法找到 <w:body>。")

        body_children = list(body)  # 直接取得 body 底下所有子元素（段落、表格等）
        if not body_children:
            raise RuntimeError("文件內容為空，無可分割的元素。")

        spans = find_section_spans(body_children)
        if not spans:
            # 沒找到分節，直接輸出一份完整副本
            base = os.path.splitext(os.path.basename(input_path))[0]
            out_path = os.path.join(output_dir, f"{base}_section_1.docx")
            with open(out_path, "wb") as fo:
                fo.write(package_bytes)
            print(f"未找到分節符號。已輸出完整文件：{out_path}")
            return
        
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        count = 0
        # 逐節切割並輸出到新的 .docx，檔案名稱來自於filename.txt檔案，
        # filename.txt檔案中的每一行依序對應到一個切割後的docx檔案的主檔名(附檔名仍為.docx)
        filename_map = {}
        with open("filename.txt", "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    # 將檔名儲存到對應的索引
                    filename_map[len(filename_map) + 1] = line

        for idx, (start, end) in enumerate(spans, 1):
            slice_elems = [etree.fromstring(etree.tostring(e)) for e in body_children[start:end+1]]
            
            # 建立新的 document.xml 內容
            new_doc_xml = build_new_document_xml(doc_xml, slice_elems)

            # 從filename_map取得對應的檔名，若無對應則使用safe_title產生
            out_filename = filename_map.get(idx, safe_title(text_of(slice_elems[0])) or f"section_{idx}")
            out_path = os.path.join(output_dir, f"{out_filename}.docx")

            # 以串流重建 zip（docx）
            with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED, allowZip64=True) as zout:
                for name in zin.namelist():
                    # 跳過目錄項（避免把目錄當檔案寫壞 ZIP）
                    if name.endswith('/'):
                        continue
                    data = zin.read(name)
                    if name == "word/document.xml":
                        data = new_doc_xml
                    # 用檔名字串寫入，避免沿用原 ZipInfo 的相容性問題
                    zout.writestr(name, data)

            count += 1
            print(f"輸出：{out_path}")

        print(f"完成，共輸出 {count} 節。")

def main():
    if len(sys.argv) != 3:
        print("用法：python split_docx_by_sections.py <input.docx> <output_dir>")
        sys.exit(1)
    input_path = sys.argv[1]
    output_dir = sys.argv[2]
    split_docx_by_sections(input_path, output_dir)

if __name__ == "__main__":
    main()
