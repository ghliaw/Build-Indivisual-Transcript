import sys
import csv
import os
import shutil
import zipfile

def prepare_filename_list_and_folders(csvfilename):
    with open(csvfilename, newline='', encoding='utf-8') as csvfile:
        # 刪除csvfile的標題列
        next(csvfile)
        rows = csv.reader(csvfile)
        # 建立一個文字檔，檔名為Filename.txt，內容為csv檔案中每一列的第三欄
        with open('filename.txt', 'w', encoding='utf-8') as f:
            for row in rows:
                print(row)
                # 取出第一欄，將此攔第一個阿拉伯數字前的非數字的文字刪除掉
                participant_id = row[0]
                participant_id = ''.join(filter(str.isdigit, participant_id))
                print(participant_id)
                dir_path = 'Participant_'+participant_id+'_'+'assignsubmission_file_'
                if not os.path.exists(dir_path):
                    os.makedirs(dir_path)
                f.write(row[2] + '\n')

'''
這個程式依序做以下的動作：
1. 讀取一個來自Moodle的批次計分用的試算表(csv格式)，將每一列第三欄的內容寫入Filename.txt檔案中(每一行一個檔名)
2. 讀取一個裡面用分節符號分為很多節的docx檔，將每一節獨立成為一個docx檔案，並根據Filename.txt中的檔名來命名這些切割後的docx檔案，然後存放到output_sections資料夾中
3. 將output_sections資料夾中的docx檔轉成pdf檔，然後存放到output_pdfs資料夾中
4. 將output_pdfs資料夾中每個pdf檔案依序移動到對應的資料夾中，資料夾名稱為Participant_參與者編號_assignsubmission_file_(參與者編號取自csv檔案的第一欄，只保留阿拉伯數字部分)
5. 最後將這些資料夾打包成一個zip檔案，方便上傳到Moodle
主程式命令列參數：python build_transscript.py <csvfilename> <input_docx> <output_zip>
'''
def main():
    if len(sys.argv) != 4:
        print("用法：python build_transscript.py <csvfilename> <input_docx> <output_zip>")
        sys.exit(1)
    csvfilename = sys.argv[1]
    input_docx = sys.argv[2]
    output_zip = sys.argv[3]
    prepare_filename_list_and_folders(csvfilename)
    # 呼叫split_docx_by_sections.py來切割docx檔案
    os.system(f'python split_docx_by_sections.py "{input_docx}" "output_sections"')
    # 呼叫doc2pdf.py來將切割後的docx檔案轉成pdf檔案
    os.system(f'python doc2pdf.py "output_sections" --outdir "output_pdfs" --no-isolate --unblock')
    # 讀取csv檔案，將output_pdfs資料夾中的pdf檔案依序移動到對應的資料夾中
    with open(csvfilename, newline='', encoding='utf-8') as csvfile:
        next(csvfile)
        rows = csv.reader(csvfile)
        for row in rows:
            participant_id = row[0]
            participant_id = ''.join(filter(str.isdigit, participant_id))
            dir_path = 'Participant_'+participant_id+'_'+'assignsubmission_file_'
            pdf_filename = row[2] + '.pdf'
            src_pdf_path = os.path.join('output_pdfs', pdf_filename)
            dst_pdf_path = os.path.join(dir_path, pdf_filename)
            if os.path.exists(src_pdf_path):
                shutil.move(src_pdf_path, dst_pdf_path)
            else:
                print(f"警告：找不到檔案 {src_pdf_path}，無法移動到 {dst_pdf_path}")
    # 將這些資料夾打包成一個zip檔案
    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for foldername in os.listdir('.'):
            if foldername.startswith('Participant_') and os.path.isdir(foldername):
                for root, dirs, files in os.walk(foldername):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, '.')
                        zipf.write(file_path, arcname)
    print(f"完成！已建立壓縮檔案：{output_zip}")

if __name__ == "__main__":
    main() 
