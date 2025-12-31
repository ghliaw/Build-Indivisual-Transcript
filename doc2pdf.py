import os
import sys
import time
import argparse
import subprocess

import pythoncom
import win32com.client as win32
import pywintypes

RPC_E_CALL_REJECTED = -2147418111        # 0x80010001
RPC_E_SERVERCALL_RETRYLATER = -2147417846 # 0x8001010A

def pump_and_sleep(seconds=0.3):
    try:
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass
    time.sleep(seconds)

def retry_call(func, *args, max_retries=12, base_sleep=0.25, **kwargs):
    attempt = 0
    while True:
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as e:
            hr = getattr(e, "hresult", e.args[0] if e.args else None)
            if hr in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER):
                attempt += 1
                if attempt > max_retries:
                    raise
                sleep_s = min(base_sleep * (1.7 ** (attempt - 1)), 3.0)
                pump_and_sleep(sleep_s)
            else:
                raise

def unblock_file(path):
    # 移除 Zone.Identifier，避免保護檢視導致 Open 卡住或被拒絕
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "Unblock-File", "-Path", path],
            check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        # 也嘗試直接刪除附加資料流
        try:
            os.remove(path + ":Zone.Identifier")
        except Exception:
            pass
    except Exception:
        pass

def convert_one_docx_with_new_word(docx_path, pdf_path, verbose=True):
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32.DispatchEx("Word.Application")  # 關鍵：新的 COM 執行個體
        word.Visible = False
        try:
            word.DisplayAlerts = 0
            opts = word.Options
            for attr, val in [("PrintBackground", False), ("BackgroundSave", False), ("SaveInterval", 0)]:
                try: setattr(opts, attr, val)
                except Exception: pass
        except Exception:
            pass

        wdExportFormatPDF = 17

        doc = retry_call(
            word.Documents.Open,
            docx_path,
            ReadOnly=True,
            ConfirmConversions=False,
            Visible=False,
            AddToRecentFiles=False
        )

        retry_call(
            doc.ExportAsFixedFormat,
            OutputFileName=pdf_path,
            ExportFormat=wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=0,      # 0=Print
            Range=0,            # 0=All
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=0,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )

        retry_call(doc.Close, False)
        if verbose:
            print(f"✅ {os.path.basename(docx_path)} → {os.path.basename(pdf_path)}")
    finally:
        if word is not None:
            try:
                retry_call(word.Quit)
            except Exception:
                pass
        pythoncom.CoUninitialize()
        pump_and_sleep(0.15)  # 給系統一點時間釋放 COM/檔案鎖

def main():
    ap = argparse.ArgumentParser(description="將資料夾內所有 .docx 轉為同名 .pdf（逐檔隔離 Word 實例，解 RPC 拒絕）")
    ap.add_argument("folder", help="含有 .docx 檔案的資料夾路徑")
    ap.add_argument("-o", "--outdir", default=None, help="PDF 輸出資料夾（預設 <資料夾>/pdf_output）")
    ap.add_argument("--no-isolate", action="store_true", help="禁用逐檔隔離（僅除錯用，不建議）")
    ap.add_argument("--unblock", action="store_true", help="轉檔前先解除檔案的保護檢視封鎖 (Unblock-File)")
    args = ap.parse_args()

    in_dir = os.path.abspath(args.folder)
    if not os.path.isdir(in_dir):
        print(f"❌ 找不到資料夾：{in_dir}")
        sys.exit(1)

    out_dir = os.path.abspath(args.outdir) if args.outdir else os.path.join(in_dir, "pdf_output")
    os.makedirs(out_dir, exist_ok=True)

    docx_list = [f for f in os.listdir(in_dir) if f.lower().endswith(".docx") and not f.startswith("~$")]
    if not docx_list:
        print("⚠️ 此資料夾沒有 .docx 檔。")
        return

    print(f"開始轉換，共 {len(docx_list)} 個檔案…（隔離模式：{'ON' if not args.no_isolate else 'OFF'}）")
    if args.unblock:
        print("🔓 解除保護檢視封鎖：已啟用")

    if args.no_isolate:
        # 單一 Word 實例（若你只想測試）
        pythoncom.CoInitialize()
        word = None
        try:
            word = win32.gencache.EnsureDispatch("Word.Application")
            word.Visible = False
            try:
                word.DisplayAlerts = 0
            except Exception:
                pass
            wdExportFormatPDF = 17
            for name in docx_list:
                src = os.path.join(in_dir, name)
                dst = os.path.join(out_dir, os.path.splitext(name)[0] + ".pdf")
                if args.unblock:
                    unblock_file(src)
                try:
                    doc = retry_call(word.Documents.Open, src, ReadOnly=True, ConfirmConversions=False, Visible=False, AddToRecentFiles=False)
                    retry_call(doc.ExportAsFixedFormat,
                               OutputFileName=dst, ExportFormat=wdExportFormatPDF, OpenAfterExport=False,
                               OptimizeFor=0, Range=0, Item=0, IncludeDocProps=True, KeepIRM=True,
                               CreateBookmarks=0, DocStructureTags=True, BitmapMissingFonts=True, UseISO19005_1=False)
                    retry_call(doc.Close, False)
                    print(f"✅ {name} → {os.path.basename(dst)}")
                except Exception as e:
                    print(f"⚠️ 失敗：{name}：{e}")
                pump_and_sleep(0.15)
        finally:
            if word is not None:
                try: retry_call(word.Quit)
                except Exception: pass
            pythoncom.CoUninitialize()
    else:
        # 預設：逐檔隔離實例（最穩）
        for name in docx_list:
            src = os.path.join(in_dir, name)
            dst = os.path.join(out_dir, os.path.splitext(name)[0] + ".pdf")
            if args.unblock:
                unblock_file(src)
            try:
                convert_one_docx_with_new_word(src, dst)
            except Exception as e:
                print(f"⚠️ 失敗：{name}：{e}")

    print(f"\n✅ 完成！PDF 位於：{out_dir}")

if __name__ == "__main__":
    main()
