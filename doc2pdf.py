import os
import sys
import time
import argparse
import subprocess
import tempfile
import shutil

import pythoncom
import pywintypes
import win32com.client.dynamic as win32_dynamic

RPC_E_CALL_REJECTED = -2147418111         # 0x80010001
RPC_E_SERVERCALL_RETRYLATER = -2147417846 # 0x8001010A

def pump_and_sleep(seconds=0.25):
    try:
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass
    time.sleep(seconds)

def retry_call(func, *args, max_retries=15, base_sleep=0.3, **kwargs):
    attempt = 0
    while True:
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as e:
            hr = getattr(e, "hresult", None)
            if hr is None and len(e.args) > 0:
                hr = e.args[0]

            if hr in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER):
                attempt += 1
                if attempt > max_retries:
                    raise
                wait_s = min(base_sleep * (1.6 ** (attempt - 1)), 3.0)
                pump_and_sleep(wait_s)
            else:
                raise

def unblock_file(path):
    try:
        subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "Unblock-File",
                "-Path", path
            ],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except Exception:
        pass

def convert_one_file(docx_path, pdf_path):
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        # 關鍵：使用 dynamic.Dispatch，避開 gen_py
        word = win32_dynamic.Dispatch("Word.Application")
        word.Visible = False

        try:
            word.DisplayAlerts = 0
        except Exception:
            pass

        doc = retry_call(
            word.Documents.Open,
            docx_path,
            False,   # ConfirmConversions
            True,    # ReadOnly
            False,   # AddToRecentFiles
        )

        wdExportFormatPDF = 17
        retry_call(
            doc.ExportAsFixedFormat,
            pdf_path,
            wdExportFormatPDF
        )

    finally:
        if doc is not None:
            try:
                retry_call(doc.Close, False)
            except Exception:
                pass

        if word is not None:
            try:
                retry_call(word.Quit)
            except Exception:
                pass

        pythoncom.CoUninitialize()
        pump_and_sleep(0.2)

def main():
    parser = argparse.ArgumentParser(description="批次將資料夾中的 DOCX 轉成 PDF（Windows dynamic 版）")
    parser.add_argument("folder", help="DOCX 資料夾路徑")
    parser.add_argument("-o", "--outdir", default=None, help="PDF 輸出資料夾")
    parser.add_argument("--unblock", action="store_true", help="先解除檔案封鎖")
    parser.add_argument("--copy-to-temp", action="store_true", help="先複製到暫存資料夾再轉")
    args = parser.parse_args()

    input_dir = os.path.abspath(args.folder)
    if not os.path.isdir(input_dir):
        print(f"❌ 找不到資料夾：{input_dir}")
        sys.exit(1)

    output_dir = os.path.abspath(args.outdir) if args.outdir else os.path.join(input_dir, "pdf_output")
    os.makedirs(output_dir, exist_ok=True)

    files = [f for f in os.listdir(input_dir) if f.lower().endswith(".docx") and not f.startswith("~$")]
    if not files:
        print("⚠️ 此資料夾中沒有 .docx 檔案")
        return

    print(f"開始轉換，共 {len(files)} 個檔案")

    for name in files:
        src = os.path.join(input_dir, name)
        dst = os.path.join(output_dir, os.path.splitext(name)[0] + ".pdf")

        work_src = src
        temp_dir = None

        try:
            if args.unblock:
                unblock_file(src)

            if args.copy_to_temp:
                temp_dir = tempfile.mkdtemp(prefix="docx2pdf_")
                work_src = os.path.join(temp_dir, name)
                shutil.copy2(src, work_src)
                if args.unblock:
                    unblock_file(work_src)

            print(f"轉換：{name}")
            convert_one_file(os.path.abspath(work_src), os.path.abspath(dst))
            print(f"✅ 完成：{os.path.basename(dst)}")

        except Exception as e:
            print(f"⚠️ 失敗：{name}：{e}")

        finally:
            if temp_dir and os.path.isdir(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)

    print(f"\n✅ 處理完成，PDF 輸出位置：{output_dir}")

if __name__ == "__main__":
    main()