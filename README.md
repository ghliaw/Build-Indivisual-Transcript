# 生成個人成績單，並上傳到Moodle
## 動機：
計算學期成績時，常常會因為公布整個成績表，導致有些學生會做比較而來與老師爭論，而且這樣做有洩漏個資的疑慮。
但是為每位學生做出專屬他自己的成績單，是一件很耗時的工作，因此想要發展出一套能自動產生個人成績單的機制，
透過程式來自動生成個人成績單，並可在Moodle上自動派送給每個學生。
## 原理：
1. 先用Excel整理所有學生的學期成績，然後用Microsoft Word中的郵件功能建立一個生成個人成績單的Template，導入Excel的內容，生成一個集成所有個人成績單的docx檔。
2. 接下來使用本專案的程式`split_docx_by_sections.py`將每個學生的個人成績單從集成檔中分離出來，形成個別的docx檔，然後再經過程式`doc2pdf.py`轉成pdf檔。
3. 在Moodle建立一個作業 \(用來派送個人成績單\)，下載此作業的計分試算表csv檔，依照檔案中每位學生的識別碼，分別建置其專屬資料夾，然後將其對應的個人成績單pdf檔複製到此資料夾。最後將所有資料夾壓縮成一個zip檔，上傳到Moodle作為這個作業的回饋檔案，Moodle就會自動派送給每個學生。
4. 學生只要打開這個作業，裡面的回饋檔就是他個人成績單pdf檔。
## 系統需求：
系統需要安裝以下軟體：
- Microsoft Word 201x版
- Python 3.1x版，要執行以下pip指令安裝pywin32與lxml套件：
  - `pip install pywin32 lxml`
## 使用流程：
流程圖如下：

<img width="500" alt="image" src="https://github.com/user-attachments/assets/7cbf1dbd-b995-4de7-8396-c6b569d39f4a" />

