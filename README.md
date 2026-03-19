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

### 1. 成績登錄到Excel檔
將學生各項作業/考試成績以及學期成績登錄到Excel檔，可以參考Example資料夾下的Score.xlsx檔。

<img width="500"  alt="image" src="https://github.com/user-attachments/assets/1667f748-cc54-4afa-9ccf-446efac8c583" />

### 2. 製作成績單word模板
- 複製一份Example裡面的「Template.docx」並以Word開啟。若出現以下視窗，請按「是」即可 (因為之前有連結到其他excel檔)。
  
  <img width="400" alt="image" src="https://github.com/user-attachments/assets/368c2b5f-649a-4670-9407-475f74fbb350" />
  
- 設定Template.docx與成績excel檔的連結：
  - 點選「郵件」-->「選取收件者」-->「使用現有清單…」

    <img width="400" alt="image" src="https://github.com/user-attachments/assets/008c2cbe-67cc-4687-a5ab-1454bb103a68" />
  
  - 在跳出的視窗中選擇成績excel檔，然後選擇登錄學期成績的工作表作為清單來源：
 
    <img width="400" height="577" alt="image" src="https://github.com/user-attachments/assets/bd8e7aef-4ec4-48ef-9646-945498e2ce86" />

- 修改Template.docx內容，使其顯示你想要的成績單內容：
  - 可以插入任一清單來源工作表的欄位來顯示：
    - 「郵件」-->「插入合併欄位」--> 選擇想要插入的欄位
   
      <img width="500"  alt="image" src="https://github.com/user-attachments/assets/1621aaaf-2e6b-48f6-9c9f-9bd32d7e5187" />
      
  - 範例成品如下圖：
 
    <img width="400" height="805" alt="image" src="https://github.com/user-attachments/assets/704818b4-5a66-472f-94d2-6ad3ddc7613b" />

### 3. 生成個人成績單集成docx檔
- (繼續在word中) 設定學生清單：
  - 點擊「郵件」-->「編輯收件者清單」--> 勾選要列印成績單的學生按下確定按鈕
 
    <img width="700"  alt="image" src="https://github.com/user-attachments/assets/d5ea0254-6dd4-4e84-9bcf-9fd8ccfcd880" />

- 預覽個人成績單內容：
  - 點擊「郵件」-->「預覽結果」，便可預覽每位同學的個人成績單內容
    
    <img width="400"  alt="image" src="https://github.com/user-attachments/assets/00a11a8b-330a-459c-90a8-88a1973776ed" />

- 產生個人成績單集成檔：
  - 點擊「郵件」-->「完成與合併」--> 編輯個別文件…
  - 會跳出「合併到新文件」視窗，點選「全部」再按下確定按鈕
    
    <img width="700" alt="image" src="https://github.com/user-attachments/assets/0426b705-9acc-47b6-9f4c-73fe614444fe" />

  - 便會產生一個新的Word檔，集成所有個人成績單，每人一頁
  - 將此集成檔存檔成**`all.docx`**備用
 
    <img width="500" alt="image" src="https://github.com/user-attachments/assets/fe67c2e2-8b60-4b6f-ba36-ec2335385ec6" />

### 4. 到Moodle建立一個放置成績單的作業
- 新增一個「作業」資源：

  <img width="500"  alt="image" src="https://github.com/user-attachments/assets/18d22e4b-f33a-4964-97f6-5a0e5f30fcd8" />

  <img width="600" height="1043" alt="image" src="https://github.com/user-attachments/assets/76b6e642-9c08-440e-bc1c-7b1553c678a8" />

- 修改作業設定：
  - (本例作業名稱取名為「個人學期成績單」)
  - 可用性：都不要勾選(即不設定繳交期限)
  - 回饋類型：**只勾選「離線計分試算表」與「回饋檔案」**
 
    <img width="500" height="632" alt="image" src="https://github.com/user-attachments/assets/b1ff41a4-8395-455e-8fd3-ad51bd48f9ad" />

    <img width="500" height="1039" alt="image" src="https://github.com/user-attachments/assets/f85ea8d6-b7a7-4abb-91ed-657b45017731" />

### 5. 下載計分用試算表(csv檔)
- 進入「個人成績單」作業，點擊「檢視所有繳交作業」按鈕，在「計分動作」的下拉選單中，點選「下載計分用試算表」，便會下載一個csv檔：

  <img width="500" alt="image" src="https://github.com/user-attachments/assets/3887dc92-a538-44d6-8a2c-7a7d06f3708c" />

  <img width="500" height="952" alt="image" src="https://github.com/user-attachments/assets/a3bf6473-28f2-4522-96a2-5bd7aa33e66d" />

  <img width="500" height="1003" alt="image" src="https://github.com/user-attachments/assets/77c45fe7-76f2-4f70-abbb-ce77d121b7b0" />

### 6. 執行build_transcript.py獲得上傳用zip壓縮檔

<img width="800"  alt="image" src="https://github.com/user-attachments/assets/008c8245-ff54-44cd-bd0f-c65d26cce163" />

<img width="400" alt="image" src="https://github.com/user-attachments/assets/a5c667dd-eb7f-4995-8fe3-742d66de9acf" />


### 7. 上傳zip壓縮檔到放置成績單的作業中

<img width="500" height="968" alt="image" src="https://github.com/user-attachments/assets/fd81d84f-4d4d-47db-87ce-867a84c3583e" />

<img width="500" height="1007" alt="image" src="https://github.com/user-attachments/assets/28778bb1-8091-4650-ad0e-f1a6c04e842d" />

<img width="400" height="1102" alt="image" src="https://github.com/user-attachments/assets/69ed68df-103f-4ca0-9845-09f16d8b4330" />

<img width="600" height="679" alt="image" src="https://github.com/user-attachments/assets/03e6f124-5147-4167-b261-a15b75ca44c2" />

<img width="1647" height="400" alt="image" src="https://github.com/user-attachments/assets/405ef03b-ef3f-4abd-8006-1d1e4a05dcc9" />













  









    








    











    













