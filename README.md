# 專案介紹 — Description

本專案屬D4SG資料英雄333計畫「**兒少保護高效能網絡合作**」成果之一，本專案目標係欲解決提案方(**臺北市家庭暴力暨性侵害防治中心**)提出之，經盤點提案方端需求如下：
1. 考量保護對象有大量機敏資訊，希望能使用純地端(On-Premises)的工作環境
2. 盡量減低後續社工接手時架構維護難度，暫不考慮跨系統介接
3. 
4. 填寫方式盡量簡化，如果可以不用排版最好
5. 希望資料後續以方便延伸做加值應用

基於以上前題，最終決定以MS Office Word作為資料輸入介面，並利用FormFill跟Content Control優化輸入方式以及限制輸入類型(Type)，再透過MS Office Excel 內建之 Visual Basic Advance 撰寫巨集(Macro)讓使用者點選檔案匯入。

# 運行環境需求 — Requirement

此專案主要運行在任何支援Content Control及FormFill版本的MS Office環境，根據微軟官方文件說法，Office 2007版及更之後釋出的版本應均有支援；但目前僅在Windows作業系統+Office 2019環境測試過。

基於使用到微軟Office單機版獨有的函式庫，故不支援OpenOffice、LibreOffice等它方軟體，亦不支援Office 365線上版(線下版可以)、Google Workspace等線上編輯環境。

# 系統架構圖 — System architecture

![structure]([https://github.com/spring28892/dvsa_doc2excel/blob/main/structure.png?raw=true])

# 使用者操作手冊 — User Manual

Step1. 在GitHub下載本專案zip檔並解壓縮。

Step2. 在input.docx(可自行更改檔名)中依據預先設計好的表格內容填寫資訊後，並完成存檔。

Step3. 使用merge_template.docx(或是用空白的Word文件，版面配置：邊界窄、方向橫向)，透過MS Word插入-文字檔(或是插入-物件-Microsoft Word Document)功能整併，將複數個案Word檔合併成一個word檔案。

    注意：整併檔案時切勿不小心修改到任何資料，以及加入純文字以外的其他內容及格式，避免轉換至Excel時出現問題。

Step4. 打開Excel檔，使用左上巨集後，選取合併後的Word檔案，資料便會自動匯入。

Step5. 如有重複的資料，會在Excel最左邊欄位以顏色標註提醒，可自行刪除。

# Demo 影片 — Demo Video

# 客製修改 - Customize 

## Dev mode



## Form Protection

input.docx檔預設開啟MS Word Form Protection功能，如要修改文件或關閉此功能，

(可自行決定要不要輸入密碼)



## Code修改

在Excel
