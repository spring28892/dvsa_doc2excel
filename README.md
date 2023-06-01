# 專案介紹 — Description

本專案屬D4SG資料英雄333計畫成果，本專案目標係欲解決提案方(**臺北市家庭暴力暨性侵害防治中心**)提出之現有工作痛點如下：
1. 考量
2. On-Premise的工作環境
3. 盡量減低maintainance難度


# 運行環境需求 — Requirement

此專案主要運行在任何支援Content Control及FormFill版本的MS Office環境，根據微軟官方文件說法，Office 2007版及更之後釋出的版本應均有支援；但目前僅在Windows作業系統+Office 2019環境測試過。

基於使用到微軟Office單機版獨有的函式庫，故不支援OpenOffice、LibreOffice等它方軟體，亦不支援Office 365線上版(線下版可以)、Google Workspace等線上編輯環境。

# 系統架構圖 — System architecture



# 使用者操作手冊 — User Manual

Step1. 在GitHub下載本專案zip檔並解壓縮後，請直接執行run.bat批次檔。該批次檔案會自動建立Python虛擬環境及安裝相關套件。

Step2. 在input.docx(可自行更改檔名)中依據預先設計好的表格內容填寫資訊後，並完成存檔

Step3. 使用merge_template.docx，透過MS Word插入-文字檔(或是插入-物件-Microsoft Word Document)功能整併，將複數個案Word檔合併成一個word檔案

    注意：整併檔案時切勿不小心修改到任何資料，以及加入純文字以外的其他內容及格式，避免轉換至excel時出現問題

Step4. 打開excel檔，

Step5. 

# Demo 影片 — Demo Video

# 客製修改 - Customize 

## Dev mode



## Form Protection

input.docx檔預設開啟MS Word Form Protection功能，如要修改文件或關閉此功能，

(可自行決定要不要輸入密碼)



## Code修改

