# 專案介紹 — Description

本專案屬D4SG資料英雄333計畫「**兒少保護高效能網絡合作**」成果之一，本專案目標係欲優化提案方(**臺北市家庭暴力暨性侵害防治中心**)現行跨網絡會議事前資料準備及彙整之效率，經盤點提案方端需求如下：
1. 考量保護對象有大量機敏資訊，希望能使用純地端(On-Premises)的工作環境
2. 考量後續社工接手時架構維護難度，暫不考慮跨系統介接
3. 盡可能在任何環境都能使用(即不需特別設置工作環境)
4. 填寫方式盡量簡化，如果可以不用排版最好
5. 希望資料後續以方便延伸做加值應用

基於以上前題，最終決定以MS Office Word作為資料輸入介面，並利用FormFill跟Content Control優化輸入方式以及限制輸入類型(Type)，再透過MS Office Excel 內建之 Visual Basic Advance 撰寫巨集(Macro)讓使用者點選檔案匯入。

# 運行環境需求 — Requirement

此專案主要運行在任何支援Content Control及FormFill版本的MS Office環境，根據微軟官方文件說法，Office 2007版及更之後釋出的版本應均有支援；但目前僅在Windows作業系統+Office 2016/2019環境測試過。

基於使用到微軟Office單機版獨有的函式庫，故不支援OpenOffice、LibreOffice等它方軟體，亦不支援Office 365線上版(線下版可以)、Google Workspace等線上編輯環境。

# 系統架構圖 — System architecture

![structure](https://github.com/spring28892/dvsa_doc2excel/blob/main/structure.png?raw=true)

# 使用者操作手冊 — User Manual

Step1. 在GitHub下載本專案zip檔並解壓縮。

Step2. 在input.docx(可自行更改檔名)中依據預先設計好的表格內容填寫資訊後，並完成存檔。

Step3. 使用merge_template.docx(或是用空白的Word文件，版面配置：邊界窄、方向橫向)，透過MS Word插入-文字檔(或是插入-物件-Microsoft Word Document)功能整併，將複數個案Word檔合併成一個word檔案。

    注意：整併檔案時切勿不小心修改到任何資料，以及加入純文字以外的其他內容及格式，避免轉換至Excel時出現問題。

Step4. 打開Excel檔，使用左上巨集後，選取合併後的Word檔案，資料便會自動匯入；注意執行前需要啟用巨集，如果出現不信任檔案而被封鎖狀況時，請自行[解除檔案封鎖](https://www.pcmarket.com.hk/microsoft-office-will-block-vba-marcos-by-default/)。

Step5. 如有重複的資料，會在Excel最左邊欄位以顏色標註提醒，可自行刪除。

    可使用本repo內的input_demo1.docx作為測試用檔案

# Demo 影片 — Demo Video



# 客製修改 - Customize 

## Dev mode

如要修改Word檔案中的Content Control及FormFill，需先開啟「開發人員」索引標籤，請參考[微軟官方說明](https://learn.microsoft.com/zh-tw/visualstudio/vsto/how-to-show-the-developer-tab-on-the-ribbon?view=vs-2022)。

## Form Protection

input.docx檔預設開啟MS Word Form Protection功能，如要修改文件或關閉此功能，可於「開發人員」索引標籤下「限制編輯」選項中修改(可自行決定要不要輸入密碼)

詳細說明可參考[微軟官方說明](https://support.microsoft.com/zh-hk/office/%E5%BB%BA%E7%AB%8B%E4%BD%BF%E7%94%A8%E8%80%85%E5%8F%AF%E5%9C%A8-word-%E4%B8%AD%E5%AE%8C%E6%88%90%E6%88%96%E5%88%97%E5%8D%B0%E7%9A%84%E8%A1%A8%E5%96%AE-040c5cc1-e309-445b-94ac-542f732c8c8b)。

    注意：為了方便維護，修改完成後切記一定要開啟限制標及，避免填寫者不小心更動刪除到表單內容或格式。

## Code修改

在Excel中按Alt+F11可以進入VBA編輯畫面，選擇左邊的模組->Module1即可修改程式碼。
    注意：請僅在在程式碼註解標註勿動以外的地方修改

VBA程式碼範例：

    sh.Cells(lr, 2).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("meeting_time")(1).Range.Text)

這段程式碼中，`(doc.SelectContentControlsByTitle("meeting_time")(1).Range.Text)`代表從目標Word檔案中抓取屬性標題是`"meeting_time"`的欄位的文字值(`.Range.Text`)，透過`Application.WorksheetFunction.Clean`函式把一些可能造成資料錯誤的文字特殊字元清乾淨後。透過`sh.Cells(lr, 2).Value`，代表在第lr行第2欄的儲存格(`sh.Cells`)中寫入我們抓取出來的值。

    另外lr行代表 Last Row = 最後一行，這值通常不動 (因為我們抓進來的資料基本上都會一直新增在最後一行)

也就是說如果我們在Word檔新增一個ContentControl屬性物件，並要抓一個新的欄位的文字值的話，可以直接複製這段程式碼，然後把本來`"meeting_time"`跟`sh.Cells(lr, 2)`這兩段分別置換成要抓的對象跟要存值的excel儲存格位置。
 
***



    If doc.SelectContentControlsByTitle("conviction_t")(1).Checked = True Then
        sh.Cells(lr, 72).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("conviction_text")(1).Range.Text)
    Else
        sh.Cells(lr, 72).Value = "0"
    End If

另外物件的`.Checked`屬性為Boolean值，True代表那項有被勾選，因此透過Checkbox(勾選項)加上If/ElseIf/Else等邏輯判斷，也可以做到比如`"conviction_t"`(前科紀錄)Checkbox選項有被勾選的話，就回傳寫入`"conviction_text"`(前科紀錄)，如沒有被勾選，則回傳`0`。


# 已知Issues

1. ContentControl無法抓到list屬性物件的Value只能抓到Title.
2. 無法事先判斷物件的值是使用者填入的值還是是先輸入的提示文字(ContentControl.Placeholder)