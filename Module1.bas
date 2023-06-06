Attribute VB_Name = "Module1"
Option Explicit
Sub extractData()

    Dim wd As New Word.Application
    Dim doc As Word.Document
    Dim FileToOpen As Variant
    Dim tbls As Object
    Dim sh As Worksheet
    Dim ff As Object    ' Word.FormField
    Dim cbValue As String, celString As String
    Dim rngFF As Object ' Word.Range
    Dim lr As Integer, i As Integer
    Dim parenting_value As String
    Dim cel As Range
    
    wd.Visible = True

    FileToOpen = Application.GetOpenFilename(Title:="選取要開啟的檔案", FileFilter:="Word Files (*.docx*), *docx*") ' 選擇開啟的檔案
    If FileToOpen <> False Then '如果使用者沒取消的話把doc設定成開啟的檔案
        Set doc = wd.Documents.Open(FileToOpen)
    Else
        Exit Sub
    End If
    
    Set tbls = doc.Tables '把doc的table都抓出來
    Set sh = ActiveSheet  '指定sh為現在開啟的excel分頁

    lr = sh.UsedRange.Rows.Count + 1 '找出有資料的最後一行+1當成第一個空行
    
    '此行以上勿動!
    
    For i = 1 To doc.SelectContentControlsByTitle("meeting_time").Count
        sh.Cells(lr, 2).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("meeting_time")(1).Range.Text)
        sh.Cells(lr, 3).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("case_id")(1).Range.Text)
        sh.Cells(lr, 4).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("project_time")(1).Range.Text)
        sh.Cells(lr, 5).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("district")(1).Range.Text)
        sh.Cells(lr, 6).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("social_worker")(1).Range.Text)
        sh.Cells(lr, 7).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("child_name")(1).Range.Text)
        sh.Cells(lr, 8).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("child_id")(1).Range.Text)
        sh.Cells(lr, 9).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("child_age")(1).Range.Text)
        sh.Cells(lr, 10).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("gender")(1).Range.Text)
        sh.Cells(lr, 11).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("disable")(1).Range.Text)
        sh.Cells(lr, 12).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("abuser_name")(1).Range.Text)
        sh.Cells(lr, 13).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("abuser_title")(1).Range.Text)
        sh.Cells(lr, 14).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("abuser_id")(1).Range.Text)
        sh.Cells(lr, 15).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("abuser_age")(1).Range.Text)
        sh.Cells(lr, 16).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("abuse_same")(1).Range.Text)
        sh.Cells(lr, 17).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("abuse_type")(1).Range.Text)
        sh.Cells(lr, 18).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("caregiver_title")(1).Range.Text)
        sh.Cells(lr, 19).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("caregiver_age")(1).Range.Text)
        sh.Cells(lr, 20).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("other_childs_count")(1).Range.Text)
        sh.Cells(lr, 21).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("risk_grade")(1).Range.Text)
        sh.Cells(lr, 22).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("score_comp")(1).Range.Text)
        
        If doc.SelectContentControlsByTitle("new_case")(1).Checked = True Then
            sh.Cells(lr, 23).Value = "1"
        ElseIf doc.SelectContentControlsByTitle("old_case")(1).Checked = True Then
            sh.Cells(lr, 23).Value = "0"
        End If
            
        If doc.SelectContentControlsByTitle("visit_1")(1).Checked = True And doc.SelectContentControlsByTitle("visit_2")(1).Checked Then
            sh.Cells(lr, 24).Value = "c"
        ElseIf doc.SelectContentControlsByTitle("visit_1")(1).Checked = True Then
            sh.Cells(lr, 24).Value = "a"
        ElseIf doc.SelectContentControlsByTitle("visit_2")(1).Checked = True Then
            sh.Cells(lr, 24).Value = "b"
        Else
            sh.Cells(lr, 24).Value = "d"
        End If
        
        If doc.SelectContentControlsByTitle("danger")(1).Checked = True Then
            sh.Cells(lr, 25).Value = "1"
        Else
            sh.Cells(lr, 25).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("sdmr")(1).Checked = True Then
            sh.Cells(lr, 26).Value = "1"
        Else
            sh.Cells(lr, 26).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("9factor_1")(1).Checked = True Then
            sh.Cells(lr, 27).Value = "1"
        Else
            sh.Cells(lr, 27).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("9factor_2")(1).Checked = True Then
            sh.Cells(lr, 28).Value = "1"
        Else
            sh.Cells(lr, 28).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("9factor_3")(1).Checked = True Then
            sh.Cells(lr, 29).Value = "1"
        Else
            sh.Cells(lr, 29).Value = "0"
        End If
    
        If doc.SelectContentControlsByTitle("9factor_4")(1).Checked = True Then
            sh.Cells(lr, 30).Value = "1"
        Else
            sh.Cells(lr, 30).Value = "0"
        End If
    
        If doc.SelectContentControlsByTitle("9factor_5")(1).Checked = True Then
            sh.Cells(lr, 31).Value = "1"
        Else
            sh.Cells(lr, 31).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("9factor_6")(1).Checked = True Then
            sh.Cells(lr, 32).Value = "1"
        Else
            sh.Cells(lr, 32).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("9factor_7")(1).Checked = True Then
            sh.Cells(lr, 33).Value = "1"
        Else
            sh.Cells(lr, 33).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("9factor_8")(1).Checked = True Then
            sh.Cells(lr, 34).Value = "1"
        Else
            sh.Cells(lr, 34).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("9factor_9")(1).Checked = True Then
            sh.Cells(lr, 35).Value = "1"
        Else
            sh.Cells(lr, 35).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("society_focus")(1).Checked = True Then
            sh.Cells(lr, 36).Value = "1"
        Else
            sh.Cells(lr, 36).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("heavy_abuse")(1).Checked = True Then
            sh.Cells(lr, 37).Value = "1"
        Else
            sh.Cells(lr, 37).Value = "0"
        End If
        
        sh.Cells(lr, 38).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("prev_listed")(1).Range.Text)
        sh.Cells(lr, 39).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("last_inform_date")(1).Range.Text)
        sh.Cells(lr, 40).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("last_inform_status")(1).Range.Text)
        sh.Cells(lr, 41).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("case_status")(1).Range.Text)
        
        If doc.SelectContentControlsByTitle("parenting_edu_f")(1).Checked = True Then
            sh.Cells(lr, 42).Value = "D"
        Else
            parenting_value = ""
            If doc.SelectContentControlsByTitle("parenting_edu_1")(1).Checked = True Then
                parenting_value = "A"
                sh.Cells(lr, 43).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("notify_hour")(1).Range.Text)
            End If
            If doc.SelectContentControlsByTitle("parenting_edu_2")(1).Checked = True Then
                parenting_value = "B"
                sh.Cells(lr, 44).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("force_hour")(1).Range.Text)
            End If
            If doc.SelectContentControlsByTitle("parenting_edu_3")(1).Checked = True Then
                parenting_value = "C"
            End If
            sh.Cells(lr, 42).Value = parenting_value
        End If
        
        sh.Cells(lr, 45).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("child_health_status")(1).Range.Text)
        sh.Cells(lr, 46).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("abuser_health_status")(1).Range.Text)
        sh.Cells(lr, 47).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("prot_order_eva")(1).Range.Text)
        sh.Cells(lr, 48).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("child_inform_counts")(1).Range.Text)
        sh.Cells(lr, 49).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("child_inform_record")(1).Range.Text)
        sh.Cells(lr, 50).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("family_evaluate")(1).Range.Text)
        sh.Cells(lr, 51).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("difficult_1")(1).Range.Text)
        sh.Cells(lr, 52).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("network_coop_1")(1).Range.Text)
        sh.Cells(lr, 53).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("network_coop_1_unit")(1).Range.Text)
        sh.Cells(lr, 54).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("difficult_2")(1).Range.Text)
        sh.Cells(lr, 55).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("network_coop_2")(1).Range.Text)
        sh.Cells(lr, 56).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("network_coop_2_unit")(1).Range.Text)
        sh.Cells(lr, 57).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("difficult_3")(1).Range.Text)
        sh.Cells(lr, 58).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("network_coop_3")(1).Range.Text)
        sh.Cells(lr, 59).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("network_coop_3_unit")(1).Range.Text)
        sh.Cells(lr, 60).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("difficult_4")(1).Range.Text)
        sh.Cells(lr, 61).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("network_coop_4")(1).Range.Text)
        sh.Cells(lr, 62).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("network_coop_4_unit")(1).Range.Text)
        
        If doc.SelectContentControlsByTitle("declass_t")(1).Checked = True Then
            sh.Cells(lr, 63).Value = "1"
        Else
            sh.Cells(lr, 63).Value = "0"
        End If
        
        sh.Cells(lr, 64).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("declass_reason")(1).Range.Text)
        sh.Cells(lr, 65).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("police_name")(1).Range.Text)
        
        If doc.SelectContentControlsByTitle("protection_order_1")(1).Checked = True Then
            sh.Cells(lr, 66).Value = "A"
        ElseIf doc.SelectContentControlsByTitle("protection_order_2")(1).Checked = True Then
            sh.Cells(lr, 66).Value = "B"
            sh.Cells(lr, 67).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("prot_apply_date")(1).Range.Text)
        ElseIf doc.SelectContentControlsByTitle("protection_order_3")(1).Checked = True Then
            sh.Cells(lr, 66).Value = "C"
        ElseIf doc.SelectContentControlsByTitle("protection_order_4")(1).Checked = True Then
            sh.Cells(lr, 66).Value = "D"
        ElseIf doc.SelectContentControlsByTitle("protection_order_5")(1).Checked = True Then
            sh.Cells(lr, 66).Value = "E"
        ElseIf doc.SelectContentControlsByTitle("protection_order_7")(1).Checked = True Then
            sh.Cells(lr, 66).Value = "F"
        ElseIf doc.SelectContentControlsByTitle("protection_order_8")(1).Checked = True Then
            sh.Cells(lr, 66).Value = "G"
        ElseIf doc.SelectContentControlsByTitle("protection_order_9")(1).Checked = True Then
            sh.Cells(lr, 66).Value = "H"
        End If
        
        sh.Cells(lr, 68).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("protection_no")(1).Range.Text)
        sh.Cells(lr, 69).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("protection_date_1")(1).Range.Text) & "至" & Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("protection_date_2")(1).Range.Text)
        
        If doc.SelectContentControlsByTitle("protection_status_t")(1).Checked = True Then
            sh.Cells(lr, 70).Value = "1"
        Else
            sh.Cells(lr, 70).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("violate_prot_t")(1).Checked = True Then
            sh.Cells(lr, 71).Value = "1"
        Else
            sh.Cells(lr, 71).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("conviction_t")(1).Checked = True Then
            sh.Cells(lr, 72).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("conviction_text")(1).Range.Text)
        Else
            sh.Cells(lr, 72).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("crime_t")(1).Checked = True Then
            sh.Cells(lr, 73).Value = "1"
            sh.Cells(lr, 74).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("inspection_freq")(1).Range.Text)
            sh.Cells(lr, 75).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("inspection_item")(1).Range.Text)
        Else
            sh.Cells(lr, 73).Value = "0"
        End If
           
        sh.Cells(lr, 76).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("police_netcoop")(1).Range.Text)
        sh.Cells(lr, 77).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("police_netcoop_unit")(1).Range.Text)
        
        sh.Cells(lr, 78).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("health_name")(1).Range.Text)
        
        If doc.SelectContentControlsByTitle("abuser_mental_t")(1).Checked = True Then
            sh.Cells(lr, 79).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("health_service")(1).Range.Text)
            sh.Cells(lr, 80).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("care_level")(1).Range.Text)
        Else
            sh.Cells(lr, 79).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("abuser_suicide_t")(1).Checked = True Then
            sh.Cells(lr, 81).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("abuser_suicide_text")(1).Range.Text)
        Else
            sh.Cells(lr, 81).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("drug_t")(1).Checked = True Then
            sh.Cells(lr, 82).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("drug_text")(1).Range.Text)
        Else
            sh.Cells(lr, 82).Value = "0"
        End If
           
        sh.Cells(lr, 83).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("health_netcoop")(1).Range.Text)
        sh.Cells(lr, 84).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("health_netcoop_unit")(1).Range.Text)
        
        sh.Cells(lr, 85).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("education_name")(1).Range.Text)
        sh.Cells(lr, 86).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("edu_status")(1).Range.Text)
    
        If doc.SelectContentControlsByTitle("edu_consult_t")(1).Checked = True Then
            sh.Cells(lr, 87).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("edu_consult_text")(1).Range.Text)
        Else
            sh.Cells(lr, 87).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("edu_resource_t")(1).Checked = True Then
            sh.Cells(lr, 88).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("edu_resource_text")(1).Range.Text)
        Else
            sh.Cells(lr, 88).Value = "0"
        End If
        
        If doc.SelectContentControlsByTitle("edu_concern_t")(1).Checked = True Then
            sh.Cells(lr, 89).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("edu_concern_text")(1).Range.Text)
        Else
            sh.Cells(lr, 89).Value = "0"
        End If
    
        sh.Cells(lr, 90).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("edu_violence")(1).Range.Text)
        sh.Cells(lr, 91).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("edu_school_safe")(1).Range.Text)
        sh.Cells(lr, 92).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("edu_home_safe")(1).Range.Text)
        sh.Cells(lr, 93).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("edu_netcoop")(1).Range.Text)
        sh.Cells(lr, 94).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("edu_netcoop_unit")(1).Range.Text)
        
        sh.Cells(lr, 95).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others1_name")(1).Range.Text)
        
        If doc.SelectContentControlsByTitle("others1_status3")(1).Checked = True Then
            sh.Cells(lr, 96).Value = "C"
        ElseIf doc.SelectContentControlsByTitle("others1_status2")(1).Checked = True Then
            sh.Cells(lr, 96).Value = "B"
        ElseIf doc.SelectContentControlsByTitle("others1_status1")(1).Checked = True Then
            sh.Cells(lr, 96).Value = "A"
        End If
        
        sh.Cells(lr, 97).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others1_safe")(1).Range.Text)
        sh.Cells(lr, 98).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others1_violence")(1).Range.Text)
        sh.Cells(lr, 99).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others1_netcoop")(1).Range.Text)
        sh.Cells(lr, 100).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others1_netcoop_unit")(1).Range.Text)
        
        sh.Cells(lr, 101).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others2_name")(1).Range.Text)
        
        If doc.SelectContentControlsByTitle("others2_status3")(1).Checked = True Then
            sh.Cells(lr, 102).Value = "C"
        ElseIf doc.SelectContentControlsByTitle("others2_status2")(1).Checked = True Then
            sh.Cells(lr, 102).Value = "B"
        ElseIf doc.SelectContentControlsByTitle("others2_status1")(1).Checked = True Then
            sh.Cells(lr, 102).Value = "A"
        End If
        
        sh.Cells(lr, 103).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others2_safe")(1).Range.Text)
        sh.Cells(lr, 104).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others2_violence")(1).Range.Text)
        sh.Cells(lr, 105).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others2_netcoop")(1).Range.Text)
        sh.Cells(lr, 106).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others2_netcoop_unit")(1).Range.Text)
        
        sh.Cells(lr, 107).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others3_name")(1).Range.Text)
        
        If doc.SelectContentControlsByTitle("others3_status3")(1).Checked = True Then
            sh.Cells(lr, 108).Value = "C"
        ElseIf doc.SelectContentControlsByTitle("others3_status2")(1).Checked = True Then
            sh.Cells(lr, 108).Value = "B"
        ElseIf doc.SelectContentControlsByTitle("others3_status1")(1).Checked = True Then
            sh.Cells(lr, 108).Value = "A"
        End If
        
        sh.Cells(lr, 109).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others3_safe")(1).Range.Text)
        sh.Cells(lr, 110).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others3_violence")(1).Range.Text)
        sh.Cells(lr, 111).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others3_netcoop")(1).Range.Text)
        sh.Cells(lr, 112).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("others3_netcoop_unit")(1).Range.Text)
        
        If doc.SelectContentControlsByTitle("declass_reso_f")(1).Checked = True Then
            sh.Cells(lr, 113).Value = "1"
        Else
            sh.Cells(lr, 113).Value = "0"
        End If
        sh.Cells(lr, 114).Value = Application.WorksheetFunction.Clean(doc.SelectContentControlsByTitle("resolution")(1).Range.Text)
        lr = lr + 1
    Next
    
    '此行以下勿動
    
    doc.Close
    wd.Quit
    Set doc = Nothing
    Set sh = Nothing
    Set wd = Nothing

End Sub
