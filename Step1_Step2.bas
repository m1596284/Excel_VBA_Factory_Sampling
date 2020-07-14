Attribute VB_Name = "Step1_Step2"
Sub Step1_customer()
    
    Call Unprotect
    
    customer_code = ActiveSheet.Range("D3").Text   '客戶代碼
    
    ' 清空舊資料
    Range("AW2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    
    Sheets("" & customer_code & "").Select
    Cells.Select
    On Error Resume Next
    ActiveSheet.ShowAllData
    
    total_amount = Sheets("" & customer_code & "").Cells(1, 2).End(xlDown).Row  '取得item數量
    
    '把所有ITEM複製到主頁
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("NEW").Select
    Range("AW1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'Sheets("NEW").Select
    'Range("D5:K6").Select   'data validation
    'With Selection.Validation
    '    .Delete
    '    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & customer_code & "!B2:B" & total_amount & ""
    '    .IgnoreBlank = True
    '    .InCellDropdown = True
    'End With
    
    'fill in test instrument
    If customer_code = "EB" Then
        Range("D7") = Sheets("Test_Instrument").Cells(3, 2).Text
        Range("D9") = Sheets("Test_Instrument").Cells(3, 3).Text
    ElseIf customer_code = "IBE" Then
        Range("D7") = Sheets("Test_Instrument").Cells(2, 2).Text
        Range("D9") = Sheets("Test_Instrument").Cells(2, 3).Text
    ElseIf customer_code = "WE" Then
        Range("D7") = Sheets("Test_Instrument").Cells(4, 2).Text
        Range("D9") = Sheets("Test_Instrument").Cells(4, 3).Text
    ElseIf customer_code = "北川" Then
        Range("D7") = Sheets("Test_Instrument").Cells(5, 2).Text
        Range("D9") = Sheets("Test_Instrument").Cells(5, 3).Text
    End If
    
    Range("D5").Select
    Selection.ClearContents
    
    Call Unprotect
    Call CleanOldData


End Sub

Sub Step2_item()
Attribute Step2_item.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Call Unprotect
    
    ActiveSheet.Range("R2") = Date
    
    '清除舊資料
    Range("A17:T65").Select
    Selection.ClearContents

    '讀取客戶和產品
    customer_code = ActiveSheet.Range("D3").Text
    item_no = ActiveSheet.Range("D5").Text
    
    '清除舊規則
    Range("F17:R41").Select
    Selection.FormatConditions.Delete
    
    If customer_code = "EB" Then
        GoTo EB
    ElseIf customer_code = "IBE" Then
        GoTo IBE
    ElseIf customer_code = "WE" Then
        GoTo WE
    Else
        GoTo 其他
    End If
    
    '----------------------------------------------------------------------------------------------------
EB:
    '資料篩選
    Sheets("" & customer_code & "").Select
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.ListObjects("表格1").Range.AutoFilter Field:=2, Criteria1:=item_no
    
    '變動欄位
    Range("A1").Select
    Selection.End(xlDown).Select
    Sheets("NEW").Range("C17") = ActiveCell.Offset(0, 2).Text   'A
    Sheets("NEW").Range("C21") = ActiveCell.Offset(0, 3).Text   'B
    Sheets("NEW").Range("C25") = ActiveCell.Offset(0, 4).Text   'N
    Sheets("NEW").Range("C26") = ActiveCell.Offset(0, 21).Text  'N(mm)
    Sheets("NEW").Range("C29") = ActiveCell.Offset(0, 17).Text  'Amp(1)
    Sheets("NEW").Range("C30") = ActiveCell.Offset(0, 5).Text   'DB50
    Sheets("NEW").Range("C33") = ActiveCell.Offset(0, 6).Text   'DB100
    Sheets("NEW").Range("C36") = ActiveCell.Offset(0, 7).Text   'DB200
    
    Sheets("NEW").Range("C39") = ActiveCell.Offset(0, 16).Text  'Amp(2)
    Sheets("NEW").Range("C40") = ActiveCell.Offset(0, 18).Text  'DB 50 (2)
    Sheets("NEW").Range("C43") = ActiveCell.Offset(0, 19).Text  'DB 100 (2)
    Sheets("NEW").Range("C46") = ActiveCell.Offset(0, 20).Text  'DB 200 (2)
    
    Sheets("NEW").Range("C50") = ActiveCell.Offset(0, 8).Text   'L
    Sheets("NEW").Range("C51") = ActiveCell.Offset(0, 9).Text   'L2
    Sheets("NEW").Range("C52") = ActiveCell.Offset(0, 22).Text  'L3

    'Sheets("NEW").Range("C46") = ActiveCell.Offset(0, 16).Text & "," & ActiveCell.Offset(0, 17).Text
    Sheets("NEW").Range("I9") = ActiveCell.Offset(0, 10).Text   'Thread
    
    
    '固定欄位
    Sheets("NEW").Range("J3") = "J30"
    Sheets("NEW").Range("A17") = "A"
    Sheets("NEW").Range("A21") = "B"
    Sheets("NEW").Range("A25") = "N"
    Sheets("NEW").Range("A29") = "Amp(1)"
    Sheets("NEW").Range("A30") = "50 MHZ"
    Sheets("NEW").Range("A33") = "100 MHZ"
    Sheets("NEW").Range("A36") = "200 MHZ"
    Sheets("NEW").Range("A39") = "Amp(2)"
    Sheets("NEW").Range("A40") = "50 MHZ(2)"
    Sheets("NEW").Range("A43") = "100 MHZ(2)"
    Sheets("NEW").Range("A46") = "200 MHZ(2)"
    Sheets("NEW").Range("A50") = "L"
    Sheets("NEW").Range("A51") = "L2"
    Sheets("NEW").Range("A52") = "L3"
    
    '設定上下限
    upper_A = ActiveCell.Offset(0, 12).Value   'A上
    lower_A = ActiveCell.Offset(0, 13).Value   'A下
    upper_B = ActiveCell.Offset(0, 14).Value   'B上
    lower_B = ActiveCell.Offset(0, 15).Value   'B下
    Sheets("NEW").Range("U17") = upper_A
    Sheets("NEW").Range("V17") = lower_A
    Sheets("NEW").Range("U21") = upper_B
    Sheets("NEW").Range("V21") = lower_B
    
    Sheets("NEW").Select
    Range("F17:R20").Select 'A的上下限
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_A & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_A & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    
    Range("F21:R24").Select 'B的上下限
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_B & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_B & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    
    Call shift2
    Call Protect
    Exit Sub
    
    '----------------------------------------------------------------------------------------------------
    
IBE:
    '資料篩選
    Sheets("" & customer_code & "").Select
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.ListObjects("表格2").Range.AutoFilter Field:=2, Criteria1:=item_no
    
    '變動欄位
    Range("A1").Select
    Selection.End(xlDown).Select
    Sheets("NEW").Range("C17") = ActiveCell.Offset(0, 2).Text
    Sheets("NEW").Range("C21") = ActiveCell.Offset(0, 3).Text
    Sheets("NEW").Range("C25") = ActiveCell.Offset(0, 4).Text 'N
    Sheets("NEW").Range("C26") = ActiveCell.Offset(0, 17).Text 'N(mm)
    Sheets("NEW").Range("C29") = ActiveCell.Offset(0, 16).Text 'Amp
    Sheets("NEW").Range("C30") = ActiveCell.Offset(0, 5).Text 'DB 50
    Sheets("NEW").Range("C33") = ActiveCell.Offset(0, 6).Text 'DB 100
    Sheets("NEW").Range("C36") = ActiveCell.Offset(0, 7).Text 'DB 200
    Sheets("NEW").Range("C39") = ActiveCell.Offset(0, 8).Text 'L
    Sheets("NEW").Range("C40") = ActiveCell.Offset(0, 9).Text 'L2
    Sheets("NEW").Range("C41") = ActiveCell.Offset(0, 18).Text 'L3
    
    
    Sheets("NEW").Range("I9") = ActiveCell.Offset(0, 10).Text & "/" & ActiveCell.Offset(0, 11).Text
    
    '固定欄位
    Sheets("NEW").Range("J3") = "J30"
    Sheets("NEW").Range("A17") = "A"
    Sheets("NEW").Range("A21") = "B"
    Sheets("NEW").Range("A25") = "N"
    Sheets("NEW").Range("A29") = "Amp"
    Sheets("NEW").Range("A30") = "50 MHZ"
    Sheets("NEW").Range("A33") = "100 MHZ"
    Sheets("NEW").Range("A36") = "200 MHZ"
    Sheets("NEW").Range("A39") = "L"
    Sheets("NEW").Range("A40") = "L2"
    Sheets("NEW").Range("A41") = "L3"
    
    '設定上下限
    upper_A = ActiveCell.Offset(0, 12).Value
    lower_A = ActiveCell.Offset(0, 13).Value
    upper_B = ActiveCell.Offset(0, 14).Value
    lower_B = ActiveCell.Offset(0, 15).Value
    Sheets("NEW").Range("U17") = upper_A
    Sheets("NEW").Range("V17") = lower_A
    Sheets("NEW").Range("U21") = upper_B
    Sheets("NEW").Range("V21") = lower_B
    
    Sheets("NEW").Select
    Range("F17:R20").Select 'A的上下限
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_A & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_A & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    
    Range("F21:R24").Select 'B的上下限
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_B & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_B & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
   
    
    Call shift2
    Call Protect
    Exit Sub
    '----------------------------------------------------------------------------------------------------
    
WE:
    '資料篩選
    Sheets("" & customer_code & "").Select
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.ListObjects("表格23").Range.AutoFilter Field:=2, Criteria1:=item_no
    
    '變動欄位
    Range("A1").Select
    Selection.End(xlDown).Select
    Sheets("NEW").Range("C17") = ActiveCell.Offset(0, 3).Text
    Sheets("NEW").Range("C21") = ActiveCell.Offset(0, 4).Text
    Sheets("NEW").Range("C25") = ActiveCell.Offset(0, 5).Text
    Sheets("NEW").Range("J7") = ActiveCell.Offset(0, 8).Text '星級
    Sheets("NEW").Range("I9") = ActiveCell.Offset(0, 9).Text & "/" & ActiveCell.Offset(0, 10).Text
    Sheets("NEW").Range("J3") = ActiveCell.Offset(0, 2).Text 'J70
    
    '固定欄位
    Sheets("NEW").Range("A17") = "A"
    Sheets("NEW").Range("A21") = "B"
    Sheets("NEW").Range("A25") = "C"
    Sheets("NEW").Range("A26") = "C"
    Sheets("NEW").Range("A27") = "C"
    Sheets("NEW").Range("A28") = "C"
    
    '特殊欄位
    Dim Item As Boolean
    Item = False
    
    Range("A1").Select
    Selection.End(xlDown).Select
    If ActiveCell.Offset(0, 6).Text <> "" Then  '查看D是否有資料
        Sheets("NEW").Range("A29") = "D"
        Sheets("NEW").Range("A30") = "D"
        Sheets("NEW").Range("A31") = "D"
        Sheets("NEW").Range("A32") = "D"
        Sheets("NEW").Range("C29") = ActiveCell.Offset(0, 6).Text
        'checking = "d only"
        
        If ActiveCell.Offset(0, 7).Text <> "" Then  '查看E是否有資料
            Sheets("NEW").Range("A33") = "E"
            Sheets("NEW").Range("A34") = "E"
            Sheets("NEW").Range("A35") = "E"
            Sheets("NEW").Range("A36") = "E"
            Sheets("NEW").Range("C33") = ActiveCell.Offset(0, 7).Text
            'checking = "d and e"
        End If
    End If
    
    If ActiveCell.Offset(0, 12).Text <> "" Then  '查看1MHZ是否有資料
            Item = True
            update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            Sheets("NEW").Range("A" & update_row & "") = "1MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 12).Text
    End If


    If ActiveCell.Offset(0, 13).Text <> "" Then  '查看10MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "10MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 13).Text
    End If
    
    If ActiveCell.Offset(0, 14).Text <> "" Then  '查看25MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "25MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 14).Text
    End If
    
    If ActiveCell.Offset(0, 15).Text <> "" Then  '查看30MHZ是否有資料
           If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "30MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 15).Text
    End If
    
    If ActiveCell.Offset(0, 16).Text <> "" Then  '查看50MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "50MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 16).Text
    End If
    
    If ActiveCell.Offset(0, 17).Text <> "" Then  '查看70MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "70MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 17).Text
    End If
    
    If ActiveCell.Offset(0, 18).Text <> "" Then  '查看100MHZ是否有資料
           If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "100MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 18).Text
    End If
    
    If ActiveCell.Offset(0, 19).Text <> "" Then  '查看200MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "200MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 19).Text
    End If
    
    If ActiveCell.Offset(0, 20).Text <> "" Then  '查看300MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "300MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 20).Text
    End If
    
    '設定上下限
    upper_A = ActiveCell.Offset(0, 21).Value
    lower_A = ActiveCell.Offset(0, 22).Value
    upper_B = ActiveCell.Offset(0, 23).Value
    lower_B = ActiveCell.Offset(0, 24).Value
    upper_C = ActiveCell.Offset(0, 25).Value
    lower_C = ActiveCell.Offset(0, 26).Value
    upper_D = ActiveCell.Offset(0, 27).Value
    lower_D = ActiveCell.Offset(0, 28).Value
    upper_E = ActiveCell.Offset(0, 29).Value
    lower_E = ActiveCell.Offset(0, 30).Value
    Sheets("NEW").Range("U17") = upper_A
    Sheets("NEW").Range("V17") = lower_A
    Sheets("NEW").Range("U21") = upper_B
    Sheets("NEW").Range("V21") = lower_B
    Sheets("NEW").Range("U25") = upper_C
    Sheets("NEW").Range("V25") = lower_C
    Sheets("NEW").Range("U29") = upper_D
    Sheets("NEW").Range("V29") = lower_D
    Sheets("NEW").Range("U33") = upper_E
    Sheets("NEW").Range("V33") = lower_E
    
    Sheets("NEW").Select
    
    Range("F17:R20").Select 'A的上下限
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_A & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_A & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    
    Range("F21:R24").Select 'B的上下限
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_B & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_B & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    
    Range("F25:R28").Select 'C的上下限
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_C & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_C & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    
    If Range("A29") = "D" Then
        Range("F29:R32").Select 'D的上下限
        Selection.FormatConditions.Delete
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_D & ""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Font.Color = -16776961
        Selection.FormatConditions(1).StopIfTrue = True
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_D & ""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Font.Color = -16776961
        Selection.FormatConditions(1).StopIfTrue = True
    End If
    
    If Range("A33") = "E" Then
        Range("F33:R36").Select 'E的上下限
        Selection.FormatConditions.Delete
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_E & ""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Font.Color = -16776961
        Selection.FormatConditions(1).StopIfTrue = True
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_E & ""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Font.Color = -16776961
        Selection.FormatConditions(1).StopIfTrue = True
    End If
    
    If Sheets("NEW").Range("A26") = "C" Then
        Sheets("NEW").Range("A26") = ""
        Sheets("NEW").Range("A27") = ""
        Sheets("NEW").Range("A28") = ""
    End If
    
    If Sheets("NEW").Range("A30") = "D" Then
        Sheets("NEW").Range("A30") = ""
        Sheets("NEW").Range("A31") = ""
        Sheets("NEW").Range("A32") = ""
    End If

    If Sheets("NEW").Range("A34") = "E" Then
        Sheets("NEW").Range("A34") = ""
        Sheets("NEW").Range("A35") = ""
        Sheets("NEW").Range("A36") = ""
    End If
    

    Call shift2
    Call Protect
    Exit Sub
    '----------------------------------------------------------------------------------------------------

其他:

    '資料篩選
    Sheets("" & customer_code & "").Select
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.ListObjects("" & customer_code & "").Range.AutoFilter Field:=2, Criteria1:=item_no
    
    '變動欄位
    Range("A1").Select
    Selection.End(xlDown).Select
    Sheets("NEW").Range("C17") = ActiveCell.Offset(0, 3).Text   'A
    Sheets("NEW").Range("C21") = ActiveCell.Offset(0, 4).Text   'B
    Sheets("NEW").Range("C25") = ActiveCell.Offset(0, 5).Text   'C
    Sheets("NEW").Range("J7") = ActiveCell.Offset(0, 8).Text    '星數
    Sheets("NEW").Range("I9") = ActiveCell.Offset(0, 9).Text & "/" & ActiveCell.Offset(0, 10).Text 'KHZ MHZ
    
    
    '固定欄位
    Sheets("NEW").Range("J3") = ActiveCell.Offset(0, 2).Text '材質
    Sheets("NEW").Range("A17") = "A"
    Sheets("NEW").Range("A21") = "B"
    Sheets("NEW").Range("A25") = "C"
    Sheets("NEW").Range("A26") = "C"
    Sheets("NEW").Range("A27") = "C"
    Sheets("NEW").Range("A28") = "C"
    
   '特殊欄位
    Range("A1").Select
    Selection.End(xlDown).Select
    If ActiveCell.Offset(0, 6).Text <> "" Then  '查看D是否有資料
        Sheets("NEW").Range("A29") = "D"
        Sheets("NEW").Range("A30") = "D"
        Sheets("NEW").Range("A31") = "D"
        Sheets("NEW").Range("A32") = "D"
        Sheets("NEW").Range("C29") = ActiveCell.Offset(0, 6).Text
        'checking = "d only"
        
        If ActiveCell.Offset(0, 7).Text <> "" Then  '查看E是否有資料
            Sheets("NEW").Range("A33") = "E"
            Sheets("NEW").Range("A34") = "E"
            Sheets("NEW").Range("A35") = "E"
            Sheets("NEW").Range("A36") = "E"
            Sheets("NEW").Range("C33") = ActiveCell.Offset(0, 7).Text
            'checking = "d and e"
        End If
    End If
    
    If ActiveCell.Offset(0, 12).Text <> "" Then  '查看1MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "1MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 12).Text
    End If


    If ActiveCell.Offset(0, 13).Text <> "" Then  '查看10MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "10MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 13).Text
    End If
    
    If ActiveCell.Offset(0, 14).Text <> "" Then  '查看25MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "25MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 14).Text
    End If
    
    If ActiveCell.Offset(0, 15).Text <> "" Then  '查看30MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "30MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 15).Text
    End If
    
    If ActiveCell.Offset(0, 16).Text <> "" Then  '查看50MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "50MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 16).Text
    End If
    
    If ActiveCell.Offset(0, 17).Text <> "" Then  '查看70MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "70MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 17).Text
    End If
    
    If ActiveCell.Offset(0, 18).Text <> "" Then  '查看100MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "100MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 18).Text
    End If
    
    If ActiveCell.Offset(0, 19).Text <> "" Then  '查看200MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "200MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 19).Text
    End If
    
    If ActiveCell.Offset(0, 20).Text <> "" Then  '查看300MHZ是否有資料
            If Item = True Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 4
            End If
            If Item = False Then
                update_row = Sheets("NEW").Cells(65, 1).End(xlUp).Row + 1
            End If
            Item = True
            Sheets("NEW").Range("A" & update_row & "") = "300MHZ"
            Sheets("NEW").Range("C" & update_row & "") = ActiveCell.Offset(0, 20).Text
    End If

    '設定上下限
    upper_A = ActiveCell.Offset(0, 21).Value
    lower_A = ActiveCell.Offset(0, 22).Value
    upper_B = ActiveCell.Offset(0, 23).Value
    lower_B = ActiveCell.Offset(0, 24).Value
    upper_C = ActiveCell.Offset(0, 25).Value
    lower_C = ActiveCell.Offset(0, 26).Value
    upper_D = ActiveCell.Offset(0, 27).Value
    lower_D = ActiveCell.Offset(0, 28).Value
    upper_E = ActiveCell.Offset(0, 29).Value
    lower_E = ActiveCell.Offset(0, 30).Value
    Sheets("NEW").Range("U17") = upper_A
    Sheets("NEW").Range("V17") = lower_A
    Sheets("NEW").Range("U21") = upper_B
    Sheets("NEW").Range("V21") = lower_B
    Sheets("NEW").Range("U25") = upper_C
    Sheets("NEW").Range("V25") = lower_C
    Sheets("NEW").Range("U29") = upper_D
    Sheets("NEW").Range("V29") = lower_D
    Sheets("NEW").Range("U33") = upper_E
    Sheets("NEW").Range("V33") = lower_E
    
    Sheets("NEW").Select
    
    Range("F17:R20").Select 'A的上下限
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_A & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_A & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    
    Range("F21:R24").Select 'B的上下限
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_B & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_B & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    
    Range("F25:R28").Select 'C的上下限
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_C & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_C & ""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Font.Color = -16776961
    Selection.FormatConditions(1).StopIfTrue = True
    
    If Range("A29") = "D" Then
        Range("F29:R32").Select 'D的上下限
        Selection.FormatConditions.Delete
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_D & ""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Font.Color = -16776961
        Selection.FormatConditions(1).StopIfTrue = True
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_D & ""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Font.Color = -16776961
        Selection.FormatConditions(1).StopIfTrue = True
    End If
    
    If Range("A33") = "E" Then
        Range("F33:R36").Select 'E的上下限
        Selection.FormatConditions.Delete
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & upper_E & ""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Font.Color = -16776961
        Selection.FormatConditions(1).StopIfTrue = True
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & lower_E & ""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Font.Color = -16776961
        Selection.FormatConditions(1).StopIfTrue = True
    End If
    
    If Sheets("NEW").Range("A26") = "C" Then
        Sheets("NEW").Range("A26") = ""
        Sheets("NEW").Range("A27") = ""
        Sheets("NEW").Range("A28") = ""
    End If
    
    If Sheets("NEW").Range("A30") = "D" Then
        Sheets("NEW").Range("A30") = ""
        Sheets("NEW").Range("A31") = ""
        Sheets("NEW").Range("A32") = ""
    End If

    If Sheets("NEW").Range("A34") = "E" Then
        Sheets("NEW").Range("A34") = ""
        Sheets("NEW").Range("A35") = ""
        Sheets("NEW").Range("A36") = ""
    End If
    

    Call shift2
    Call Protect
    Exit Sub

    '------------------------------------------------------------------------------------------------------------------
    

    
    
End Sub





