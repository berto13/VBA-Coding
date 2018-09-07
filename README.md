# VBA-Coding

'1.EXCEL MULTIPLE TABS CREATOR

Sub Tab_Creator_sym()

' Tab_Creator_sym Macro

'Variable Declaration
Dim ALDRDate1 As String
Dim ALDRDate2 As String
Dim Symbol As String
Dim SymRow As Integer
Dim ws As Worksheet
Dim ALDRDateShrt1 As String
Dim ALDRDateShrt2 As String
Dim ALDRTab1 As String
Dim ALDRTab2 As String
Dim strDataRange As Range
Dim keyRange As Range
Dim tbl As ListObject
Dim SymTab1 As String
Dim SymTab2 As String
Dim ExpF As String
Dim ExportFile As Workbook
Dim TabCreator As Workbook
Dim TabCreatorName As String

Set TabCreator = ThisWorkbook
TabCreatorName = ThisWorkbook.Name

'First row with Symbol Value
SymRow = 11

'Variable Designation
ALDRTab1 = ActiveWorkbook.Sheets("Main").Range("E3")
ALDRTab2 = ActiveWorkbook.Sheets("Main").Range("E4")
ALDRDateShrt1 = ActiveWorkbook.Sheets("Main").Range("D3")
ALDRDateShrt2 = ActiveWorkbook.Sheets("Main").Range("D4")
ALDRDate1 = ActiveWorkbook.Sheets("Main").Range("C3")
ALDRDate2 = ActiveWorkbook.Sheets("Main").Range("C4")
Symbol = ActiveWorkbook.Sheets("Main").Range("B" & SymRow)


ExpF = ActiveWorkbook.Sheets("Main").Range("K8")

'Initial Errors Check
If ExpF = "" Then
    MsgBox ("Export File not Selected."), , "Error with Export File"
    Range("J7:J8").Style = "Bad"
    Range("K8").Style = "Bad"
    Exit Sub
End If
If Sheets("Main").CheckBox2.Value = False Then
    MsgBox ("Please confirm that the ALDR data has been updated. Verify Process Check #1"), , "Error with Process Check #1"
    Range("J3:J4").Style = "Bad"
    Exit Sub
End If
If Sheets("Main").CheckBox1.Value = False Then
    MsgBox ("Please confirm that 'As Of Date' and ALDR TabNames have been updated. Verify Process Check #2"), , "Error with Process Check #2"
    Range("J5:J6").Style = "Bad"
    Exit Sub
End If
    
    
Set ExportFile = Workbooks.Open(Filename:=ExpF)
TabCreator.Activate
'Start of main code
'Loop to create tabs until cell is blank
Do
    'Verify if checkbox is checked for tab with or without date
    If Sheets("Main").WDateChk.Value = True Then
        SymTab1 = Symbol & " " & ALDRDateShrt1
        SymTab2 = Symbol & " " & ALDRDateShrt2
    Else
        SymTab1 = Symbol
    End If
    
'With ALDR1
    Sheets(ALDRTab1).Select
    'To remove any previous filter
    ActiveSheet.Range("A2").AutoFilter
    'Identify last row with value
    Lastrow = ActiveSheet.Cells(Rows.Count, "L").End(xlUp).Row
    'Set the range of data
    Set strDataRange = ActiveSheet.Range("A2:L" & Lastrow)
    Set keyRange = ActiveSheet.Range("A2")
    'Filter by Symbol
    ActiveSheet.Range("A2").AutoFilter Field:=1, Criteria1:=Symbol
    'Select and copy visible cells of returned data after filter
    strDataRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    'Create and name tab with Symbol
    ExportFile.Activate
    Set ws = ExportFile.Sheets.Add(After:= _
             ExportFile.Sheets(ExportFile.Sheets.Count))
    ws.Name = SymTab1
    'Paste data in the new tab created
    ActiveSheet.Paste
    'Format to table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.TableStyle = "TableStyleMedium2"
    'Format Short QTY and sum
    Range("H:H").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Range("H1").Select
    Lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Selection.End(xlDown).Offset(1, 0).Select
    Selection.Formula = "=SUM(H2:H" & Lastrow & ")"
    'Go to Main tab and copy template for reportable/non-reportable Shorts with appropriate date
    TabCreator.Activate
    Sheets("Main").Select
    Range("G1:H7").Select
    Selection.Copy
    ExportFile.Activate
    Sheets(SymTab1).Select
    Range("B1").Select
    Selection.End(xlDown).Offset(4, 0).Select
    ActiveSheet.Paste
    Range("A1").Select
    Selection.End(xlDown).Offset(4, 0).Select
    Selection.Value = ALDRDate1
    'Fit columns to content
    Cells.Select
    Cells.EntireColumn.AutoFit
    If ActiveSheet.Range("A2").Value = "" Then
        'Highlight Tab red if no data returned
        ActiveSheet.Tab.ColorIndex = 3 '3=Red , 4=green,5=blue,6=yellow
        TabCreator.Activate
        Sheets("Main").Select
    Else
        'Paste Short QTY total into the Main tab
        Range("H1").Select
        Selection.End(xlDown).Select
        Selection.Copy
        TabCreator.Activate
        Sheets("Main").Select
        Range("C" & SymRow).Select
        Selection.PasteSpecial xlPasteValues
    
    End If
    
'Do only if more than one ALDR date selected
If Sheets("Main").WDateChk.Value = True Then
'Same code as above but now for ALDR2
'With ALDR2
    Sheets(ALDRTab2).Select
    'To remove any previous filter
    ActiveSheet.Range("A2").AutoFilter
    'Identify last row with value
    Lastrow = ActiveSheet.Cells(Rows.Count, "L").End(xlUp).Row
    'Set the range of data
    Set strDataRange = ActiveSheet.Range("A2:L" & Lastrow)
    Set keyRange = ActiveSheet.Range("A2")
    'Filter by Symbol
    ActiveSheet.Range("A2").AutoFilter Field:=1, Criteria1:=Symbol
    'Select and copy visible cells of returned data after filter
    strDataRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    'Create and name tab with Symbol
    ExportFile.Activate
    Set ws = ExportFile.Sheets.Add(After:= _
             ExportFile.Sheets(ExportFile.Sheets.Count))
    ws.Name = SymTab2
    'Paste data in the new tab created
    ActiveSheet.Paste
    'Format to table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.TableStyle = "TableStyleMedium2"
    'Format Short QTY and sum
    Range("H:H").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Range("H1").Select
    Lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Selection.End(xlDown).Offset(1, 0).Select
    Selection.Formula = "=SUM(H2:H" & Lastrow & ")"
    'Go to Main tab and copy template for reportable/non-reportable Shorts with appropriate date
    TabCreator.Activate
    Sheets("Main").Select
    Range("G1:H7").Select
    Selection.Copy
    ExportFile.Activate
    Sheets(SymTab2).Select
    Range("B1").Select
    Selection.End(xlDown).Offset(4, 0).Select
    ActiveSheet.Paste
    Range("A1").Select
    Selection.End(xlDown).Offset(4, 0).Select
    Selection.Value = ALDRDate2
    'Fit columns to content
    Cells.Select
    Cells.EntireColumn.AutoFit
    If ActiveSheet.Range("A2").Value = "" Then
        'Highlight Tab red if no data returned
        ActiveSheet.Tab.ColorIndex = 3 '3=Red , 4=green,5=blue,6=yellow
        TabCreator.Activate
        Sheets("Main").Select
    Else
        'Paste Short QTY total into the Main tab
        Range("H1").Select
        Selection.End(xlDown).Select
        Selection.Copy
        TabCreator.Activate
        Sheets("Main").Select
        Range("D" & SymRow).Select
        Selection.PasteSpecial xlPasteValues
    
    End If
End If
    'Go the the next row for the next Symbol
    SymRow = SymRow + 1
    Symbol = ActiveWorkbook.Sheets("Main").Range("B" & SymRow)
    
Loop While (Symbol <> "")
'End of main code

Range("B10").Select
Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
'Set the range of data
Set strDataRange = ActiveSheet.Range("B11:D" & Lastrow)
Set keyRange = ActiveSheet.Range("B10")
strDataRange.Select
Selection.Copy
Workbooks.Open ("\\rutvnazcti0089\OPS_REGULATORY_CONTROL\Short Interest Reporting\SIR Enhancements\LogTabCreator.xlsx")
Range("A2").Select
Selection.End(xlDown).Offset(1, 0).Select
ActiveSheet.Paste
Range("A2").Select
Selection.End(xlDown).Offset(0, 3).Select
ActiveSheet.Range(Selection, Selection.End(xlUp).Offset(1, 0)).Select
Selection.Value = Now
Workbooks("LogTabCreator.xlsx").Save
Workbooks("LogTabCreator.xlsx").Close


TabCreator.Activate
Range("B10").Select
Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
Set strDataRange = ActiveSheet.Range("B9:D" & Lastrow)
Set keyRange = ActiveSheet.Range("B10")
strDataRange.Select
Selection.Copy
Dim UserLog As Workbook
Set UserLog = Workbooks.Add
Range("A2").Select
ActiveSheet.Paste
TabCreator.Activate
Range("C10:D10").Copy
UserLog.Activate
Range("B3:C3").PasteSpecial xlValues
Range("A1").Value = "Tab Creator Log"
Cells.Select
Cells.EntireColumn.AutoFit

TabCreator.Activate
Sheets("Main").CheckBox2.Value = False
Sheets("Main").CheckBox1.Value = False
Range("K8").Value = ""
Range("J7:J8").Style = "Bad"
Range("K8").Style = "Bad"
Range("B11:D260").ClearContents
ActiveWorkbook.Save

ExportFile.Activate
MsgBox ("Tabs Created Successfully!"), , "Done!"

TabCreator.Close

End Sub
---------------------------------------------------------------------------------------------------------------------------------------

'2. ACCESS DATABASE DATA SCRAPER

Option Compare Database


Private Sub Command11_Click()

 Dim MyXL As Object

    Set MyXL = CreateObject("Excel.Application")
    With MyXL
        .Application.Visible = True
        .Workbooks.Open "L:\WWSSLA\LaBilling\RPCPROD\Argentina\SEBILL PROD\Acreedias\Acreedias_DB\dividend_charges.xlsx"
    End With


End Sub

Private Sub Command14_Click()

 Dim MyXL As Object

    Set MyXL = CreateObject("Excel.Application")
    With MyXL
        .Application.Visible = True
        .Workbooks.Open "L:\WWSSLA\LaBilling\RPCPROD\Argentina\SEBILL PROD\Acreedias\Acreedias_DB\safekeeping_no_charge.xlsx"
    End With
End Sub

Private Sub Command7_Click()


DoCmd.SetWarnings False
DoCmd.OpenQuery "qry_delete_tbl_scrape", acViewNormal, acEdit


'-------
' Get the main system object
    Dim Sessions As Object
    Dim System As Object
    Set System = CreateObject("EXTRA.System")   ' Gets the system object
    If (System Is Nothing) Then
        MsgBox "Could not create the EXTRA System object.  Stopping macro playback."
        Stop
    End If
    Set Sessions = System.Sessions

    If (Sessions Is Nothing) Then
        MsgBox "Could not create the Sessions collection object.  Stopping macro playback."
        Stop
    End If
'------
' Set the default wait timeout value
    g_HostSettleTime = 50    ' milliseconds

    OldSystemTimeout& = System.TimeoutValue
    If (g_HostSettleTime > OldSystemTimeout) Then
        System.TimeoutValue = g_HostSettleTime
    End If

' Get the necessary Session Object
    Dim Sess0 As Object
    Set Sess0 = System.ActiveSession
    If (Sess0 Is Nothing) Then
        MsgBox "Could not create the Session object.  Stopping macro playback."
        Stop
    End If
    If Not Sess0.Visible Then Sess0.Visible = True
    Sess0.Screen.WaitHostQuiet (g_HostSettleTime)



'** Prepare the database:

   Dim db As Database, rs As Recordset, rs2 As Recordset
   Dim i As Integer
   Dim opening_balance As String
   Dim trans_ref As String
   Dim trcd As String
   Dim customer_reference As String
   Dim tran_amount As String
   Dim settle_date As String
   Dim Drcr As String
   Dim broker As String
   Dim status As String
   Dim closing_balance As String
   Dim a As Integer
   Dim b As Integer
   Dim debit_account_number As String
   Dim credit_account_number As String
   Dim ccy As String
   Dim amount As String
   Dim Payment_method As String
   Dim customer_ref As String
   Dim instructed_on As String
   Dim value_date As String
   Dim free_text1 As String
   Dim free_text2 As String
   Dim free_text3 As String
   Dim debit_ccy_amount As String
   Dim credit_ccy_amount As String
   Dim exchange_rate As String
   Dim TIME_STARTED, TIME_COMPLETED
   Dim ACCOUNTS_OK, ACCOUNTS_ERR
   
   TIME_STARTED = Time()

   Set db = CurrentDb
   Set rs = db.OpenRecordset("tbl_scrape")
  
   'rs.MoveLast
   
   'txtTotRecs = rs.RecordCount
   
   'rs.AddNew
   
      
   DoEvents

If (Sess0.Screen.getstring(1, 33, 4)) = "HIST" Then
GoTo de_NUEVO
Else
MsgBox "INCORRECT HOST SCREEN, PLEASE TRY AGAIN"
GoTo FINAL
End If

de_NUEVO:

For a = 7 To 17


'** Start the main loop:
      
        rs.AddNew
        
        Sess0.Screen.MoveTo a, 2
                
        If (Sess0.Screen.getstring(a, 73, 3)) = "SET" Then
        GoTo siguiente
        Else
        End If
        
        If (Sess0.Screen.getstring(a, 17, 3)) = "800" Then
        GoTo comenzar
        Else
        If (Sess0.Screen.getstring(a, 17, 3)) = "900" Then
        GoTo comenzar
        Else
        GoTo FINAL
        End If
        End If
        
comenzar:
        rs!opening_balance = LTrim(Sess0.Screen.getstring(3, 57, 24))
        rs!trans_ref = Sess0.Screen.getstring(a, 5, 11)
        rs!trcd = Sess0.Screen.getstring(a, 17, 3)
        rs!customer_reference = Sess0.Screen.getstring(a, 22, 16)
        rs!tran_amount = LTrim(Sess0.Screen.getstring(a, 55, 26))
        rs!settle_date = Sess0.Screen.getstring(a + 1, 5, 11)
        rs!Drcr = Sess0.Screen.getstring(a + 1, 18, 1)
        rs!broker = Sess0.Screen.getstring(a + 1, 28, 39)
        rs!status = Sess0.Screen.getstring(a + 1, 69, 7)
        rs!closing_balance = LTrim(Sess0.Screen.getstring(19, 57, 24))
        
        If (Sess0.Screen.getstring(a, 17, 3)) = "800" Then
        GoTo pago_local
        Else
        GoTo fx
        End If

pago_local:
        Sess0.Screen.SendKeys ("ds")
        Sess0.Screen.SendKeys ("<enter>")
        Sess0.Screen.WaitHostQuiet (g_HostSettleTime)
        Sess0.Screen.WaitForCursor 23, 10
        rs!debit_account_number = Sess0.Screen.getstring(6, 23, 10)
        rs!credit_account_number = Sess0.Screen.getstring(7, 23, 10)
        rs!ccy = Sess0.Screen.getstring(8, 23, 3)
        rs!amount = LTrim(Sess0.Screen.getstring(9, 23, 24))
        rs!Payment_method = Sess0.Screen.getstring(10, 23, 2)
        rs!customer_ref = Sess0.Screen.getstring(12, 23, 24)
        rs!instructed_on = Sess0.Screen.getstring(13, 23, 11)
        rs!value_date = Sess0.Screen.getstring(14, 23, 11)
        rs!free_text1 = RTrim(Sess0.Screen.getstring(14, 23, 50))
        rs!free_text2 = RTrim(Sess0.Screen.getstring(15, 23, 50))
        rs!free_text3 = RTrim(Sess0.Screen.getstring(16, 23, 50))
        Sess0.Screen.MoveTo 23, 10
        Sess0.Screen.SendKeys ("ret")
        Sess0.Screen.SendKeys ("<enter>")
        Sess0.Screen.WaitForCursor 7, 2
        rs.Update
        GoTo siguiente
        
fx:
        Sess0.Screen.SendKeys ("ds")
        Sess0.Screen.SendKeys ("<enter>")
        Sess0.Screen.WaitHostQuiet (g_HostSettleTime)
        Sess0.Screen.WaitForCursor 23, 10
        rs!debit_account_number = Sess0.Screen.getstring(6, 23, 10)
        rs!debit_ccy_amount = RTrim(Sess0.Screen.getstring(7, 23, 25))
        rs!credit_account_number = Sess0.Screen.getstring(8, 23, 10)
        rs!credit_ccy_amount = RTrim(Sess0.Screen.getstring(9, 23, 25))
        'rs!ccy = Sess0.Screen.GetString(8, 23, 3)
        rs!customer_ref = RTrim(Sess0.Screen.getstring(12, 23, 24))
        rs!instructed_on = Sess0.Screen.getstring(13, 23, 11)
        rs!value_date = Sess0.Screen.getstring(16, 23, 11)
        rs!free_text1 = RTrim(Sess0.Screen.getstring(18, 23, 50))
        rs!free_text2 = RTrim(Sess0.Screen.getstring(19, 23, 50))
        'rs!free_text3 = RTrim(Sess0.Screen.GetString(16, 23, 50))
        Sess0.Screen.MoveTo 23, 10
        Sess0.Screen.SendKeys ("ret")
        Sess0.Screen.SendKeys ("<enter>")
        Sess0.Screen.WaitForCursor 7, 2
        rs.Update
        GoTo siguiente
    
    
siguiente:

Next a

Sess0.Screen.MoveTo 23, 10
Sess0.Screen.SendKeys ("nxt")
Sess0.Screen.SendKeys ("<enter>")
Sess0.Screen.WaitHostQuiet (g_HostSettleTime)
Sess0.Screen.WaitForCursor 7, 2



GoTo next_pages
'End If



next_pages:

For b = 9 To 17


'** Start the main loop:
      
        rs.AddNew
        
        Sess0.Screen.MoveTo b, 2
                
        If (Sess0.Screen.getstring(b, 73, 3)) = "SET" Then
        GoTo siguiente1
        Else
        End If
        
        If (Sess0.Screen.getstring(b, 17, 3)) = "800" Then
        GoTo comenzar1
        Else
        If (Sess0.Screen.getstring(b, 17, 3)) = "900" Then
        GoTo comenzar1
        Else
        GoTo FINAL
        End If
        End If
        
comenzar1:
        rs!opening_balance = LTrim(Sess0.Screen.getstring(3, 57, 24))
        rs!trans_ref = Sess0.Screen.getstring(b, 5, 11)
        rs!trcd = Sess0.Screen.getstring(b, 17, 3)
        rs!customer_reference = Sess0.Screen.getstring(b, 22, 16)
        rs!tran_amount = LTrim(Sess0.Screen.getstring(b, 55, 26))
        rs!settle_date = Sess0.Screen.getstring(b + 1, 5, 11)
        rs!Drcr = Sess0.Screen.getstring(b + 1, 18, 1)
        rs!broker = Sess0.Screen.getstring(b + 1, 28, 39)
        rs!status = Sess0.Screen.getstring(b + 1, 69, 7)
        rs!closing_balance = LTrim(Sess0.Screen.getstring(19, 57, 24))
        
        If (Sess0.Screen.getstring(b, 17, 3)) = "800" Then
        GoTo pago_local1
        Else
        GoTo fx1
        End If

pago_local1:
        Sess0.Screen.SendKeys ("ds")
        Sess0.Screen.SendKeys ("<enter>")
        Sess0.Screen.WaitHostQuiet (g_HostSettleTime)
        Sess0.Screen.WaitForCursor 23, 10
        rs!debit_account_number = Sess0.Screen.getstring(6, 23, 10)
        rs!credit_account_number = Sess0.Screen.getstring(7, 23, 10)
        rs!ccy = Sess0.Screen.getstring(8, 23, 3)
        rs!amount = LTrim(Sess0.Screen.getstring(9, 23, 24))
        rs!Payment_method = Sess0.Screen.getstring(10, 23, 2)
        rs!customer_ref = Sess0.Screen.getstring(12, 23, 24)
        rs!instructed_on = Sess0.Screen.getstring(13, 23, 11)
        rs!value_date = Sess0.Screen.getstring(14, 23, 11)
        rs!free_text1 = RTrim(Sess0.Screen.getstring(14, 23, 50))
        rs!free_text2 = RTrim(Sess0.Screen.getstring(15, 23, 50))
        rs!free_text3 = RTrim(Sess0.Screen.getstring(16, 23, 50))
        Sess0.Screen.MoveTo 23, 10
        Sess0.Screen.SendKeys ("ret")
        Sess0.Screen.SendKeys ("<enter>")
        Sess0.Screen.WaitForCursor 7, 2
        rs.Update
        GoTo siguiente1
        
fx1:
        Sess0.Screen.SendKeys ("ds")
        Sess0.Screen.SendKeys ("<enter>")
        Sess0.Screen.WaitHostQuiet (g_HostSettleTime)
        Sess0.Screen.WaitForCursor 23, 10
        rs!debit_account_number = Sess0.Screen.getstring(6, 23, 10)
        rs!debit_ccy_amount = RTrim(Sess0.Screen.getstring(7, 23, 25))
        rs!credit_account_number = Sess0.Screen.getstring(8, 23, 10)
        rs!credit_ccy_amount = RTrim(Sess0.Screen.getstring(9, 23, 25))
        'rs!ccy = Sess0.Screen.GetString(8, 23, 3)
        rs!customer_ref = RTrim(Sess0.Screen.getstring(12, 23, 24))
        rs!instructed_on = Sess0.Screen.getstring(13, 23, 11)
        rs!value_date = Sess0.Screen.getstring(16, 23, 11)
        rs!free_text1 = RTrim(Sess0.Screen.getstring(18, 23, 50))
        rs!free_text2 = RTrim(Sess0.Screen.getstring(19, 23, 50))
        'rs!free_text3 = RTrim(Sess0.Screen.GetString(16, 23, 50))
        Sess0.Screen.MoveTo 23, 10
        Sess0.Screen.SendKeys ("ret")
        Sess0.Screen.SendKeys ("<enter>")
        Sess0.Screen.WaitForCursor 7, 2
        rs.Update
        GoTo siguiente1
    
 
siguiente1:

Next b

Sess0.Screen.MoveTo 23, 10
Sess0.Screen.SendKeys ("nxt")
Sess0.Screen.SendKeys ("<enter>")
Sess0.Screen.WaitHostQuiet (g_HostSettleTime)
Sess0.Screen.WaitForCursor 7, 2


GoTo next_pages


FINAL:
   rs.Close
   TIME_COMPLETED = Time()
   MsgBox "FIN DE PROCESO !!!!" & Chr(10) & _
          " " & Chr(10) & _
          "STARTED: " & TIME_STARTED & "  COMPLETED: " & TIME_COMPLETED

Exit_COMMAND7_Click:
    Exit Sub


End Sub

---------------------------------------------------------------------------------------------------------------------------------------
'3. DATA ORGANIZER - REPORT

Private Sub CommandButton2_Click()

Dim strDataRange As Range
Dim keyRange As Range

Dim TabName As String
Dim RowNum As Integer
Dim Entity As String
Dim Entity2 As String

Dim SFX1 As String
Dim SFX2 As String
Dim SFX3 As String
SFX1 = ActiveWorkbook.Sheets("Summary").Range("K2")
SFX2 = ActiveWorkbook.Sheets("Summary").Range("K3")
SFX3 = ActiveWorkbook.Sheets("Summary").Range("K4")

Entity = ActiveWorkbook.Sheets("Summary").Range("C9")
Entity2 = ActiveWorkbook.Sheets("Summary").Range("J9")
RowNum = 1

TabName = ActiveWorkbook.Sheets("Summary").Range("A1").Offset(RowNum, 0).Value

Application.ScreenUpdating = False

ActiveWorkbook.Sheets(TabName).Activate
Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row

'PID Tab
    Do Until Not IsEmpty(Sheets(TabName).Range("E1").Value)

        ActiveSheet.Range("D1:D2").Copy
        ActiveSheet.Range("E1").Select
        ActiveSheet.Paste
        ActiveSheet.Range("E1").Value = "ABS Variance"
        ActiveSheet.Range("E2").Formula = "=ABS(D2)"

    'Application.ScreenUpdating = False

        ActiveSheet.Range("E2").Copy
        ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Offset(0, 3).Select
        ActiveSheet.Range(Selection, Selection.End(xlUp)).Select
        ActiveSheet.Paste
    
    'PID Filter Entity1

        Set strDataRange = ActiveSheet.Range("A1:E" & Lastrow)
        Set keyRange = ActiveSheet.Range("E1")
        strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
        strDataRange.Select
        Selection.AutoFilter
        If Entity = "All" Then
            Selection.AutoFilter
        Else
            Selection.AutoFilter Field:=1, Criteria1:=Entity 'column A
        End If
    Loop

    'PID Filter if already formatted
    Set strDataRange = ActiveSheet.Range("A1:E" & Lastrow)
    Set keyRange = ActiveSheet.Range("E1")
    strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
    strDataRange.Select
    Selection.AutoFilter
    If Entity = "All" Then
        Selection.AutoFilter
    Else
        Selection.AutoFilter Field:=1, Criteria1:=Entity 'column A
    End If

    'CopyPaste data into Temp sheet and then into Summary tab

    Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
    Set strDataRange = ActiveSheet.Range("A1:F" & Lastrow)
    Set keyRange = ActiveSheet.Range("F1")
    strDataRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
            
    ActiveWorkbook.Sheets("TempSheet").Activate
    ActiveSheet.Range("A1").Select
    ActiveSheet.Paste
            
    If SFXChkBox.Value = True Then
        Selection.AutoFilter Field:=2, Criteria1:=Array(SFX1, SFX2, SFX3), Operator:=xlFilterValues

        Set strDataRange = ActiveSheet.Range("A2:G" & Lastrow)
        Set keyRange = ActiveSheet.Range("G1")
        Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
        strDataRange.SpecialCells(xlCellTypeVisible).Select
        Selection.EntireRow.Delete
        ActiveSheet.Range("A1").Select
        Selection.AutoFilter
    End If
    If OtherPidExcl.Value = True Then
        Application.Run "Module1.AdditionalExclusions"
    End If
        
    ActiveSheet.Range("A2:D41").Copy
    ActiveWorkbook.Sheets("Summary").Activate
    Range("A14").PasteSpecial xlPasteValues
    Set strDataRange = ActiveSheet.Range("A14:F53")
    Set keyRange = ActiveSheet.Range("B13")
    strDataRange.Sort Key1:=keyRange, Order1:=xlDescending


    ActiveWorkbook.Sheets("TempSheet").Range("A:L").Delete

    'PID Entity2 Filter
    ActiveWorkbook.Sheets(TabName).Activate
    Set strDataRange = ActiveSheet.Range("A1:E" & Lastrow)
    Set keyRange = ActiveSheet.Range("E1")
    Selection.AutoFilter
    strDataRange.Select
    Selection.AutoFilter
    If Entity2 = "All" Then
        Selection.AutoFilter
    Else
        Selection.AutoFilter Field:=1, Criteria1:=Entity2 'column A
    End If
    'CopyPaste data into Temp sheet and then into Summary tab
    Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
    Set strDataRange = ActiveSheet.Range("A1:F" & Lastrow)
    Set keyRange = ActiveSheet.Range("F1")
    strDataRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
            
    ActiveWorkbook.Sheets("TempSheet").Activate
    ActiveSheet.Range("A1").Select
    ActiveSheet.Paste
            
    If SFXChkBox.Value = True Then
        Selection.AutoFilter Field:=2, Criteria1:=Array(SFX1, SFX2, SFX3), Operator:=xlFilterValues

        Set strDataRange = ActiveSheet.Range("A2:G" & Lastrow)
        Set keyRange = ActiveSheet.Range("G1")
        Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
        strDataRange.SpecialCells(xlCellTypeVisible).Select
        Selection.EntireRow.Delete
        ActiveSheet.Range("A1").Select
        Selection.AutoFilter
    End If
            
            
    If OtherPidExcl.Value = True Then
        Application.Run "Module1.AdditionalExclusions"
    End If
        
    ActiveSheet.Range("A2:D41").Copy
    ActiveWorkbook.Sheets("Summary").Activate
    Range("H14").PasteSpecial xlPasteValues
    Set strDataRange = ActiveSheet.Range("H14:M53")
    Set keyRange = ActiveSheet.Range("I13")
    strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
    ActiveWorkbook.Sheets("TempSheet").Range("A:L").Delete


    RowNum = RowNum + 1
    TabName = ActiveWorkbook.Sheets("Summary").Range("A1").Offset(RowNum, 0).Value

'End of PID Tab

'All Other Tabs
    Do Until (RowNum = 8)

        ActiveWorkbook.Sheets(TabName).Activate
    
    'For Maturity Tab Only
        If TabName = "Maturity " Then
    
            If ActiveWorkbook.Sheets(TabName).Range("G1").Value = "" Then
            
                Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
            
                ActiveSheet.Range("F1:F2").Copy
                ActiveSheet.Range("G1").Select
                ActiveSheet.Paste
                ActiveSheet.Range("G1").Value = "ABS Total Variance"
                ActiveSheet.Range("G2").Formula = "=ABS(D2+E2+F2)"

            'Application.ScreenUpdating = False

                ActiveSheet.Range("G2").Copy
                ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Offset(0, 5).Select
                ActiveSheet.Range(Selection, Selection.End(xlUp)).Select
                ActiveSheet.Paste
        
            Else
                'nothing
            End If
        
    'Maturity Entity1 Filter
        ActiveSheet.Range("A1").Select
        Selection.AutoFilter
        Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
        Set strDataRange = ActiveSheet.Range("A1:G" & Lastrow)
        Set keyRange = ActiveSheet.Range("G1")
        strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
        Selection.AutoFilter
        strDataRange.Select
        Selection.AutoFilter
        If Entity = "All" Then
            Selection.AutoFilter
        Else
            Selection.AutoFilter Field:=1, Criteria1:=Entity, Field:=4, Criteria1:="<>" 'column A
        End If
            
        Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
        Set strDataRange = ActiveSheet.Range("A1:G" & Lastrow)
        Set keyRange = ActiveSheet.Range("G1")
        strDataRange.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
            
        ActiveWorkbook.Sheets("TempSheet").Activate
        ActiveSheet.Range("A1").Select
        ActiveSheet.Paste
            
        Set strDataRange = ActiveSheet.Range("A2:G" & Lastrow)
        Set keyRange = ActiveSheet.Range("G1")
        Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
        If SFXChkBox.Value = True Then
            Selection.AutoFilter Field:=2, Criteria1:=Array(SFX1, SFX2, SFX3), Operator:=xlFilterValues
        
            strDataRange.SpecialCells(xlCellTypeVisible).Select
            Selection.EntireRow.Delete
            ActiveSheet.Range("A1").Select
            Selection.AutoFilter
        End If
         
        If OtherPidExcl.Value = True Then
            Application.Run "Module1.AdditionalExclusions"
        End If
            
        ActiveSheet.Range("A2:F41").Copy
        ActiveWorkbook.Sheets("Summary").Activate
        Range("A60").PasteSpecial xlPasteValues
        Set strDataRange = ActiveSheet.Range("A60:F99")
        Set keyRange = ActiveSheet.Range("B59")
        strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
        
        ActiveWorkbook.Sheets("TempSheet").Range("A:L").Delete
        
    'Maturity Entity2 Filter
        ActiveWorkbook.Sheets(TabName).Activate
        Set strDataRange = ActiveSheet.Range("A1:G" & Lastrow)
        Set keyRange = ActiveSheet.Range("G1")
        Selection.AutoFilter
        strDataRange.Select
        Selection.AutoFilter
        If Entity2 = "All" Then
            Selection.AutoFilter
        Else
            Selection.AutoFilter Field:=1, Criteria1:=Entity2 'column A
        End If
            
        Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
        Set strDataRange = ActiveSheet.Range("A1:G" & Lastrow)
        Set keyRange = ActiveSheet.Range("G1")
        strDataRange.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
            
        ActiveWorkbook.Sheets("TempSheet").Activate
        ActiveSheet.Range("A1").Select
        ActiveSheet.Paste
            
        Set strDataRange = ActiveSheet.Range("A2:G" & Lastrow)
        Set keyRange = ActiveSheet.Range("G1")
        Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
        If SFXChkBox.Value = True Then
            Selection.AutoFilter Field:=2, Criteria1:=Array(SFX1, SFX2, SFX3), Operator:=xlFilterValues
        
            strDataRange.SpecialCells(xlCellTypeVisible).Select
            Selection.EntireRow.Delete
            ActiveSheet.Range("A1").Select
            Selection.AutoFilter
        End If
            
        If OtherPidExcl.Value = True Then
            Application.Run "Module1.AdditionalExclusions"
        End If
            
        ActiveSheet.Range("A2:F41").Copy
        ActiveWorkbook.Sheets("Summary").Activate
        Range("H60").PasteSpecial xlPasteValues
        Set strDataRange = ActiveSheet.Range("H60:M99")
        Set keyRange = ActiveSheet.Range("I59")
        strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
        
        ActiveWorkbook.Sheets("TempSheet").Range("A:L").Delete
        
        RowNum = RowNum + 1
        TabName = ActiveWorkbook.Sheets("Summary").Range("A1").Offset(RowNum, 0).Value
        
        Else
        
'End of Maturity Tab
    
'All Other Tabs
    
            If ActiveWorkbook.Sheets(TabName).Range("F1").Value = "" Then

                Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
            
                ActiveSheet.Range("E1:E2").Copy
                ActiveSheet.Range("F1").Select
                ActiveSheet.Paste
                ActiveSheet.Range("F1").Value = "ABS Variance"
                ActiveSheet.Range("F2").Formula = "=ABS(E2)"

                'Application.ScreenUpdating = False

                ActiveSheet.Range("F2").Copy
                ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Offset(0, 4).Select
                ActiveSheet.Range(Selection, Selection.End(xlUp)).Select
                ActiveSheet.Paste
    
                Set strDataRange = ActiveSheet.Range("A1:F" & Lastrow)
                Set keyRange = ActiveSheet.Range("F1")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending

                'Filter Out Blanks
                strDataRange.Select
                Selection.AutoFilter
                If Entity = "All" Then
                    Selection.AutoFilter
                    Selection.AutoFilter Field:=4, Criteria1:="<>" 'column D
                Else
                    Selection.AutoFilter Field:=1, Criteria1:=Entity 'column A
                    Selection.AutoFilter Field:=4, Criteria1:="<>" 'column D
                End If

            
        
            End If
        
            'Entity1 Filter
                Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
                Set strDataRange = ActiveSheet.Range("A1:F" & Lastrow)
                Set keyRange = ActiveSheet.Range("F1")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
                strDataRange.Select
            If Entity = "All" Then
                Selection.AutoFilter
                Selection.AutoFilter Field:=4, Criteria1:="<>" 'column D
            Else
                Selection.AutoFilter Field:=1, Criteria1:=Entity 'column A
                Selection.AutoFilter Field:=4, Criteria1:="<>" 'column D
            End If
            
            Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
            Set strDataRange = ActiveSheet.Range("A1:F" & Lastrow)
            Set keyRange = ActiveSheet.Range("F1")
            strDataRange.SpecialCells(xlCellTypeVisible).Select
            Selection.Copy
            
            ActiveWorkbook.Sheets("TempSheet").Activate
            ActiveSheet.Range("A1").Select
            ActiveSheet.Paste
            
            Set strDataRange = ActiveSheet.Range("A2:F" & Lastrow)
            Set keyRange = ActiveSheet.Range("F1")
            Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
        
            If SFXChkBox.Value = True Then
                Selection.AutoFilter Field:=2, Criteria1:=Array(SFX1, SFX2, SFX3), Operator:=xlFilterValues
        
                strDataRange.SpecialCells(xlCellTypeVisible).Select
                Selection.EntireRow.Delete
                ActiveSheet.Range("A1").Select
                Selection.AutoFilter
            End If
            
            If OtherPidExcl.Value = True Then
                Application.Run "Module1.AdditionalExclusions"
            End If
            
            ActiveSheet.Range("A2:E41").Copy
            ActiveWorkbook.Sheets("Summary").Activate
            
            'Entity1 paste into Summary tab
            If RowNum = 2 Then
                Range("A198").PasteSpecial xlPasteValues
                Set strDataRange = ActiveSheet.Range("A198:F237")
                Set keyRange = ActiveSheet.Range("B197")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
            ElseIf RowNum = 3 Then
                Range("A106").PasteSpecial xlPasteValues
                Set strDataRange = ActiveSheet.Range("A106:F145")
                Set keyRange = ActiveSheet.Range("B105")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
            ElseIf RowNum = 4 Then
                Range("A152").PasteSpecial xlPasteValues
                Set strDataRange = ActiveSheet.Range("A152:F191")
                Set keyRange = ActiveSheet.Range("B151")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
            ElseIf RowNum = 6 Then
                Range("A244").PasteSpecial xlPasteValues
                Set strDataRange = ActiveSheet.Range("A244:F283")
                Set keyRange = ActiveSheet.Range("B243")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
            ElseIf RowNum = 7 Then
                Range("A290").PasteSpecial xlPasteValues
                Set strDataRange = ActiveSheet.Range("A290:F329")
                Set keyRange = ActiveSheet.Range("B289")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
            End If
        
            ActiveWorkbook.Sheets("TempSheet").Range("A:L").Delete
            
            'Entity2 Filter
            ActiveWorkbook.Sheets(TabName).Activate
            Set strDataRange = ActiveSheet.Range("A1:F" & Lastrow)
            Set keyRange = ActiveSheet.Range("F1")
            strDataRange.Select
            Selection.AutoFilter
            If Entity2 = "All" Then
                Selection.AutoFilter
                Selection.AutoFilter Field:=4, Criteria1:="<>" 'column D
            Else
                Selection.AutoFilter Field:=1, Criteria1:=Entity2, Field:=4, Criteria1:="<>" 'column A
                Selection.AutoFilter Field:=4, Criteria1:="<>" 'column D
            End If
            
            Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
            Set strDataRange = ActiveSheet.Range("A1:F" & Lastrow)
            Set keyRange = ActiveSheet.Range("F1")
            strDataRange.SpecialCells(xlCellTypeVisible).Select
            Selection.Copy
            
            ActiveWorkbook.Sheets("TempSheet").Activate
            ActiveSheet.Range("A1").Select
            ActiveSheet.Paste
            
            Set strDataRange = ActiveSheet.Range("A2:F" & Lastrow)
            Set keyRange = ActiveSheet.Range("F1")
            Lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
        
            If SFXChkBox.Value = True Then
                Selection.AutoFilter Field:=2, Criteria1:=Array(SFX1, SFX2, SFX3), Operator:=xlFilterValues
        
                strDataRange.SpecialCells(xlCellTypeVisible).Select
                Selection.EntireRow.Delete
                ActiveSheet.Range("A1").Select
                Selection.AutoFilter
            End If
            
            If OtherPidExcl.Value = True Then
                Application.Run "Module1.AdditionalExclusions"
            End If
            
            ActiveSheet.Range("A2:E41").Copy
            ActiveWorkbook.Sheets("Summary").Activate
        
            'Entity2 paste into Summary tab
            If RowNum = 2 Then
                Range("H198").PasteSpecial xlPasteValues
                Set strDataRange = ActiveSheet.Range("H198:M237")
                Set keyRange = ActiveSheet.Range("I197")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
            ElseIf RowNum = 3 Then
                Range("H106").PasteSpecial xlPasteValues
                Set strDataRange = ActiveSheet.Range("H106:M145")
                Set keyRange = ActiveSheet.Range("I105")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
            ElseIf RowNum = 4 Then
                Range("H152").PasteSpecial xlPasteValues
                Set strDataRange = ActiveSheet.Range("H152:M191")
                Set keyRange = ActiveSheet.Range("I151")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
            ElseIf RowNum = 6 Then
                Range("H244").PasteSpecial xlPasteValues
                Set strDataRange = ActiveSheet.Range("H244:M283")
                Set keyRange = ActiveSheet.Range("I243")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
            ElseIf RowNum = 7 Then
                Range("H290").PasteSpecial xlPasteValues
                Set strDataRange = ActiveSheet.Range("H290:M329")
                Set keyRange = ActiveSheet.Range("I289")
                strDataRange.Sort Key1:=keyRange, Order1:=xlDescending
            End If
        
            ActiveWorkbook.Sheets("TempSheet").Range("A:L").Delete
            RowNum = RowNum + 1
            TabName = ActiveWorkbook.Sheets("Summary").Range("A1").Offset(RowNum, 0).Value
    
        End If
    
    
    
    Loop

ActiveWorkbook.Sheets("PID").Activate
ActiveSheet.Range("A1").Select
Selection.AutoFilter
ActiveSheet.Range("L2:M3").Copy
ActiveWorkbook.Sheets("Summary").Activate
ActiveSheet.Range("G2:H3").Select
ActiveSheet.Paste
ActiveSheet.Range("G2:H3").Select
MsgBox RnmdFiles & "Summary Updated Successfully", vbInformation, "Done!"

---------------------------------------------------------------------------------------------------------------------------------------
'4. RENAME PDF Files

Private Sub CommandButton1_Click()

    Dim MyFolder As String
    Dim MyFile As String
    Dim i As Long
    Dim MyOldFile As String
    Dim MyNewFile As String
    Dim InvName As String
    Dim CellNum As Integer
    Dim ReName As String
    
    MyFolder = ActiveWorkbook.Sheets("Macro").Range("B2")
    
    CellNum = 7
    i = 7
    
    InvName = ActiveWorkbook.Sheets("Macro").Range("G" & CellNum)
    
    ReName = ActiveWorkbook.Sheets("Macro").Range("I" & CellNum)
    
    MyFile = Dir(MyFolder & "\" & InvName & ".pdf")
    
    Do Until CellNum = 28
        MyFolder = ActiveWorkbook.Sheets("Macro").Range("B2")
        InvName = ActiveWorkbook.Sheets("Macro").Range("G" & CellNum)
        ReName = ActiveWorkbook.Sheets("Macro").Range("I" & CellNum)
        MyFile = Dir(MyFolder & "\" & InvName & ".pdf")
        MyOldFile = MyFolder & "\" & MyFile
        MyNewFile = MyFolder & "\" & ReName & ".pdf"
        On Error GoTo ErrorHandler
        Name MyOldFile As MyNewFile
        On Error GoTo ErrorHandler
        MyFile = Dir
        CellNum = CellNum + 1
    Loop
    
ErrorHandler:
   Exit Sub
   Resume Next

End Sub
---------------------------------------------------------------------------------------------------------------------------------------

'5. DERIVATIVES ADJUSTMENTS

Sub Run_All_Dervs()

Dim strDataRange As Range
Dim keyRange As Range


Dim ReportingEnt As String
Dim CounterEnt As String
Dim TabName As String
Dim ToNext As Integer
Dim TabNum As Integer
Dim FirstCell As Integer

Application.ScreenUpdating = False
FirstCell = 1017
ToNext = FirstCell
ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)

TabName = "O.1 - Master Derivatives Data"
TabNum = 0


Do Until (TabNum = 2)
ToNext = FirstCell
ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)


'USD
Do Until (ReportingEnt = "")
    ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
    CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)
    ActiveWorkbook.Sheets(TabName).Activate
    Lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Range("A1").Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=47, Criteria1:=ReportingEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=48, Criteria1:=CounterEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=49, Criteria1:="USD", Operator:=xlFilterValues 'Filter by CCY
    Set strDataRange = ActiveSheet.Range("A1:BB" & Lastrow)
    Set keyRange = ActiveSheet.Range("A1")
    strDataRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ActiveWorkbook.Sheets("Derivatives Adjs").Activate
    ActiveSheet.Range("B1").Select
    Selection.PasteSpecial xlPasteValues
    Application.Run "Module5.Export_Dervs_Adjs_NoBox"
    ToNext = ToNext + 1
    
Loop


'AUD
ToNext = FirstCell
ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)
Do Until (ReportingEnt = "")
    ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
    CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)
    ActiveWorkbook.Sheets(TabName).Activate
    Lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Range("A1").Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=47, Criteria1:=ReportingEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=48, Criteria1:=CounterEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=49, Criteria1:="AUD", Operator:=xlFilterValues 'Filter by CCY
    Set strDataRange = ActiveSheet.Range("A1:BB" & Lastrow)
    Set keyRange = ActiveSheet.Range("A1")
    strDataRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ActiveWorkbook.Sheets("Derivatives Adjs").Activate
    ActiveSheet.Range("B1").Select
    Selection.PasteSpecial xlPasteValues
    Application.Run "Module5.Export_Dervs_Adjs_NoBox"
    ToNext = ToNext + 1
Loop


'GBP
ToNext = FirstCell
ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)
Do Until (ReportingEnt = "")
    ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
    CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)
    ActiveWorkbook.Sheets(TabName).Activate
    Lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Range("A1").Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=47, Criteria1:=ReportingEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=48, Criteria1:=CounterEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=49, Criteria1:="GBP", Operator:=xlFilterValues 'Filter by CCY
    Set strDataRange = ActiveSheet.Range("A1:BB" & Lastrow)
    Set keyRange = ActiveSheet.Range("A1")
    strDataRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ActiveWorkbook.Sheets("Derivatives Adjs").Activate
    ActiveSheet.Range("B1").Select
    Selection.PasteSpecial xlPasteValues
    Application.Run "Module5.Export_Dervs_Adjs_NoBox"
    ToNext = ToNext + 1
Loop


'EUR
ToNext = FirstCell
ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)
Do Until (ReportingEnt = "")
    ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
    CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)
    ActiveWorkbook.Sheets(TabName).Activate
    Lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Range("A1").Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=47, Criteria1:=ReportingEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=48, Criteria1:=CounterEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=49, Criteria1:="EUR", Operator:=xlFilterValues 'Filter by CCY
    Set strDataRange = ActiveSheet.Range("A1:BB" & Lastrow)
    Set keyRange = ActiveSheet.Range("A1")
    strDataRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ActiveWorkbook.Sheets("Derivatives Adjs").Activate
    ActiveSheet.Range("B1").Select
    Selection.PasteSpecial xlPasteValues
    Application.Run "Module5.Export_Dervs_Adjs_NoBox"
    ToNext = ToNext + 1
Loop


'JPY
ToNext = FirstCell
ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)
Do Until (ReportingEnt = "")
    ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
    CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)
    ActiveWorkbook.Sheets(TabName).Activate
    Lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Range("A1").Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=47, Criteria1:=ReportingEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=48, Criteria1:=CounterEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=49, Criteria1:="JPY", Operator:=xlFilterValues 'Filter by CCY
    Set strDataRange = ActiveSheet.Range("A1:BB" & Lastrow)
    Set keyRange = ActiveSheet.Range("A1")
    strDataRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ActiveWorkbook.Sheets("Derivatives Adjs").Activate
    ActiveSheet.Range("B1").Select
    Selection.PasteSpecial xlPasteValues
    Application.Run "Module5.Export_Dervs_Adjs_NoBox"
    ToNext = ToNext + 1
Loop


'CHF
ToNext = FirstCell
ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)
Do Until (ReportingEnt = "")
    ReportingEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("N" & ToNext)
    CounterEnt = ActiveWorkbook.Sheets("Derivatives Adjs").Range("O" & ToNext)
    ActiveWorkbook.Sheets(TabName).Activate
    Lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Range("A1").Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=47, Criteria1:=ReportingEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=48, Criteria1:=CounterEnt, Operator:=xlFilterValues
    Selection.AutoFilter Field:=49, Criteria1:="CHF", Operator:=xlFilterValues 'Filter by CCY
    Set strDataRange = ActiveSheet.Range("A1:BB" & Lastrow)
    Set keyRange = ActiveSheet.Range("A1")
    strDataRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ActiveWorkbook.Sheets("Derivatives Adjs").Activate
    ActiveSheet.Range("B1").Select
    Selection.PasteSpecial xlPasteValues
    Application.Run "Module5.Export_Dervs_Adjs_NoBox"
    ToNext = ToNext + 1
Loop


TabName = "O.2 - Master Derivatives Data"
TabNum = TabNum + 1
Loop

MsgBox "Derivatives Adjustments Completed Successfully!", vbInformation, "Done!"


End Sub

Sub Export_Dervs_Adjs_NoBox()

'Inflow Adjustments
ActiveSheet.Range("$A$1103:$R$1143").AutoFilter Field:=8, Criteria1:=">0.01", Operator:=xlAnd
On Error Resume Next
ActiveSheet.Range("$A$1104:$R$1143").SpecialCells(xlCellTypeVisible).Copy
Dim wsSheet1 As Worksheet
On Error Resume Next
Set wsSheet = Sheets("Dervs Inflows - Temp")
On Error Resume Next
If wsSheet Is Nothing Then
    Sheets.Add.Name = "Dervs Inflows - Temp"
End If

ActiveWorkbook.Sheets("Dervs Inflows - Temp").Activate

If Range("A1") = "" Then
    Range("A1").Select
Else
    ActiveSheet.Range("A1").End(xlDown).Offset(1, 0).Select
End If

Selection.PasteSpecial xlPasteValues

'Outflow Adjustments
ActiveWorkbook.Sheets("Derivatives Adjs").Activate
ActiveSheet.Range("$A$1103:$R$1143").AutoFilter

ActiveSheet.Range("$A$1145:$R$1185").AutoFilter Field:=8, Criteria1:=">0.01", Operator:=xlAnd
ActiveSheet.Range("$A$1146:$R$1185").SpecialCells(xlCellTypeVisible).Copy
Dim wsSheet2 As Worksheet
On Error Resume Next
Set wsSheet1 = Sheets("Dervs Outflows - Temp")
On Error Resume Next
If wsSheet1 Is Nothing Then
    Sheets.Add.Name = "Dervs Outflows - Temp"
End If

ActiveWorkbook.Sheets("Dervs Outflows - Temp").Activate

If Range("A1") = "" Then
    Range("A1").Select
Else
    ActiveSheet.Range("A1").End(xlDown).Offset(1, 0).Select
End If

Selection.PasteSpecial xlPasteValues

ActiveWorkbook.Sheets("Derivatives Adjs").Activate
ActiveSheet.Range("$A$1145:$R$1185").AutoFilter

    Range("B1:BD1015").Select
    Selection.ClearContents
    Range("B2").Select
    Selection.End(xlUp).Select
    


End Sub
