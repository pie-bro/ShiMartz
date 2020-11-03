Sub Data_Clean_Xero()
'
' Data Clean
'Final version

'

'Copy GL Code to another sheet and change the sheet name to "AA"
'Change the first report name to "GL"

Sheets("GL").Select
Dim i As Long
Dim j As Long
'Dim t As String


i = Range("A1").SpecialCells(xlLastCell).Row
'Range("A884").Value = i
t = Range("a1").Value
'Range("o1").Value = t


'Add column and change name
Cells.Find(What:="Date").Select
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
ActiveCell.Value = "GL NAME"
ActiveCell.Offset(0, 1).Value = "Month"
ActiveCell.Offset(0, 2).Value = "Amount"
ActiveCell.Offset(2, 0).Select


'Add GL NAME
j = ActiveCell.Column
'Range("M1").Value = i
ActiveCell.FormulaR1C1 = _
    "=IF(ISTEXT(RC[3]),RC[3],IF(ISNUMBER(RC[3]),R[-1]C,""""))"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))
ActiveCell.Offset(1, 1).Select

'Add month
j = ActiveCell.Column
    ActiveCell.Formula2R1C1 = "=IF(LEFT(CELL(""format"",RC[2]))=""D"",MONTH(RC[2]),"""")"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))
ActiveCell.Offset(0, 1).Select

'Add amount
j = ActiveCell.Column
    ActiveCell.Formula2R1C1 = "=IF(LEFT(CELL(""format"",RC[1]))=""D"",RC[5]-RC[6],"""")"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))

'Add AA
    Cells.Find(What:="Month").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.Value = "AA"
    ActiveCell.Offset(2, 0).Select
    j = ActiveCell.Column
'    if AA code more than 10000, change the R10000 to Rx, x is the new last row number of AA code
    ActiveCell.FormulaR1C1 = _
         "=IF(ISBLANK(RC[3]),"""",INDEX('AA'!R1C2:R10000C2,MATCH('GL'!RC[-1],'AA'!R1C1:R10000C1,0)))"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))
    
'Add Fincial Year
    Cells.Find(What:="Amount").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    j = ActiveCell.Column
    ActiveCell.Value = "Financial Year"
    ActiveCell.Offset(3, 0).Select
      ActiveCell.Formula2R1C1 = _
        "=IF(LEFT(CELL(""format"",RC[2]))=""D"", IF(RC[2]<=DATE(YEAR(RC[2]),Programme!R2C2,Programme!R3C2),YEAR(RC[2]),YEAR(RC[2])+1),"""")"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))
    ActiveCell.Offset(0, 1).Select


End Sub











Previous Sub Data_Clean()
'
' Data Clean
'

'

'Copy GL Code to another sheet and change the sheet name to "AA"
'Change the first report name to "GL"

Sheets("GL").Select
Dim i As Long
Dim j As Long
'Dim t As String


i = Range("A1").SpecialCells(xlLastCell).Row
'Range("A884").Value = i
t = Range("a1").Value
'Range("o1").Value = t


'Add column and change name
Cells.Find(What:="Date").Select
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
ActiveCell.Value = "GL NAME"
ActiveCell.Offset(0, 1).Value = "Month"
ActiveCell.Offset(0, 2).Value = "Amount"
ActiveCell.Offset(2, 0).Select


'Add GL NAME
j = ActiveCell.Column
'Range("M1").Value = i
ActiveCell.FormulaR1C1 = _
    "=IF(ISTEXT(RC[3]),RC[3],IF(ISNUMBER(RC[3]),R[-1]C,""""))"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))
ActiveCell.Offset(1, 1).Select

'Add month
j = ActiveCell.Column
    ActiveCell.Formula2R1C1 = "=IF(LEFT(CELL(""format"",RC[2]))=""D"",MONTH(RC[2]),"""")"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))
ActiveCell.Offset(0, 1).Select

'Add amount
j = ActiveCell.Column
    ActiveCell.Formula2R1C1 = "=IF(LEFT(CELL(""format"",RC[1]))=""D"",RC[5]-RC[6],"""")"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))

'Add AA
    Cells.Find(What:="Month").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.Value = "AA"
    ActiveCell.Offset(2, 0).Select
    j = ActiveCell.Column
'    if AA code more than 300, change the R300 to Rx, x is the new last row number of AA code
    ActiveCell.FormulaR1C1 = _
         "=IF(ISBLANK(RC[3]),"""",INDEX('AA'!R1C2:R300C2,MATCH('GL'!RC[-1],'AA'!R1C1:R300C1,0)))"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))


End Sub


Financial Year
'Add Fincial Year
j = ActiveCell.Column
      ActiveCell.Formula2R1C1 = _
        "=IF(LEFT(CELL(""format"",RC[2]))=""D"", IF(RC[2]<=DATE(YEAR(RC[2]),Programme!R2C2,Programme!R3C2),YEAR(RC[2]),YEAR(RC[2])+1),"""")"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))
ActiveCell.Offset(0, 1).Select

Financial Year Calculation
=IF(LEFT(CELL("format",E14))="D", IF(E14<=DATE(YEAR(E14),$Q$9,$R$9),YEAR(E14),YEAR(E14)+1),"")


=IF(LEFT(CELL("format",A14))="D",
 IF(AND(A14<=L10,A14>EDATE($L$10,-12)),YEAR($L$10),FALSE),"")




VLOOK UP
=IF(ISBLANK(E8),"",INDEX(Sheet5!$B$1:$B$52,MATCH(A8,Sheet5!$A$1:$A$52,0)))






Private Sub CommandButton1_Click()

' Data Clean
'

'

'Copy GL Code to another sheet and change the sheet name to "AA"
'Change the first report name to "GL"

'ActiveSheet.Shapes.Range(Array("Button 1")).Select
'Selection.OnAction = "Data_Clean"
'ActiveSheet.Shapes.Range(Array("Button 1")).OnAction = "Data_Clean"

Sheets("GL").Select

'Worksheets(1).Select
Dim i As Long
Dim j As Long
'Dim z As Long
'Dim t As String
Dim a As String


i = Range("A1").SpecialCells(xlLastCell).Row

'j = ActiveCell.Column

'Range("Q1").Value = i
'Range("Q2").Value = ActiveCell.Value


'Range("A884").Value = i
t = Range("a1").Value
'Range("o1").Value = t


'Add column and change name
'Worksheets(1).Range("A1").Select
'Set result = Worksheets(1).Range("A:A").Find(What:="Date")

ActiveSheet.Range("A:A").Find(What:="Date").Select
'Cells(1, "A").Select

'Set result = Range("A:A").Find("Date")




'Cells.Find(What:="Date").Select
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
ActiveCell.Value = "GL NAME"
ActiveCell.Offset(0, 1).Value = "Month"
ActiveCell.Offset(0, 2).Value = "Amount"
ActiveCell.Offset(2, 0).Select


'Add GL NAME
j = ActiveCell.Column

a = Cells(i, j).Address

'Range("M1").Value = i
ActiveCell.FormulaR1C1 = _
    "=IF(ISTEXT(RC[3]),RC[3],IF(ISNUMBER(RC[3]),R[-1]C,""""))"
Selection.AutoFill Destination:=Range(ActiveCell.Address, a)
ActiveCell.Offset(1, 1).Select

'Add month
j = ActiveCell.Column
    ActiveCell.Formula2R1C1 = "=IF(LEFT(CELL(""format"",RC[2]))=""D"",MONTH(RC[2]),"""")"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))
ActiveCell.Offset(0, 1).Select

'Add amount
j = ActiveCell.Column
    ActiveCell.Formula2R1C1 = "=IF(LEFT(CELL(""format"",RC[1]))=""D"",RC[5]-RC[6],"""")"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))

'Add AA
    Cells.Find(What:="Month").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.Value = "AA"
    ActiveCell.Offset(2, 0).Select
    j = ActiveCell.Column
'    if AA code more than 300, change the R300 to Rx, x is the new last row number of AA code
    ActiveCell.FormulaR1C1 = _
         "=IF(ISBLANK(RC[3]),"""",INDEX('AA'!R1C2:R300C2,MATCH('GL'!RC[-1],'AA'!R1C1:R300C1,0)))"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))


End Sub












Sub Macro1()
'
' Macro1 Macro


Dim i As Integer
Dim j As Integer

'For i = 1 To 20
'
'If Cells(i, "A").Value = "Date" Then
'Range("Ai").Select
'i = Range("Ai").Row
'j = Range("Ai").Column

'Next i

'Range("ij").Select

j = Range("A1").SpecialCells(xlLastCell).Row


Range("A6").Select
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Range("A6").Value = "GL NAME"
Range("B6").Value = "Month"
Range("C6").Value = "Amount"
Range("A8").Select

For i = 8 To j

If Application.WorksheetFunction.IsNumber(Cells(i, "D").Value) Then
Cells(i, "A").Value = Cells(i - 1, "A")

ElseIf IsDate(Cells(i, "D").Value) Then
Cells(i, "A").Value = Cells(i - 1, "A")


ElseIf Application.WorksheetFunction.IsText(Cells(i, "D").Value) Then
Cells(i, "A").Value = Cells(i, "D")


End If

Cells(i, "B").Select

If IsDate(Cells(i, "D").Value) Then
Cells(i, "B").Value = Month(Cells(i, "D"))
End If


Cells(i, "C").Select

If Application.WorksheetFunction.IsNumber(Cells(i, "D").Value) Then
Cells(i, "C").Value = Format((Cells(i, "H").Value - Cells(i, "I").Value), Cells(i, "H").NumberFormat)

ElseIf IsDate(Cells(i, "D").Value) Then
Cells(i, "C").Value = Format((Cells(i, "H").Value - Cells(i, "I").Value), Cells(i, "H").NumberFormat)


End If


Next i

    
End Sub












Sub Data_Cleaning1()
'
'Prepare GL NAME, Month, Amount

'*Prepare copy first for the original data in case Macro not working good

'1.DATE must at A6, if not, move the table to A6

'2.IF the last row number over 900, change 900 to the correct row number
'Instruction :Click CTRL+H, in the pop up window, write 900 in "Find What:",
'Write the correct row number in "Replace with",
'Then Click "Replace All"

'3.Check Date, if filter date not right, change year
'
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.AutoFilter
    Range("A6").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "GL Name"
    Range("A8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISTEXT(RC[1]),RC[1],IF(ISNUMBER(RC[1]),R[-1]C,""""))"
    Range("A8").Select
    Selection.AutoFill Destination:=Range("A8:A900")
    Range("A8:A900").Select

    Range("B6").Select
    ActiveSheet.Range("$B$6:$H$900").AutoFilter Field:=1, Operator:= _
        xlFilterValues, Criteria2:=Array(0, "12/31/2020", 0, "12/31/2019")
'    Change Date Above
    Range("B4").Activate
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C4").Activate
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "Month"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "Amount"
    Range("B9").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=MONTH(RC[2])"
    Range("B9").Select
    Selection.Copy
    Range("B9:B900").Select
    ActiveSheet.Paste
    Range("C9").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[5]-RC[6]"
    Range("C9").Select
    Selection.Copy
    Range("C9:C900").Select
    ActiveSheet.Paste
End Sub
