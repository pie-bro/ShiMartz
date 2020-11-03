Sub Data_Clean_GreenTree()
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
Cells.Find(What:="Account No.").Select
ActiveCell.Offset(1, 0).Select
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
ActiveCell.Value = "Month"
ActiveCell.Offset(0, 1).Value = "Farm"
ActiveCell.Offset(0, 2).Value = "Code"
ActiveCell.Offset(0, 3).Value = "GL Number"
ActiveCell.Offset(0, 4).Value = "AA"
ActiveCell.Offset(0, 5).Value = "GL Name"
ActiveCell.Offset(0, 6).Value = "Amount"
'ActiveCell.Offset(1, 0).Select
Cells.Find(What:="GL Name").Offset(1, 0).Select

'Add GL NAME
j = ActiveCell.Column
'Range("M1").Value = i
ActiveCell.FormulaR1C1 = _
    "=IF(ISTEXT(RC[7]),RC[7],R[-1]C)"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))
ActiveCell.Offset(1, 1).Select

'Add Amount
Cells.Find(What:="Amount").Offset(1, 0).Select
j = ActiveCell.Column
    ActiveCell.Formula2R1C1 = "=IF(ISBLANK(RC[5]),"""", RC[13]-RC[14])"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))

'Add GL Number
Cells.Find(What:="GL Number").Offset(1, 0).Select
j = ActiveCell.Column
    ActiveCell.Formula2R1C1 = "=IF(ISBLANK(RC[8]),RC[4],R[-1]C)"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))

'Add Farm
Cells.Find(What:="Farm").Offset(1, 0).Select
j = ActiveCell.Column
    ActiveCell.Formula2R1C1 = "=MID(RC[2],4,3)"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))


'Add Code
Cells.Find(What:="Code").Offset(1, 0).Select
j = ActiveCell.Column
    ActiveCell.Formula2R1C1 = "=RIGHT(RC[1],4)"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))

'Add AA
    Cells.Find(What:="AA").Offset(1, 0).Select
'    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
'    ActiveCell.Value = "AA"
'    ActiveCell.Offset(2, 0).Select
    j = ActiveCell.Column
'    if AA code more than 10000, change the R10000 to Rx, x is the new last row number of AA code
    ActiveCell.FormulaR1C1 = _
         "=IF(ISBLANK(RC[3]),"""",INDEX('AA'!R1C2:R10000C2,MATCH('GL'!RC[-1],'AA'!R1C1:R10000C1,0)))"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))

'Add Month
Cells.Find(What:="Month").Offset(1, 0).Select
j = ActiveCell.Column
    ActiveCell.Formula2R1C1 = "=IF(ISBLANK(RC[11]),"""",MONTH(RC[11]))"
Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))


'Add Fincial Year
    Cells.Find(What:="Amount").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    j = ActiveCell.Column
    ActiveCell.Value = "FY"
    ActiveCell.Offset(1, 0).Select
      ActiveCell.Formula2R1C1 = _
         "=IF(ISBLANK(RC[6]),"""",IF(RC[6]<=DATE(YEAR(RC[6]),Programme!R2C2,Programme!R3C2),YEAR(RC[6]),YEAR(RC[6])+1))"
'        "=IF(LEFT(CELL(""format"",RC[2]))=""D"", IF(RC[2]<=DATE(YEAR(RC[2]),Programme!R2C2,Programme!R3C2),YEAR(RC[2]),YEAR(RC[2])+1),"""")"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(i, j))
    ActiveCell.Offset(0, 1).Select


End Sub
