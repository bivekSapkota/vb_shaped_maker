Attribute VB_Name = "Gas_Shaped_maker"
Sub finding_data_from_DCS()
Attribute finding_data_from_DCS.VB_ProcData.VB_Invoke_Func = "J\n14"
' script written by Bivek Sapkota
' to find data from dcs for creating shaped file

Dim volume() As Double
Dim startDate As Date
Dim Deal_Id As Integer

bookName = ActiveWorkbook.Name
bookPath = ActiveWorkbook.Path
Worksheets(1).Activate
Row = 1
Column = 1
Do While Cells(Row, Column) <> "" ' reading column nos from the original sheet
If Cells(Row, Column) = "ATCO Transaction Number" Then ATrNo = Column
If Cells(Row, Column) = "ATCO Strategy Number" Then AStNo = Column
If Cells(Row, Column) = "Commodity" Then comdty = Column
If Cells(Row, Column) = "Volume" Then vol = Column
If Cells(Row, Column) = "Start Date" Then Sdate = Column
If Cells(Row, Column) = "End Date" Then EDate = Column
Column = Column + 1
Loop
datarow = Row + 1

Deal_Id = InputBox("Enter the Reference id")
Do While Cells(datarow, ATrNo) <> Deal_Id  'finding the starting of the shaped volume for the reference id
    datarow = datarow + 1
Loop

dealStartRow = datarow   'starting row of the required deal
startDate = Cells(dealStartRow, Sdate).Value 'storing the start date of the deal
Do While Cells(datarow, ATrNo) = Deal_Id
datarow = datarow + 1
Loop
dealendrow = datarow - 1 ' ending row of the required deal
endDate = Cells(dealendrow, EDate) ' storing the end date of the deal

If Cells(datarow, comdty) = "gas" Then  ' finding out if it is a gas deal or power deal
MsgBox ("This is a gas deal")
End If

ReDim volume(dealStartRow To dealendrow) As Double  'declaring array for storing volume

For loopDealRow = dealStartRow To dealendrow
volume(loopDealRow) = Cells(loopDealRow, vol)
Next loopDealRow

'End of finding volume from the dcs file

'now adding new workbook
Workbooks.Add
Sheets("Sheet1").Select
Sheets("Sheet1").Name = Deal_Id

'shaped making starts now---------------------------------------------------------------------------------------------------------------

'shaped auto filler

Dim loopDate As Date

deal_prefix = InputBox("Select prefix for Deal ID")
Range("a1") = "Deal_id"
Range("b1") = "Term_date"
Range("c1") = "Hour"
Range("d1") = "is_dst"
Range("e1") = "Volume"
Range("f1") = "Price"
Range("g1") = "Leg"

Row = 2
Column = 1

loopmonth = Month(startDate) 'initial value for the looping month
For loopDate = startDate To endDate
If Month(loopDate) > loopmonth Then dealStartRow = dealStartRow + 1
    For hours = 0 To 23
        Cells(Row, Column).Value = UCase(deal_prefix) & "_" & Deal_Id
        Cells(Row, Column + 1) = loopDate
        Cells(Row, Column + 2) = hours
        Cells(Row, Column + 3).Value = 0
        If hours = 0 Then Cells(Row, Column + 4).Value = volume(dealStartRow)
        Cells(Row, Column + 5).Value = "NULL"
        Cells(Row, Column + 6).Value = 1
        Row = Row + 1
    Next hours
loopmonth = Month(loopDate)
Next loopDate

'simple formatting for date column
Columns("B:B").EntireColumn.AutoFit
Columns("B:B").NumberFormat = "m/d/yyyy"

'macro magic to eliminate trailing commas while saving in csv
'making TRMTracker import compatible
'script written by Bivek Sapkota


    Range("a1").Select
    Selection.End(xlToRight).Offset(0, 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("a1").Select
    Selection.End(xlDown).Offset(1, 0).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("a1").Select
    
    
'trailing commas now removed

'now saving the file as csv in the current folder
    ActiveWorkbook.SaveAs Filename:=bookPath & "\" & Deal_Id, FileFormat:=xlCSV, CreateBackup:=False

Windows(bookName).Activate ' now returning to the original workbook
End Sub

