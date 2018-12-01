Attribute VB_Name = "Module1"
Public inputList As String
Public inputName As String
Public matchedRowNumbers As String          'row number in the database that matches with input
Public matchedRowNumberScores As String     'score of each row that qualifies for output
Public matchedRowNumberWeightage As String  'weightage of each row that qualifies for output
Public matchedPositionList As String        'position matching in each row
Public matchedCountOfRow As String          'count of position matching in each row
Sub PatternChecker()
Attribute PatternChecker.VB_Description = "Macro will run once shorcut key is pressed.\nNote: Save the workbook before running the script"
Attribute PatternChecker.VB_ProcData.VB_Invoke_Func = "r\n14"
Dim inputRowSize As Integer
Dim inputRowCount As Integer
Dim rowSize As Integer
Dim rowCount As Integer
Dim inputStrings() As String
Dim rowIndex As Integer
Dim dbString As String
Dim inputString As Variant
Dim inputIndex As Integer
Dim sheetName As String

Application.ScreenUpdating = False
Workbooks("PatternMatcher.xlsm").Activate
ActiveWorkbook.Sheets("Sheet2").Activate
Range("A1").Select

'get the input string from sheet1 and storing them in an array after concating
ActiveWorkbook.Sheets("Sheet1").Activate
ActiveSheet.Range("A1").Select
inputRowSize = Range(ActiveCell, ActiveCell.End(xlToRight)).Count
inputRowCount = Range(ActiveCell.Offset(1), ActiveCell.End(xlDown)).Count
Call concatRowValues(inputRowSize, inputRowCount, True, 0)
inputStrings = Split(inputList, ",")

'selecting the data sheet
ActiveWorkbook.Sheets("Sheet2").Activate
ActiveSheet.Range("A1").Select
rowSize = Range(ActiveCell, ActiveCell.End(xlToRight)).Count
rowCount = Range(ActiveCell.Offset(1), ActiveCell.End(xlDown)).Count

'TODO: add a check for the difference in row size of input an db
If inputRowSize <> rowSize Then
    MsgBox "Input row size & dabatase row size do not match. Execution stopped!", vbOKOnly
    End
End If

'begin comparison process and create sheet if there is a match
rowIndex = 1
inputIndex = 1
For Each inputString In inputStrings
    While rowIndex <= rowCount
        dbString = concatRowValues(rowSize, rowCount, False, rowIndex)
        Call compareData(inputString, dbString, rowSize, rowIndex)
        rowIndex = rowIndex + 1
    Wend
    If Len(matchedRowNumbers) > 0 And Len(matchedRowNumberScores) > 0 Then
        'create a worksheet for current input string
        sheetName = Split(inputName, ",")(inputIndex - 1)
        Call createWorkbook(inputIndex, rowSize, sheetName)
    End If
    inputIndex = inputIndex + 1
    rowIndex = 1
    matchedRowNumbers = ""
    matchedRowNumberScores = ""
    matchedRowNumberWeightage = ""
    matchedCountOfRow = ""
    matchedPositionList = ""
Next inputString
inputList = ""
inputName = ""
'not yet implemented
'Call GetUniqueCount(rowSize, rowCount)
Application.ScreenUpdating = True
MsgBox "Complete.", vbOKOnly
End Sub
'create a new sheet and copy the desired rows
Function createWorkbook(ByVal inputIndex As Integer, _
                        ByVal rowSize As Integer, _
                        ByVal sheetName As String)
Dim rowNumbers() As String
Dim rowScores() As String
Dim rowWeightage() As String
Dim rowMatchCount() As String
Dim rowPositions() As String
Dim addr As String
Dim dataRanges As Range
Dim rankRanges As Range
Dim rng As Range

rowNumbers = Split(matchedRowNumbers, ",")
rowScores = Split(matchedRowNumberScores, ",")
rowWeightages = Split(matchedRowNumberWeightage, ",")
rowMatchCount = Split(matchedCountOfRow, ",")
rowPositions = Split(matchedPositionList, "|")
If UBound(rowNumbers) <= 0 Then
    rowNumbers(0) = matchedRowNumbers
    rowScores(0) = matchedRowNumberScores
    rowWeightages(0) = matchedRowNumberWeightage
    rowMatchCount(0) = matchedCountOfRow
    rowPositions(0) = matchedPositionList
End If
addr = Split(ActiveCell.End(xlToRight).Address, "$")(1)
Set dataRanges = Range("A1:" & addr & "1")
For i = LBound(rowNumbers) To UBound(rowNumbers)
    x = CInt(rowNumbers(i)) + 1
    Set dataRanges = Union(dataRanges, Range("A" & x & ":" & addr & x))
Next i
With ActiveWorkbook
    Worksheets.Add(after:=Worksheets(Worksheets.Count)) _
                .Name = sheetName
    Sheets("Sheet2").Activate
End With
'copy selected to output sheet
dataRanges.Copy _
    Destination:=Sheets(sheetName).Range("A1")
ActiveWorkbook.Sheets(Worksheets.Count).Activate
ActiveWindow.Zoom = 74
ActiveCell.Offset(0, rowSize).Value2 = "Weightage"
ActiveCell.Offset(0, rowSize + 1).Value2 = "MatchCount"
ActiveCell.Offset(0, rowSize + 2).Value2 = "%age"
'assigning score/weightage of each row
i = 0
While i <= UBound(rowScores)
    ActiveCell.Offset(i + 1, rowSize).Value2 = rowWeightages(i)
    ActiveCell.Offset(i + 1, rowSize + 1).Value2 = rowMatchCount(i)
    ActiveCell.Offset(i + 1, rowSize + 2).Value2 = rowScores(i)
    ActiveCell.Offset(i + 1, rowSize + 2).NumberFormat = "0.00"
    i = i + 1
Wend
'highlight matching cells of the row
rowIndex = 0
While rowIndex <= UBound(rowPositions)
    cols = Split(rowPositions(rowIndex), ",")
    colIndex = 0
    While colIndex <= UBound(cols)
        ActiveCell.Offset(rowIndex + 1, cols(colIndex)).Interior.Color = RGB(0, 255, 0)
        colIndex = colIndex + 1
    Wend
    rowIndex = rowIndex + 1
Wend
'sorting the rows based on scores
If ActiveCell.End(xlDown).Row - 1 > 2 Then
    Set rng = Range(ActiveCell, ActiveCell.Offset(ActiveCell.End(xlDown).Row - 1, rowSize + 2))
    Set sortRange = Range(ActiveCell.Offset(0, rowSize + 2), _
                    ActiveCell.Offset(ActiveCell.End(xlDown).Row - 1, rowSize + 2))
    rng.Sort key1:=sortRange, _
                order1:=xlDescending, _
                Header:=xlYes
End If
'assigning rank to according to each row
ActiveCell.Offset(0, rowSize + 3).Value2 = "Rank"
Set rankRanges = Range(ActiveCell.Offset(1, rowSize + 2), _
                    ActiveCell.Offset(ActiveCell.End(xlDown).Row - 1, rowSize + 2))
i = 0
While i <= UBound(rowScores)
    ActiveCell.Offset(i + 1, rowSize + 3).Value2 = Application.WorksheetFunction. _
            Rank(ActiveCell.Offset(i + 1, rowSize + 2).Value2, rankRanges)
    i = i + 1
Wend

ActiveWorkbook.Sheets("Sheet2").Activate
End Function
'match input string with the database string
Function compareData(ByVal inputString As String, _
                    ByVal dataString As String, _
                    ByVal stringLength As Integer, _
                    ByVal rownum As Integer)
Dim pos As Integer
Dim matchCount As Integer
Dim matchScore As Double
Dim dbString As String
Dim rowWeightage As Integer
Dim positionList As String

'extract the string & weightage
dbString = Split(dataString, "|")(0)
rowWeightage = Split(dataString, "|")(1)
positionList = ""
pos = 1
matchCount = 0
While pos <= stringLength - 1
    'skip the position in the input string if the value is '-'
    If Mid(UCase(inputString), pos, 1) <> "-" _
        And Mid(UCase(inputString), pos, 1) = Mid(UCase(dbString), pos, 1) Then
        matchCount = matchCount + 1
        positionList = pos & "," & positionList
    End If
    pos = pos + 1
Wend
If matchCount > 0 Then
    positionList = Left(positionList, Len(positionList) - 1)    'eliminating the end comma in the string
End If
'score calculation will depend on row weightage
matchScore = matchCount / rowWeightage * 100
If matchScore > 20 And matchScore <= 100 Then
    If Len(matchedRowNumbers) <= 0 _
        And Len(matchedRowNumberScores) <= 0 _
        And Len(matchedRowNumberWeightage) <= 0 _
        And Len(matchedCountOfRow) <= 0 _
        And Len(matchedPositionList) <= 0 Then
        
        matchedRowNumbers = rownum
        matchedRowNumberScores = matchScore
        matchedRowNumberWeightage = rowWeightage
        matchedCountOfRow = matchCount
        matchedPositionList = positionList
    Else
        matchedRowNumbers = matchedRowNumbers & "," & rownum
        matchedRowNumberScores = matchedRowNumberScores & "," & matchScore
        matchedRowNumberWeightage = matchedRowNumberWeightage & "," & rowWeightage
        matchedCountOfRow = matchedCountOfRow & "," & matchCount
        matchedPositionList = matchedPositionList & "|" & positionList
    End If
End If
End Function
'Concatenating cells for each row
Function concatRowValues(ByVal rowSize As Integer, _
                        ByVal rowCount As Integer, _
                        ByVal flag As Boolean, _
                        ByVal rownum As Integer) As String
Dim colIndex As Integer
Dim rowIndex As Integer
Dim rowSet As String
Dim retFlag As Boolean
Dim rowWeightage As Integer

rowWeightage = 0
retFlag = False
colIndex = 1    'name column should be omitted
'if flag is off - use rowNum as the row to be concated
If Not flag Then
    rowIndex = rownum
Else
    rowIndex = 1
End If
Do While Not IsEmpty(ActiveCell.Offset(rowIndex, colIndex)) And rowIndex <= rowCount
    While Not IsEmpty(ActiveCell.Offset(rowIndex, colIndex)) And colIndex < rowSize
        If Len(rowSet) <= 0 Then
            rowSet = ActiveCell.Offset(rowIndex, colIndex).Text
        Else
            rowSet = rowSet & ActiveCell.Offset(rowIndex, colIndex).Text
        End If
        'calculate row weightage based on number of '-' for data row
        If Not flag And ActiveCell.Offset(rowIndex, colIndex).Text <> "-" Then
            rowWeightage = rowWeightage + 1
        End If
        colIndex = colIndex + 1
    Wend
    If (flag) Then
        If Len(inputList) <= 0 And Len(inputName) <= 0 Then
            inputList = rowSet
            inputName = ActiveCell.Offset(rowIndex).Text
        Else
            inputList = inputList & "," & rowSet
            inputName = inputName & "," & ActiveCell.Offset(rowIndex).Text
        End If
    Else
        'return the concat string and exit loop for each data row
        retFlag = True
        Exit Do
    End If
    rowSet = ""
    colIndex = 1
    rowIndex = rowIndex + 1
Loop
If retFlag Then
    concatRowValues = rowSet & "|" & rowWeightage
End If
End Function
