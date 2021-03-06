Attribute VB_Name = "Module3"
'This is intended for Motif search
'inputSeq --> data provided by user as input to be searched
'inputRowSize --> length of input data (row), omitting '-' char
'inputRowCount --> number of input rows
'rowSet --> concatenate row data without '-'
'
'The execution should begin only when at least 2 characted in caps is entered.
'scanning will not proceed if input sequence length is greater than input data.
'
'If no search result is obtained, user will be asked to re-enter sequence to search.
'If incorrect data is provied, input form displayed again after proper message.

Sub MotifSearch()
Attribute MotifSearch.VB_Description = "This macro will search for a sequence provided by the user against each input row only. Data sheet will not be referred in this case. To run this macro use the shorcut key.\n\nNote: Save the workbook before running the script"
Attribute MotifSearch.VB_ProcData.VB_Invoke_Func = "m\n14"

Dim inputSeq As Variant
Dim regex As New RegExp
Dim expStr As String

Application.ScreenUpdating = False
'select current workbook
Workbooks("PatternMatcher.xlsm").Activate
inputSeq = Application.InputBox("Enter the sequence to search (A-Z). Use * as wildcard only")
'data sanity - only a-z is allowed
expStr = "^[A-Z*]{2,}$"
With regex
    .Global = True
    .MultiLine = True
    .IgnoreCase = False
    .Pattern = expStr
End With
'input form validation
If (inputSeq = vbFalse) Then
    'exit prog when clicked cancel
    Exit Sub
Else
    'data sanity check
    If (regex.Test(inputSeq)) Then
        'sanity check pass - start processing
        Call process(inputSeq)
    Else
        'sanity check fail - alert msg and exit
        MsgBox "Incorrect data provided. It should accept A to Z (caps) only and at least 2 characters."
        'Exit Sub
        MotifSearch
    End If
End If
End Sub

'actual processing of data should begin from here, after sanity check confirmed
Function process(ByVal inputSeq As String)

Dim inputRowSize As Integer
Dim inputRowCount As Integer
Dim rowIndex As Integer
Dim colIndex As Integer
Dim rowSet As String
'collection will store 'input name-match count' (key-value pair) for each input
Dim searchResult As New Scripting.Dictionary
Dim expStr As String
Dim regex As New RegExp
Dim match As Variant

'select input sheet and starting cell
ActiveWorkbook.Sheets("Sheet1").Activate
Range("A1").Select
rowIndex = 1
colIndex = 1
rowSet = vbNullString
expStr = vbNullString
'get row size & count
inputRowSize = Range(ActiveCell, ActiveCell.End(xlToRight)).Count
inputRowCount = Range(ActiveCell.Offset(1), ActiveCell.End(xlDown)).Count
'scan input data sheet
While Not IsEmpty(ActiveCell.Offset(rowIndex, colIndex)) And rowIndex <= inputRowCount
    While Not IsEmpty(ActiveCell.Offset(rowIndex, colIndex)) And colIndex <= inputRowSize
        'concat input row data eliminating '-'
        If Not ActiveCell.Offset(rowIndex, colIndex).Text = "-" Then
            rowSet = rowSet & ActiveCell.Offset(rowIndex, colIndex).Text
        End If
        colIndex = colIndex + 1
    Wend
    'search will start only when concat len is greater than input else move to next
    If Len(rowSet) >= Len(inputSeq) Then
        If InStr(inputSeq, "*") Then
            expStr = Replace(inputSeq, "*", "(.)")
        Else
            expStr = inputSeq
        End If
        With regex
            .Global = True
            .IgnoreCase = True
            .MultiLine = True
            .Pattern = expStr
        End With
        Set match = regex.Execute(rowSet)
        If match.Count > 0 Then
            'add to collection
            searchResult.Add ActiveCell.Offset(rowIndex, 0).Text, match.Count
        End If
    End If
    'reset all loop data
    rowSet = vbNullString
    colIndex = 1
    rowIndex = rowIndex + 1
Wend
'motif search result check
If searchResult.Count > 0 Then
    'create motif search result sheet
    Call createResultSheet(searchResult, inputSeq)
    MsgBox "Motif search complete!"
Else
    'generic alert msg
    MsgBox "No result found for Motif search: " & inputSeq
    redo = MsgBox("Do you want search again?", vbYesNo + vbQuestion)
    If redo = vbYes Then
        'redo search - call motifsearch subroutine
        Call MotifSearch
    End If
End If
End Function

'create search result sheet based on input name and count of occurrence
Function createResultSheet(ByRef searchResult As Scripting.Dictionary, _
                            ByRef inputSeq As String)
Dim sheetName As String
Dim colIndex As Integer
Dim rowIndex As Integer

'header row count is 2, hence result should start from 3
colIndex = 0
rowIndex = 2
'avoid wild character in sheet name
If InStr(inputSeq, "*") > 0 Then
    inputSeq = Replace(inputSeq, "*", "-")
End If
sheetName = Left("Motif_ " & inputSeq, 30)
'avoid duplicate sheet name, remove existing and create new one
For Each ws In ActiveWorkbook.Worksheets
    If InStr(ws.Name, sheetName) > 0 Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Sheets(ws.Name).Delete
    End If
Next
'create new sheet for motif search result
With ActiveWorkbook
    Worksheets.Add(after:=Worksheets(Worksheets.Count)) _
            .Name = sheetName
End With
ActiveWorkbook.Sheets(sheetName).Activate
ActiveWindow.Zoom = 74
Range("A1").Select
ActiveCell.Value2 = "Motif:"
ActiveCell.Offset(0, 1) = inputSeq
With ActiveCell.Offset(0, 1)
    .Font.Bold = True
    .Interior.Color = RGB(200, 220, 0)
End With
ActiveCell.Offset(1, 0).Value2 = "Name"
ActiveCell.Offset(1, 1).Value2 = "Count"
'start filling result from 3rd row
For Each Key In searchResult.Keys
    ActiveCell.Offset(rowIndex, colIndex).Value2 = Key
    ActiveCell.Offset(rowIndex, colIndex + 1).Value2 = searchResult(Key)
    rowIndex = rowIndex + 1
Next Key
End Function
