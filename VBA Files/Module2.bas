Attribute VB_Name = "Module2"
Sub GetUniqueCount(ByVal rowSize As Integer, _
                    ByVal rowCount As Integer)

Dim headerRange As Range
Dim colIndex As Integer
Dim constRowCount As Integer
Dim rowIndex As Integer
constRowCount = 26
'get data row headers from sheet2
Set headerRange = Range(ActiveCell.Offset(0, 1), _
                        ActiveCell.Offset(0, rowSize - 1))


'clear contents before copying header
ActiveWorkbook.Sheets("Sheet3").Activate
Range("A1").Select
Range(ActiveCell.Offset(0, 1), ActiveCell.Offset(constRowCount, rowSize - 1)).Clear

'copy header row from datasheet to sheet3
headerRange.Copy _
    Destination:=Sheets("Sheet3").Range("B1")

'apply countif for each value in the range
colIndex = 1
rowIndex = 1
While (colIndex < rowSize)
    While rowIndex <= constRowCount
        ActiveCell.Offset(rowIndex, colIndex).Value2 = ActiveCell.Offset(rowIndex, 0).Text & ":" & ActiveCell.Offset(0, colIndex).Text
        ActiveCell.Offset(rowIndex, colIndex).Value2 = Application.WorksheetFunction _
                                .CountIf("datarange", ActiveCell.Offset(rowIndex, 0).Text)
        rowIndex = rowIndex + 1
    Wend
colIndex = colIndex + 1
rowIndex = 1
Wend
End Sub

