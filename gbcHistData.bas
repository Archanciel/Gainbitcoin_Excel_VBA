Attribute VB_Name = "gbcHistData"
Option Explicit
Sub AddEntries()
Attribute AddEntries.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Macro2 Macro
'
' Touche de raccourci du clavier: Ctrl+q
'
    Dim startCell As Range
    
    Set startCell = Selection
    
    startCell.Offset(, 2).Activate
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "General"
    Selection.Replace What:=" BTC", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Cut startCell.Offset(, 4)
    Range("A1").CurrentRegion.Select
    Selection.RemoveDuplicates Columns:=1, Header:=xlYes
    SplitDateTime startCell.Offset(, 1)
    Range(startCell.Offset(, 2), startCell.Offset(, 4).End(xlDown)).Select
    Selection.Cut startCell.Offset(, 1)
    startCell.Select
End Sub
Sub SplitDateTime(startCell As Range)
    Dim cell As Range
    Dim timeCol As Range
    
    Range(startCell, startCell.End(xlDown)).Select
    
    For Each cell In Selection
        If IsDate(cell.Value) Then
            cell.Offset(, 1).Resize(, 2).Value _
                = Array(DateSerial(Year(cell.Value), Month(cell.Value), Day(cell.Value)), _
                        TimeSerial(Hour(cell.Value), Minute(cell.Value), Second(cell.Value)))
        End If
    Next
    
    Set timeCol = startCell.Offset(, 2)
    Range(timeCol, timeCol.End(xlDown)).Select
    Selection.NumberFormat = "[$-F400]h:mm:ss AM/PM"
End Sub
