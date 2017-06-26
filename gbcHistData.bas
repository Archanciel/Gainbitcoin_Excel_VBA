Attribute VB_Name = "gbcHistData"
Option Explicit
Sub AddEntries()
Attribute AddEntries.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' AddEntries Macro
'
' Touche de raccourci du clavier: Ctrl+q
'
    Dim startCell As Range
    Dim currentRow
    Dim titleRow
    Application.ScreenUpdating = False
    
    Set startCell = Selection
    currentRow = startCell.Row
    
    'pasting content copied by the Firefox TableTool2 plugin
    ActiveSheet.Paste startCell
    
    If currentRow = 1 Then
        'first time data is added to the worksheet --> title line must be kept, but emptied
        titleRow = 1
    Else
        Rows(currentRow).EntireRow.Delete
        titleRow = 0
        Set startCell = Selection 'as line was deleted, startCell must be reset
    End If
    
    startCell.Offset(, 2).Activate
    
    'removing BTC from earning amounts
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "General"
    Selection.Replace What:=" BTC", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    'cut & paste on col 4 to make room for raw date col splitting in date and time cols
    Selection.Cut startCell.Offset(, 4)
    
    'eliminating possible double entries
    Range("A1").CurrentRegion.Select
    Selection.RemoveDuplicates Columns:=1, Header:=xlYes
    
    'date/time splitting
    SplitDateTime startCell.Offset(titleRow, 1) 'titleRow used here to prevent string
                                                'of title to interfere with date/time
                                                'data type !
    Range(startCell.Offset(, 2), startCell.Offset(, 4).End(xlDown)).Select
    Selection.Cut startCell.Offset(, 1)
    
    startCell.Select

    'setting col titles
    If currentRow = 1 Then
        'first time data is added to the worksheet --> titles must be set
        startCell.Value = "NR"
        startCell.Offset(, 1).Value = "DATE"
        startCell.Offset(, 2).Value = "TIME"
        startCell.Offset(, 3).Value = "EARNED"
    End If
    
    Application.ScreenUpdating = True
End Sub
Private Sub SplitDateTime(startCell As Range)
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
