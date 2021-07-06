Attribute VB_Name = "Building_Adjacency_Matrix"
Option Explicit
Public ImportBook, ExportBook As Workbook
Public ImportSheet, ExportSheet1, ExportSheet2, ExportSheet3  As Worksheet
Public LastRow, TempRow, TempLastRow, LastColumn, TempLastColumn, SelectColumn, r, c, i, i2 As Long
Public SelectCell, TempCell As Range
Public Arr(), RowArr(), SplitArr As Variant
Public OriginalAutofilter, MultipleData As Boolean



Sub Recover_from_errors()

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub



Sub Conversion_from_Individuals_to_Groups()

Application.Dialogs(xlDialogActivate).Show
Set ImportBook = ActiveWorkbook
If ImportBook.FullName = ThisWorkbook.FullName Then GoTo DoNotStart
Set ImportSheet = ActiveSheet
OriginalAutofilter = ImportSheet.AutoFilterMode
ImportSheet.AutoFilterMode = False

Set SelectCell = Nothing
On Error Resume Next
Set SelectCell = Application.InputBox(prompt:="Click a cell in the column of intended category and click the OK button. Any cell in the intended column will do.", Type:=8)
If SelectCell Is Nothing Then
    On Error GoTo 0
    GoTo EndInTheMiddle
End If
On Error GoTo 0
SelectColumn = SelectCell.Column

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Set ExportBook = Workbooks.Add
Set ExportSheet1 = ExportBook.Worksheets(1)
Set ExportSheet2 = Worksheets.Add(After:=ExportSheet1)
Set ExportSheet3 = Worksheets.Add(After:=ExportSheet2)

With ExportSheet2
    ImportSheet.UsedRange.Copy .Cells(1, 1)
    Arr() = .UsedRange.Value
    .UsedRange.NumberFormatLocal = "0_ "
    .UsedRange.Value = Arr()
End With
ImportSheet.Columns(SelectColumn).Copy ExportSheet3.Cells(1, 1) '目的のデータの列をeSheet2の1列目へコピペ

MultipleData = False
With ExportSheet3
    LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    For r = 2 To LastRow
        If .Cells(r, 1) Like "*;*" Then
            MultipleData = True
            GoTo Multiple_Data
        End If
    Next r
    .Columns(1).RemoveDuplicates Columns:=1, Header:=xlYes
    .Columns(1).Sort key1:=.Cells(1, 1), Order1:=xlAscending, Header:=xlYes
    LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
End With

With ExportSheet2
    For r = 2 To LastRow
        ExportSheet1.Cells(1, r - 1) = ExportSheet3.Cells(r, 1)
        .UsedRange.AutoFilter Field:=SelectColumn, Criteria1:=ExportSheet3.Cells(r, 1)
        For Each TempCell In .Range(.Cells(2, 1), .Cells(.AutoFilter.Range.Rows.Count, 1)).SpecialCells(xlCellTypeVisible)
            TempLastRow = ExportSheet1.Cells(ExportSheet1.Rows.Count, r - 1).End(xlUp).Row
            ExportSheet1.Cells(TempLastRow + 1, r - 1) = TempCell.Value
        Next TempCell
    Next r
    .UsedRange.AutoFilter Field:=SelectColumn, Criteria1:="="
    If WorksheetFunction.Subtotal(3, .Columns(1)) > 1 Then
        For Each TempCell In .Range(.Cells(2, 1), .Cells(.AutoFilter.Range.Rows.Count, 1)).SpecialCells(xlCellTypeVisible)
            TempLastColumn = ExportSheet1.Cells(1, ExportSheet1.Columns.Count).End(xlToLeft).Column
            ExportSheet1.Cells(1, TempLastColumn + 1) = TempCell.Value
            ExportSheet1.Cells(2, TempLastColumn + 1) = TempCell.Value
        Next TempCell
    End If
End With
GoTo Complete

Multiple_Data:
With ExportSheet3
    .Cells(1, 2) = "Groups"
    LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    For r = 2 To LastRow
        If .Cells(r, 1) Like "*;*" Then
            SplitArr = Split(.Cells(r, 1), ";")
            For i = 0 To UBound(SplitArr)
                TempLastRow = .Cells(.Rows.Count, 2).End(xlUp).Row
                If SplitArr(i) <> "" Then .Cells(TempLastRow + 1, 2) = SplitArr(i)
            Next i
            Erase SplitArr
        Else
            TempLastRow = .Cells(.Rows.Count, 2).End(xlUp).Row
            If .Cells(r, 1) <> "" Then .Cells(TempLastRow + 1, 2) = .Cells(r, 1)
        End If
    Next r
    .UsedRange.RemoveDuplicates Columns:=2, Header:=xlYes
    LastRow = .Cells(.Rows.Count, 2).End(xlUp).Row
End With

With ExportSheet2
    For r = 2 To LastRow
        ExportSheet1.Cells(1, r - 1) = ExportSheet3.Cells(r, 2)
        .UsedRange.AutoFilter Field:=SelectColumn, Criteria1:=ExportSheet3.Cells(r, 2)
        If WorksheetFunction.Subtotal(3, .Columns(1)) > 1 Then
            For Each TempCell In .Range(.Cells(2, 1), .Cells(.AutoFilter.Range.Rows.Count, 1)).SpecialCells(xlCellTypeVisible)
                TempLastRow = ExportSheet1.Cells(ExportSheet1.Rows.Count, r - 1).End(xlUp).Row
                ExportSheet1.Cells(TempLastRow + 1, r - 1) = TempCell.Value
            Next TempCell
        End If
        
        .UsedRange.AutoFilter Field:=SelectColumn, Criteria1:=ExportSheet3.Cells(r, 2) & ";*"
        If WorksheetFunction.Subtotal(3, .Columns(1)) > 1 Then
            For Each TempCell In .Range(.Cells(2, 1), .Cells(.AutoFilter.Range.Rows.Count, 1)).SpecialCells(xlCellTypeVisible)
                TempLastRow = ExportSheet1.Cells(ExportSheet1.Rows.Count, r - 1).End(xlUp).Row
                ExportSheet1.Cells(TempLastRow + 1, r - 1) = TempCell.Value
                ExportSheet2.Cells(TempCell.Row, SelectColumn).Replace What:=ExportSheet3.Cells(r, 2), Replacement:=""
            Next TempCell
        End If
        
        .UsedRange.AutoFilter Field:=SelectColumn, Criteria1:="*;" & ExportSheet3.Cells(r, 2) & ";*"
        If WorksheetFunction.Subtotal(3, .Columns(1)) > 1 Then
            For Each TempCell In .Range(.Cells(2, 1), .Cells(.AutoFilter.Range.Rows.Count, 1)).SpecialCells(xlCellTypeVisible)
                TempLastRow = ExportSheet1.Cells(ExportSheet1.Rows.Count, r - 1).End(xlUp).Row
                ExportSheet1.Cells(TempLastRow + 1, r - 1) = TempCell.Value
                ExportSheet2.Cells(TempCell.Row, SelectColumn).Replace What:=ExportSheet3.Cells(r, 2), Replacement:=""
            Next TempCell
        End If
    
        .UsedRange.AutoFilter Field:=SelectColumn, Criteria1:="*;" & ExportSheet3.Cells(r, 2)
        If WorksheetFunction.Subtotal(3, .Columns(1)) > 1 Then
            For Each TempCell In .Range(.Cells(2, 1), .Cells(.AutoFilter.Range.Rows.Count, 1)).SpecialCells(xlCellTypeVisible)
                TempLastRow = ExportSheet1.Cells(ExportSheet1.Rows.Count, r - 1).End(xlUp).Row
                ExportSheet1.Cells(TempLastRow + 1, r - 1) = TempCell.Value
                ExportSheet2.Cells(TempCell.Row, SelectColumn).Replace What:=ExportSheet3.Cells(r, 2), Replacement:=""
            Next TempCell
        End If
    Next r
    
    .UsedRange.AutoFilter Field:=SelectColumn, Criteria1:="="
    If WorksheetFunction.Subtotal(3, .Columns(1)) > 1 Then
        For Each TempCell In .Range(.Cells(2, 1), .Cells(.AutoFilter.Range.Rows.Count, 1)).SpecialCells(xlCellTypeVisible)
            TempLastColumn = ExportSheet1.Cells(1, ExportSheet1.Columns.Count).End(xlToLeft).Column
            ExportSheet1.Cells(1, TempLastColumn + 1) = TempCell.Value
            ExportSheet1.Cells(2, TempLastColumn + 1) = TempCell.Value
        Next TempCell '次のセルへ
    End If '無所属の人がいるかのIf文の終了
End With

Complete:
ExportSheet1.UsedRange.EntireColumn.AutoFit
Application.DisplayAlerts = False
ExportSheet2.Delete
ExportSheet3.Delete
Application.DisplayAlerts = True

EndInTheMiddle:
ImportSheet.AutoFilterMode = False
If OriginalAutofilter = True Then ImportSheet.UsedRange.AutoFilter
 
DoNotStart:
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub



Sub Building_Adjacency_Matrix_from_Groups()

Application.Dialogs(xlDialogActivate).Show
Set ImportBook = ActiveWorkbook
If ImportBook.FullName = ThisWorkbook.FullName Then GoTo DoNotStart
Set ImportSheet = ActiveSheet
OriginalAutofilter = ImportSheet.AutoFilterMode
ImportSheet.AutoFilterMode = False

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Set ExportBook = Workbooks.Add
Set ExportSheet1 = ExportBook.Worksheets(1)
Set ExportSheet2 = Worksheets.Add(After:=ExportSheet1)
ExportSheet1.Cells(1, 1) = "ID"
 
ImportSheet.UsedRange.Copy ExportSheet2.Cells(1, 1)
With ExportSheet2
    Arr() = .UsedRange.Value
    .UsedRange.NumberFormatLocal = "0_ "
    .UsedRange.Value = Arr()
    LastColumn = .Cells(1, .Columns.Count).End(xlToLeft).Column
    For c = 1 To LastColumn
        LastRow = .Cells(.Rows.Count, c).End(xlUp).Row
        For r = 2 To LastRow
            TempLastRow = ExportSheet1.Cells(ExportSheet1.Rows.Count, 1).End(xlUp).Row
            .Cells(r, c).Copy ExportSheet1.Cells(TempLastRow + 1, 1)
        Next r
    Next c
End With

With ExportSheet1
    .Columns(1).RemoveDuplicates Columns:=1, Header:=xlYes
    .Columns(1).Sort key1:=.Cells(1, 1), Order1:=xlAscending, Header:=xlYes
    LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    For c = 2 To LastRow
        .Cells(c, 1).Copy .Cells(1, c)
    Next c
End With

With ExportSheet2
    For c = 1 To LastColumn
        LastRow = .Cells(.Rows.Count, c).End(xlUp).Row
        If LastRow < 3 Then GoTo NotGroup
        Erase RowArr()
        ReDim RowArr(0)
        For r = 2 To LastRow
            TempRow = WorksheetFunction.Match(.Cells(r, c), ExportSheet1.Columns(1), 0)
            RowArr(UBound(RowArr)) = TempRow
            ReDim Preserve RowArr(UBound(RowArr) + 1)
        Next r
        ReDim Preserve RowArr(UBound(RowArr) - 1)
        For i = 0 To UBound(RowArr)
            For i2 = 0 To UBound(RowArr)
                If i <> i2 Then
                    ExportSheet1.Cells(RowArr(i), RowArr(i2)) = 1
                End If
            Next i2
        Next i
NotGroup:
    Next c
End With

ExportSheet1.UsedRange.SpecialCells(xlCellTypeBlanks) = 0
ExportSheet1.UsedRange.EntireColumn.AutoFit
Application.DisplayAlerts = False
ExportSheet2.Delete
Application.DisplayAlerts = True
ImportSheet.AutoFilterMode = False
If OriginalAutofilter = True Then ImportSheet.UsedRange.AutoFilter
 
DoNotStart:
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
