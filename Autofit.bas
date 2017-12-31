Attribute VB_Name = "Autofit"
Option Explicit

Public Sub btnMergedAreaRowAutofit_onAction(control As IRibbonControl)
    Call MergedAreaRowAutofit
End Sub

Sub MergedAreaRowAutofit()
    Dim rng As Range ' Range to operate on
    
    Dim curRow As Range ' current row in rng being operated on
    Dim curRowIdx As Long ' current row index in rng
    Dim curCell As Range ' current cell (column) in curRow being operated on
    Dim curColIdx As Long ' current column index in curRow being worked on
    
    Dim mergedArea As Range ' Merged area to find the height for
    Dim mAreaCol As Range ' Column in the merged area being worked on
    Dim spareCell As Range ' A spare cell to work out the height for mergedArea
    
    Dim mergedAreaWidth As Double 'merge width
    Dim RH As Double 'row height
    Dim MaxRH As Double ' Max autofit height for current row being worked on
    Dim doAutofit As Boolean ' Do an autofit on row if MaxRH is not obtained
    
    Const SpareCol  As Long = 16384 ' Spare column used to work out row heights
    
    If Not TypeOf Application.Selection Is Range Then
        MsgBox "No Range Selected"
        Exit Sub
    End If
    Set rng = Application.Selection
    Set rng = Application.Intersect(rng, rng.Worksheet.UsedRange)
    
    Application.ScreenUpdating = False

    ' loop through each row in the working range
    For curRowIdx = 1 To rng.Rows.Count
        Set curRow = rng.Rows(curRowIdx)
        Application.StatusBar = "Autofitting row " & curRowIdx & " of " & rng.Rows.Count
         'if the row is not hidden
        If Not curRow.EntireRow.Hidden Then
            'if the cells have data
            If Application.WorksheetFunction.CountA(curRow) Then
                
                MaxRH = 0
                doAutofit = False
                ' Loop through each column in working row
                For curColIdx = rng.Columns.Count To 1 Step -1
                    Set curCell = rng.Cells(curRowIdx, curColIdx)
                    ' if the current cell contains text
                    ' Note: Should this be Application.WorksheetFunction.Type = 2 ?
                    '       It seems like it would acheive the same thing.
                    If VarType(curCell.Value) = vbString Then
                        ' if current cell is merged
                        If curCell.MergeCells Then
                            ' Get the merged area
                            Set mergedArea = curCell.MergeArea
    
                            If mergedArea.WrapText Then
                                'get the total width
                                mergedAreaWidth = 0
                                For Each mAreaCol In mergedArea.Columns
                                    mergedAreaWidth = mergedAreaWidth + mAreaCol.ColumnWidth + 0.66
                                Next mAreaCol
                                
                                'use the spare column,
                                'put the value,
                                'make autofit,
                                'get the row height
                                Set spareCell = mergedArea.EntireRow.Cells(1, SpareCol)
                                spareCell.Value = mergedArea.Value
                                spareCell.ColumnWidth = mergedAreaWidth
                                spareCell.WrapText = True
                                spareCell.EntireRow.Autofit
                                RH = spareCell.RowHeight
                                If RH > MaxRH Then MaxRH = RH
                                spareCell.Value = vbNullString
                                spareCell.WrapText = False
                                spareCell.ColumnWidth = 8.43
                            End If
                            
                        ElseIf curCell.WrapText Then
                            doAutofit = True
                        End If
                    End If
                Next curColIdx
                
                If MaxRH > 0 Then
                    mergedArea.RowHeight = MaxRH
                ElseIf doAutofit Then
                    curRow.Autofit
                End If
            
            End If
        End If
    Next curRowIdx
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
End Sub

