Private Sub Worksheet_Change(ByVal Target As Range)
    Dim TotalRows As Long
    Dim FirstRow As Long
    Dim OtherRow As Long
    Dim CurrentColumn As Long
    Dim FirstRowOccurrences As Long
    Dim OtherRowOccurrences As Long
    Dim StartColumn As Long
    Dim EndColumn As Long
    Dim DuplicateRow As Boolean
    Dim MainRange As Range
    Dim OtherRange As Range
    Dim ColorIndex As Long
    Dim BlankColor As Long
    
    BlankColor = 0
    ColorIndex = 0
    StartColumn = 65
    EndColumn = 65 + 15
    TotalRows = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    
    For FirstRow = 1 To TotalRows
        'init color highlighting
        Range(Chr(StartColumn) & FirstRow & ":" & Chr(EndColumn) & FirstRow).Interior.ColorIndex = 0
        Next FirstRow
    
    For FirstRow = 1 To TotalRows - 1
        Set MainRange = Range(Chr(StartColumn) & FirstRow & ":" & Chr(EndColumn) & FirstRow)
        For OtherRow = FirstRow + 1 To TotalRows
            DuplicateRow = True
            Set OtherRange = Range(Chr(StartColumn) & OtherRow & ":" & Chr(EndColumn) & OtherRow)
            'loop on cols as ascii value
            For CurrentColumn = StartColumn To EndColumn
                'comparing occurrences of every index on two lines
                If Range(Chr(CurrentColumn) & FirstRow).Value <> "" Then
                    'by the first row
                    FirstRowOccurrences = Application.WorksheetFunction.CountIf(MainRange, Range(Chr(CurrentColumn) & FirstRow).Value)
                    OtherRowOccurrences = Application.WorksheetFunction.CountIf(OtherRange, Range(Chr(CurrentColumn) & FirstRow).Value)
                    If FirstRowOccurrences <> OtherRowOccurrences Then
                        DuplicateRow = False
                    End If
                End If
                If Range(Chr(CurrentColumn) & OtherRow).Value <> "" Then
                    'by the other row - the second line could be longer for example
                    FirstRowOccurrences = Application.WorksheetFunction.CountIf(MainRange, Range(Chr(CurrentColumn) & OtherRow).Value)
                    OtherRowOccurrences = Application.WorksheetFunction.CountIf(OtherRange, Range(Chr(CurrentColumn) & OtherRow).Value)
                    If FirstRowOccurrences <> OtherRowOccurrences Then
                        DuplicateRow = False
                    End If
                End If
                Next CurrentColumn
            If DuplicateRow = True And Application.WorksheetFunction.CountIf(MainRange, "") <> 16 Then
                If MainRange.Interior.ColorIndex > BlankColor Then
                    'check if it a new row join to existing couple
                    OtherRange.Interior.ColorIndex = MainRange.Interior.ColorIndex
                Else
                    MainRange.Interior.ColorIndex = (ColorIndex Mod 53) + 3
                    OtherRange.Interior.ColorIndex = (ColorIndex Mod 53) + 3
                    ColorIndex = ColorIndex + 1
                End If
            End If
            Next OtherRow
        Next FirstRow
        
                
End Sub

