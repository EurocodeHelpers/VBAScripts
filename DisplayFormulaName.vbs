Function display_formula(rngIn As Range) As String
    If rngIn.Value <> "" Then
        Dim i As Integer
        Dim j As Integer
        Dim l As Integer
        Dim s As String
        Dim t As String
        s = rngIn.Formula
        l = Len(s)
        t = ""
        i = 1
        Do
            t = t & Mid(s, i, 1)
            If InStr("=+-*/^()", Mid(s, i, 1)) Then
                j = i + 1
                Do
                    i = i + 1
                Loop Until InStr("=+-*/^()", Mid(s, i, 1))
                If Mid(s, i, 1) = "(" Or InStr("0123456789", Mid(s, j, 1)) Then
                    t = t & Mid(s, j, i - j)
                ElseIf i > j Then
                    t = t & Range(Mid(s, j, i - j)).Text
                End If
            Else
                i = i + 1
            End If
        Loop Until i > l
        display_formula = Right(t, Len(t) - 1)
    Else
        display_formula = ""
    End If
End Function

Function display_formula_name(cell)
' dispalys cell formula in another cell
    display_formula_name = Right(cell.Formula, Len(cell.Formula) - 1)
End Function


Sub fix_display_formula()
no_sheets = ActiveWorkbook.Sheets.Count
For Sheet = 1 To no_sheets
    Sheets(Sheet).Activate
    For rowa = 1 To 65
        For Column = 1 To 35
        cell_contents = Cells(rowa, Column).Formula
        If InStr(cell_contents, "formula") Then
            Cells(rowa, Column) = cell_contents
        End If
        Next Column
    Next rowa
Next Sheet
Sheets(1).Select
End Sub
