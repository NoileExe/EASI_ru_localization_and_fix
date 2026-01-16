' Главный файл патча
Sub ApplyAllChanges()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    ApplyPart1 wb
    ApplyPart2 wb
    MsgBox "Готово!", vbInformation
End Sub
