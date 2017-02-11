Attribute VB_Name = "ThisWorkbook"
' AutoBackup VBA code everytime workbook is saved
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    VBABackup()
End Sub ' Workbook_BeforeSave
