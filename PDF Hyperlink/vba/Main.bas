Attribute VB_Name = "Main"
' Open PDF to specific page
Sub OpenPDF()

    App_Path = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    File_Path = "C:\Users\mike.boiko\Documents\Code\Hyperlink DB\PPL006-Electrical_Stick_File-160714.pdf"
    'Page_Num = "A16.02-1-0105B_R3"
    Page_Num = "5"
    'Shell Chr(34) & App_Path & Chr(34) & " /A Nameddest=" & Page_Num & " " & Chr(34) & File_Path & Chr(34), vbMaximizedFocus
    Shell Chr(34) & App_Path & Chr(34) & " /A Page=" & Page_Num & " " & Chr(34) & File_Path & Chr(34), vbMaximizedFocus
    'Shell App_Path & " /A Page=" & Page_Num & " " & File_Path, vbMaximizedFocus

End Sub
