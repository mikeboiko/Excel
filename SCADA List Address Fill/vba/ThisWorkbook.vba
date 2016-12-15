' ===
' Install
' ===

' - Copy this code into "ThisWorkbook" on any projects you wish to enable the plain text backup
' - Enable "Microsoft Visual Basic for Applications Extensibility 5.x"
'   -> Tools>References
'   -> Find "... 5.x" and check to enable
'   -> "OK"
'
' In order for auto-save Macro to work, Application.EnableEvents needs to be True

Private Const CODE_START_LINE_DEFAULT = 1

Option Explicit

' Triggered automatically on user save event and used to create a backup of the current workbook's vba as plain-text
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' ===
    ' Debug
    ' ===

    ' Define types
    Dim DebugEnabled As Boolean

    ' Enable/disable debugging here
    DebugEnabled = False

    ' ===
    ' Main
    ' ===

    ' Define types
    Dim Code As CodeModule
    Dim CodeLine As Long
    Dim CodeLineCount As Long
    Dim FilePath As String
    Dim FolderPath As String
    ' ---
    ' Read
    ' ---

    Set Code = Application.VBE.ActiveCodePane.CodeModule
    ' Number of lines in the code
    CodeLineCount = Code.CountOfLines()
    ' Get current module code
    ' Determine file path for the generated vba file
    
    'If vba subfolder doesn't exist, create it
    FolderPath = Application.ActiveWorkbook.Path & "/vba/"
    If Dir(FolderPath, vbDirectory) = "" Then MkDir FolderPath
    
    FilePath = Application.ActiveWorkbook.Path & "/vba/" & Code.Name & ".vba"

    ' ---
    ' Write
    ' ---
    
    ' Open file by file path
    ' TODO [161205] - Create folders/files when they do not exist
    Open FilePath For Output As #1
        ' Print current module code to the open vba file
        Print #1, Code.Lines(CODE_START_LINE_DEFAULT, CodeLineCount)
    Close #1 ' Close file

    ' * Debug Output
    If DebugEnabled Then
        ' Display the output of the current module
        Debug.Print Code.Lines(CODE_START_LINE_DEFAULT, CodeLineCount)
    End If


End Sub




