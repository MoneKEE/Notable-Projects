Attribute VB_Name = "modFileSystem"
'*******************************************
'Revision History
'Date       Developer       Description
'12-Oct-16  M Gore          Initial Version
'*******************************************

Option Compare Text

Private Const mModuleName As String = "modFileSystem"

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public Function GetFileName() As String
    Path = "C:\users\mig9119\Downloads\"
    file = Dir(Path)
    
    Do Until InStr(file, "WorkCenter")
        file = Dir()
    Loop
    
    GetFileName = Path + file
End Function
