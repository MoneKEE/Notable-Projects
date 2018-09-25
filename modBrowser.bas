Attribute VB_Name = "modBrowser"
'*******************************************
'Revision History
'Date       Developer       Description
'12-Oct-16  M Gore          Initial Version
'*******************************************

Option Compare Text

Private Const mModuleName As String = "modBrowser"

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public Function GetIE() As SHDocVw.InternetExplorer
' Returns an internet explorer object that is bound to
' either a new instance or an active instance of Internet Explorer.

    Dim IE          As SHDocVw.InternetExplorer
    Dim shellWins   As New SHDocVw.ShellWindows
    Dim ThisBrwsr   As SHDocVw.InternetExplorer
    
    On Error Resume Next
    
    ' Loop through the collection of active shell
    ' windows and set the first window with a
    ' Location URL that begins with http.
    For Each IE In shellWins
        If InStr(IE.LocationURL, "http") Then
            Set ThisBrwsr = IE
        End If
    Next
    
    Debug.Print Err.Description
    
    On Error GoTo 0
    
    ' If no active IE objects were found create
    ' a new one.
    If ThisBrwsr Is Nothing Then
        Set ThisBrwsr = Interaction.CreateObject("InternetExplorer.Application")
        ThisBrwsr.Visible = True
    End If
    
    Set GetIE = ThisBrwsr
    
    Set IE = Nothing
    Set ThisBrwsr = Nothing
End Function

Public Function FocusWindow()
' Puts the IE browser window in focus
    Dim iret    As Long
    Dim IE      As SHDocVw.InternetExplorer
    
    Set IE = GetIE
    
    iret = BringWindowToTop(IE.hwnd)
    
    Set IE = Nothing
End Function

