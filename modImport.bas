Attribute VB_Name = "modImport"
Option Compare Database

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub main()
    Dim filename As String

    filename = modFileSystem.GetFileName
    
    TouchWB (filename)
    ImportSheet (filename)
    modFileSystem.DeleteFile (filename)
End Sub

Sub ImportSheet(ByVal filename As String)
    Dim LastRow As Long
    
    DoCmd.TransferSpreadsheet acImport, , "Staging", filename, True
End Sub

Sub GrabReportSeries()
    Dim span As Date
    Dim sdDate As Date
    Dim i As Integer
    
    i = 0
    sdDate = #10/6/2016#
    
    Do While span < Date
        span = DateAdd("d", 4 * i, sdDate)
        
        Debug.Print "Beginning Extract for: " & DateAdd("d", -3, span) & " - " & span
        modExperian.GetExtract (span)
        i = i + 1
        Sleep 3000
    Loop
    
End Sub

Private Function GetLastRow(ByVal filename As String) As Long
    Dim wkb As Excel.Application
    
    Set wkb = CreateObject("Excel.Application")
    wkb.Application.Visible = False
    wkb.Workbooks.Open filename
    
    GetLastRow = wkb.Workbooks(1).ActiveSheet.Range("A655565").End(xlUp).Row
End Function

Function TouchWB(ByVal filename As String)
    Dim wkb As Excel.Application
    
    Set wkb = CreateObject("Excel.Application")
    wkb.Application.Visible = False
    wkb.Workbooks.Open filename
    wkb.Workbooks(1).Save
    wkb.Workbooks.Close
    wkb.Quit
    Set wkb = Nothing
End Function
