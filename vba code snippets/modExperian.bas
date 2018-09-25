Attribute VB_Name = "modExperian"
'*******************************************
'Revision History
'Date       Developer       Description
'29-Sep-16  M Gore          Initial Version
'*******************************************

Option Compare Text

Private Const mModuleName As String = "modExperian"

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub GetExtract(Optional DtStart As String, Optional DtEnd As String)
' Drive the
    Dim IE As SHDocVw.InternetExplorer
    Dim TargetURL As String
    Dim sDate As String
    Dim eDate As String
    
    Application.DisplayAlerts = False
    
    Set IE = modBrowser.GetIE
    
    TargetURL = "https://onesource.passporthealth.com/"
    sDate = Str$(DateAdd("d", -2, Date))
    eDate = Str$(DateAdd("d", 0, Date))
    
    ' Show or hide the browser window?
    IE.Visible = True
    
    Sleep 2000
    
    IE.Navigate TargetURL
    
    Do: DoEvents: Sleep 100: Loop Until IE.Busy = False
    Do: DoEvents: Sleep 100: Loop Until IE.ReadyState = READYSTATE_COMPLETE
    
    Login
    
    LaunchECareNext
    LaunchWorkCenterStatusReport
    GenerateStatusReport sDate, eDate
    ExportStatusReport
    SaveFileAs

    IE.Quit
    
    Set IE = Nothing
End Sub

Private Function LaunchSelfServicePortal(ByVal target As String)
    Dim IE As SHDocVw.InternetExplorer
    
    Set IE = modBrowser.GetIE
    IE.Navigate target + "/_members/CustSup/SelfService/NPI_Maintenance.aspx?Reset=1"
    
    Do: DoEvents: Sleep 100: Loop Until IE.Busy = False
    Do: DoEvents: Sleep 100: Loop Until IE.ReadyState = READYSTATE_COMPLETE
    
    Set IE = Nothing
End Function

Private Function LaunchECareNext()
    Dim IE As SHDocVw.InternetExplorer
    Dim link As MSHTML.IHTMLAnchorElement
    Dim path1 As String, path2 As String, path3 As String
    
    Set IE = modBrowser.GetIE
    Sleep 5000
    Set link = IE.Document.getElementById("pnlMenuLinks").all(8)
    path1 = Split(link.pathname, ",")(0)
    path2 = Split(path1, "'")(1)
    path3 = Split(path2, "'")(0)
    IE.Navigate path3
    
    Do: DoEvents: Sleep 100: Loop Until IE.Busy = False
    Do: DoEvents: Sleep 100: Loop Until IE.ReadyState = READYSTATE_COMPLETE
    
    Set IE = Nothing
End Function

Private Function LaunchWorkCenterStatusReport()
    Dim IE As SHDocVw.InternetExplorer
    Dim Lnk As MSHTML.HTMLAnchorElement

    
    Set IE = modBrowser.GetIE
    Set Lnk = IE.Document.getElementById("WorkCenterStatusReportToolLink")
    Lnk.Click
    
    Do: DoEvents: Sleep 100: Loop Until IE.Busy = False
    Do: DoEvents: Sleep 100: Loop Until IE.ReadyState = READYSTATE_COMPLETE
    
    Set IE = Nothing
    Set Lnk = Nothing
End Function

Private Function GenerateStatusReport(ByVal sDate As String, ByVal fDate As String)
    Dim IE As SHDocVw.InternetExplorer
    Dim Frm As Object
    Dim btn As MSHTML.HTMLInputElement
    Dim btnBack As MSHTML.HTMLInputElement
    Dim Selct As MSHTML.HTMLSelectElement
    Dim startdate As MSHTML.HTMLInputElement
    Dim enddate As MSHTML.HTMLInputElement

    Set IE = modBrowser.GetIE
    Set startdate = IE.Document.getElementsByClassName("datepicker start-date hasDatepicker")(0)
    Set enddate = IE.Document.getElementsByClassName("datepicker end-date hasDatepicker")(0)
    Set btn = IE.Document.getElementsByClassName("btnProcess btnModal newbtn")(0)
    Set btnBack = IE.Document.getElementsByClassName("btnBack btnModal newbtn")(0)
    Set Selct = IE.Document.getElementsByName("D2")(0)
    
    Do
        startdate.Value = sDate
        enddate.Value = fDate
        Selct.selectedIndex = 3
    Loop Until startdate.Value = sDate
    
    Sleep 2000
    btn.Click
    
    Set Frm = IE.Document.getElementsByClassName("processing")

    Do: DoEvents: Sleep 100: Loop Until InStr(btnBack.outerHTML, "block")
    
    Set IE = Nothing
    Set startdate = Nothing
    Set enddate = Nothing
    Set btn = Nothing
    Set btnBack = Nothing
    Set Frm = Nothing
    Set Selct = Nothing
End Function

Private Function ExportStatusReport()
    Dim IE As SHDocVw.InternetExplorer
    Dim btnExport As MSHTML.HTMLInputElement
    Set IE = modBrowser.GetIE
    
    Set btnExport = IE.Document.getElementsByClassName("btnExport btnModal newbtn")(0)
    
    Do: DoEvents: Sleep 100: Loop Until InStr(btnExport.outerHTML, "block")
    btnExport.Click
    
    Set Frm = IE.Document.getElementsByClassName("processing")
    
    Do: DoEvents: Sleep 100: Loop Until Frm.Length = 1
    Do: DoEvents: Sleep 100: Loop Until IE.ReadyState = READYSTATE_COMPLETE
    Do: DoEvents: Sleep 100: Loop Until Frm.Length = 0
    
    Set btnExport = Nothing
    Set IE = Nothing
    Set Frm = Nothing
End Function

Private Function SaveFileAs()
' Would love to find a way to initiate the file download without using
' SendKeys.  Perhaps google chrome?
    Dim IE As SHDocVw.InternetExplorer
    Dim SA As Object
    
    Set IE = modBrowser.GetIE
    
    Sleep 10000
    modBrowser.FocusWindow
    Sleep 5000
    Interaction.SendKeys "%N"
    Sleep 5000
    Interaction.SendKeys "{TAB}{ENTER}"
    Sleep 5000
    
    Set IE = Nothing
End Function

Private Function Login()
    Dim IE As SHDocVw.InternetExplorer
    Dim User As MSHTML.IHTMLElementCollection
    Dim Pwd As MSHTML.IHTMLElementCollection
    Dim Frm As HTMLFormElement
    
    Set IE = modBrowser.GetIE
    Set User = IE.Document.getElementsByName("OSAuthLoginUN")
    Set Pwd = IE.Document.getElementsByName("OSAuthLoginPWD")
    Set Frm = IE.Document.forms(0)
    
    For Each rw In User
        rw.Value = "mig9119"
    Next
    
    For Each rw In Pwd
        rw.Value = "rooD3nee"
    Next
    
    Frm.submit
    
    Set User = Nothing
    Set Pwd = Nothing
    Set Frm = Nothing
End Function

