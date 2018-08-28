Attribute VB_Name = "modToolKit"
'*******************************************
'Revision History
'Date       Developer       Description
'14-Nov-16  M Gore          Initial Version
'21-Dec-16  M Gore
'*******************************************
' Tool Kit

' Source and Destination sheets
Public Const gACCOUNTSSHEET As String = "Accounts"
Public Const gRAWDATASHEET As String = "raw data alt"
Public Const gGENERALSHEET As String = "General"
Public Const gAGENTSSHEET As String = "Agents"
Public Const gSCRAPSHEET As String = "scrap"
Public Const gSTATEPATHSHEET As String = "State Path Distribution"
Public Const gSTATUSFREQSHEET As String = "Status Frequency"
Public Const gREFERENCESHEET As String = "reference"
Public Const gTIMESERIESSHEET As String = "Time Series"

' Account Features List
Public Const gPatTypeColumn As Integer = 2
Public Const gPriInsColumn As Integer = gPatTypeColumn + 1
Public Const gStatusColumn As Integer = gPriInsColumn + 1
Public Const gDtCreatedColumn As Integer = gStatusColumn + 1
Public Const gDoneStatusColumn As Integer = gDtCreatedColumn + 1
Public Const gLeadtimeColumn As Integer = gDoneStatusColumn + 1
Public Const gHopsColumn As Integer = gLeadtimeColumn + 1
Public Const gFirstTouchColumn As Integer = gHopsColumn + 1
Public Const gCompleteColumn As Integer = gFirstTouchColumn + 1
Public Const gTTCColumn As Integer = gCompleteColumn + 1
Public Const gAgeColumn As Integer = gTTCColumn + 1
Public Const gLullTimeColumn As Integer = gAgeColumn + 1
Public Const gTransDelayColumn As Integer = gLullTimeColumn + 1
Public Const gMaxDelayColumn As Integer = gTransDelayColumn + 1
Public Const gTotDelayColumn As Integer = gMaxDelayColumn + 1
Public Const gFTouchByColumn As Integer = gTotDelayColumn + 1
Public Const gLTouchByColumn As Integer = gFTouchByColumn + 1
Public Const gCWXColumn As Integer = gLTouchByColumn + 1
Public Const gCBDOSColumn As Integer = gCWXColumn + 1
Public Const gState1Column As Integer = gCBDOSColumn + 1
Public Const gState2Column As Integer = gState1Column + 1
Public Const gState3Column As Integer = gState2Column + 1
Public Const gState4Column As Integer = gState3Column + 1
Public Const gState5Column As Integer = gState4Column + 1
Public Const gState6Column As Integer = gState5Column + 1
Public Const gState7Column As Integer = gState6Column + 1
Public Const gState8Column As Integer = gState7Column + 1
Public Const gState9Column As Integer = gState8Column + 1
Public Const gState10Column As Integer = gState9Column + 1
Public Const gState11Column As Integer = gState10Column + 1
Public Const gState12Column As Integer = gState11Column + 1
Public Const gState13Column As Integer = gState12Column + 1
Public Const gState14Column As Integer = gState13Column + 1
Public Const gState15Column As Integer = gState14Column + 1
Public Const gState16Column As Integer = gState15Column + 1
Public Const gBindColumn As Integer = gState16Column + 1
Public Const gBindTColumn   As Integer = gBindColumn + 1

' Agent Features List

' State Path Features List
Public Const gTraverseColumn = 2
Public Const gTraverseCntColumn = gTraverseColumn + 1
Public Const gAvgAgeColumn = gTraverseCntColumn + 1
Public Const gMedAgeColumn = gAvgAgeColumn + 1
Public Const gWorstAgeColumn = gMedAgeColumn + 1
Public Const gAvgActvAgeColumn = gWorstAgeColumn + 1
Public Const gMedActvAgeColumn = gAvgActvAgeColumn + 1
Public Const gWorstActvAgeColumn = gMedActvAgeColumn + 1
Public Const gAvgFTouchColumn = gWorstActvAgeColumn + 1
Public Const gMedFTouchColumn = gAvgFTouchColumn + 1
Public Const gWorstFTouchColumn = gMedFTouchColumn + 1
Public Const gTransitionsColumn = gWorstFTouchColumn + 1

' Status Frequency Features List


' Time Series Features List
Public Const gCaseCountColumn As Integer = 1
Public Const gActiveCasesColumn As Integer = gCaseCountColumn + 1
Public Const gTransCountColumn As Integer = gActiveCasesColumn + 1
Public Const gTransCompleteColumn As Integer = gTransCountColumn + 1
Public Const gTransOpenColumn As Integer = gTransCompleteColumn + 1
Public Const gStatCompleteColumn As Integer = gTransOpenColumn + 1
Public Const gStatOpenColumn As Integer = gStatCompleteColumn + 1
Public Const gCreatedColumn As Integer = gStatOpenColumn + 1
Public Const gOutToInColumn As Integer = gCreatedColumn + 1
Public Const gOutToTotalColumn As Integer = gOutToInColumn + 1
Public Const gStatDistrColumn As Integer = gOutToTotalColumn + 1

Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

' Function Declarations
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function CleanRawData()
    Dim lLastRow As Long
    Dim lRefRow As Long
    Dim lTrueRow As Long
    Dim rCatch As Range
    Dim iStatRow As Integer
    
    modToolKit.OptimizeCodeBegin (gRAWDATASHEET)
    
    lRefRow = Sheets(gRAWDATASHEET).Range("A1048576").End(xlUp).Row
    lLastRow = Sheets(gRAWDATASHEET).Range("R1048576").End(xlUp).Row
    
    If lRefRow = lLastRow Then
        Exit Function
    Else
        lTrueRow = lLastRow + 1
        iStatRow = Sheets(gREFERENCESHEET).Range("A1048576").End(xlUp).Row
        Set rCatch = Sheets(gRAWDATASHEET).Range("J" & lTrueRow & ":R" & lRefRow)
        
    ' Correct Timestamps
        Sheets(gRAWDATASHEET).Range("S" & lTrueRow & ":S" & lRefRow).Formula = "=VALUE(J" & lTrueRow & ")"
        Sheets(gRAWDATASHEET).Range("T" & lTrueRow & ":T" & lRefRow).Formula = "=VALUE(K" & lTrueRow & ")"
        Sheets(gRAWDATASHEET).Range("U" & lTrueRow & ":U" & lRefRow).Formula = "=VALUE(L" & lTrueRow & ")"
        
        Sheets(gRAWDATASHEET).Range("S" & lTrueRow & ":U" & lRefRow).Copy
        Sheets(gRAWDATASHEET).Range("S" & lTrueRow & ":U" & lRefRow).PasteSpecial xlPasteValues
        Sheets(gRAWDATASHEET).Range("S" & lTrueRow & ":U" & lRefRow).Copy
        Sheets(gRAWDATASHEET).Range("J" & lTrueRow & ":L" & lRefRow).PasteSpecial xlPasteValues
        
    ' Add TS Date Part
        rCatch.Columns(8).Formula = "=INT(J" & lTrueRow & ")"
        
    ' Add DC Date Part
        rCatch.Columns(9).Formula = "=INT(L" & lTrueRow & ")"
        
    ' Correct DoneStatus Values
        rCatch.Columns(4).Formula = "=VLOOKUP(I" & lTrueRow & ",reference!$A$5:$B$" & iStatRow & ",2,FALSE)"
        
    ' Add Leadtime
        rCatch.Columns(5).Formula = "=IF(K" & lTrueRow & "- L" & lTrueRow & "< 0,0, K" & lTrueRow & "- L" & lTrueRow & ")"
        
    ' Add DOS - STATUS
        rCatch.Columns(6).Formula = "=K" & lTrueRow & "- J" & lTrueRow
        
    ' Add STATUS - CREATED
        rCatch.Columns(7).Formula = "=J" & lTrueRow & "- L" & lTrueRow
 
    ' Format Time Stamps
        rCatch.Columns(1).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
        rCatch.Columns(2).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
        rCatch.Columns(3).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
        
        rCatch.Columns(8).NumberFormat = "mm/dd/yyyy"
        rCatch.Columns(9).NumberFormat = "mm/dd/yyyy"
        
        Sheets(gRAWDATASHEET).Range("S:U").Delete
    End If

    modToolKit.OptimizeCodeEnd (gRAWDATASHEET)
End Function

Public Function ClearRange(ByVal strRngName As String)
    Dim rngData As Range
    Dim lLastRow As Long
    Dim lFirstRow As Long
    Dim iLastCol As Integer
    Dim strLastCol As String
    
    lLastRow = Sheets(strRngName).Range("A1048576").End(xlUp).Row
    
    Select Case strRngName
        Case gRAWDATASHEET
            If blnHeader = True Then
                lFirstRow = 1
            Else
                lFirstRow = 2
            End If
        Case gACCOUNTSSHEET, gTIMESERIESSHEET, gAGENTSSHEET
            If blnHeader = True Then
                lFirstRow = 3
            Else
                lFirstRow = 4
            End If
    End Select
    
    If lLastRow <> 3 Then
        iLastCol = Sheets(strRngName).Range("AZ3").End(xlToLeft).Column
        strLastCol = GetColumnLetter(iLastCol)
        
        Sheets(strRngName).Range("A4" & ":" & strLastCol & lLastRow).ClearContents
    End If
    
End Function
Public Function GetRange(ByVal strRngName As String, ByVal blnHeader As Boolean) As Range
    Dim rngData As Range
    Dim lLastRow As Long
    Dim lFirstRow As Long
    Dim iLastCol As Integer
    Dim strLastCol As String
    
    Sheets(strRngName).Activate
    
    If Sheets(strRngName).FilterMode = True Then
        Sheets(strRngName).ShowAllData
        Sleep 2000
    End If
    
    iLastCol = Sheets(strRngName).Range("AZ3").End(xlToLeft).Column
    strLastCol = GetColumnLetter(iLastCol)
    
    lLastRow = Sheets(strRngName).Range("A1048576").End(xlUp).Row
    lFirstRow = Sheets(strRngName).Range("B1").End(xlDown).Row

    Select Case strRngName
        Case gRAWDATASHEET
            If blnHeader = True Then
                lFirstRow = 1
            Else
                lFirstRow = 2
            End If
        Case gACCOUNTSSHEET, gTIMESERIESSHEET, gAGENTSSHEET, gSTATEPATHSHEET, gSTATUSFREQSHEET
            If blnHeader = True Then
                lFirstRow = 3
            ElseIf lFirstRow = lLastRow Then
                lFirstRow = 4
                lLastRow = 4
            Else
                lFirstRow = 4
            End If
    End Select
    
    Set rngData = Sheets(strRngName).Range("A" & lFirstRow & ":" & strLastCol & lLastRow)
    
    Set GetRange = rngData
    
    Set rngData = Nothing
End Function

Public Function GetColumnLetter(ColumnNumber As Integer) As String
    ' Convert the column number to the column letter
    If ColumnNumber > 26 Then
        '1st character: Subtract 1 to map the characters to 0-25,
        '               but you don't have to remap back to 1-26
        '               after the 'Int' operation since columns
        '               1-26 have no prefix letter
        
        '2nd character: Subtract 1 to map the characters to 0-25,
        '               but then must remap back to 1-26 after the
        '               'Mod' operation by adding 1 back in
        '               (included in the '65')
        
        GetColumnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & _
                       Chr(((ColumnNumber - 1) Mod 26) + 65)
    Else
        ' Columns A-Z
        GetColumnLetter = Chr(ColumnNumber + 64)
    End If
End Function

Public Function GetConcatString(ByVal StartCol As Integer) As String
    Dim ColLetter As String
    Dim ColLetter2 As String
    Dim ConcatString As String
    Dim i As Integer
    
    
    For i = StartCol To StartCol + 15
        ColLetter = GetColumnLetter(i)
        ColLetter2 = GetColumnLetter(i + 1)
        If i = StartCol + 15 Then
            ConcatString = ConcatString & "$" & ColLetter & "4"
        Else
            ConcatString = ConcatString & "$" & ColLetter & "4,IF(Len(" & ColLetter2 & 4 & ")=0,"""","" -> ""),"
        End If
    Next i
    
    GetConcatString = ConcatString
End Function

Public Function FormatTimeValues()
    Dim DateRange As Range
    Dim lLastRow As Long
    Dim rSort As Range
    
    Set rSort = GetSortRange
    
    lLastRow = rSort.Rows.Count
    
    Set DateRange = rSort.Range("J:L")

    DateRange.Insert
    
    With rSort.Range("J2:L" & lLastRow)
        .Formula = "=VALUE(M2)"
        .Copy
        .PasteSpecial (xlPasteValues)
        .NumberFormat = "mm/dd/yyyy hh:mm am/pm"
    End With
    
    rSort.Range("J1:L1").Delete Shift:=xlShiftToLeft
    DateRange.Range("A2:C" & lLastRow).Delete Shift:=xlShiftToLeft
    
    Set DateRange = Nothing
    Set rSort = Nothing
End Function

Public Function AddMetricsColumns()
    Dim MetricsRange As Range
    Dim lLastRow As Long
    Dim rSort As Range
    
    Set rSort = GetSortRange
    
    lLastRow = rSort.Rows.Count
    
    Set MetricsRange = rSort.Range("N1:P" & lLastRow)
    
    With MetricsRange.Columns(1)
        .Range("A1").Value = "Leadtime"
        .Range("A1").Font.Bold = True
        .Offset(1, 0).Formula = "=K2-L2"
    End With
    
    With MetricsRange.Columns(2)
        .Range("A1").Value = "DOS - Status"
        .Range("A1").Font.Bold = True
        .Offset(1, 0).Formula = "=K2-J2"
    End With

    With MetricsRange.Columns(3)
        .Range("A1").Value = "Status - Created"
        .Range("A1").Font.Bold = True
        .Offset(1, 0).Formula = "=J2-L2"
    End With
    
    With MetricsRange
        .Copy
        .PasteSpecial (xlPasteValues)
    End With
    
    Set MetricsRange = Nothing
    Set rSort = Nothing
End Function

Public Function AddDoneStatus()
    Dim StatusRange As Range
    Dim lLastRow As Long
    Dim rSource As Range
    
    Set rSource = GetSourceRange
    
    lLastRow = rSource.Rows.Count
    
    Set StatusRange = rSource.Columns("M")
    
    StatusRange.Formula = "=VLOOKUP($I2,reference!$A$2:$B$35,2,True)"
    
    With StatusRange
        .Copy
        .PasteSpecial (xlPasteValues)
    End With
    
    Set StatusRange = Nothing
    Set rSource = Nothing
End Function

Public Function OptimizeCodeBegin(ByVal sSheetName As String)
    Application.ScreenUpdating = False
    
    EventState = Application.EnableEvents
    Application.EnableEvents = False
    
    CalcState = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    PageBreakState = Sheets(sSheetName).DisplayPageBreaks

    Sheets(sSheetName).DisplayPageBreaks = False
End Function

Public Function OptimizeCodeEnd(sSheetName)
    Sheets(sSheetName).DisplayPageBreaks = PageBreakState

    Application.Calculation = CalcState
    Application.EnableEvents = EventState
    Application.ScreenUpdating = True
End Function
