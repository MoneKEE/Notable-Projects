Attribute VB_Name = "modPivotTables"
'*******************************************
'Revision History
'Date       Developer       Description
'29-Sep-16  M Gore          Initial Version
'*******************************************

Option Compare Text

Private Const mModuleName As String = "modPivotTables"

Public FromDate As Variant
Public ToDate As Variant

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function InitCompletedCasesSheet()
    Dim Header1 As String
    Dim Header2 As String
    Dim Header3 As String
    
    Header1 = "Experian Extract"
    Header2 = "# of cases in ""Done"" Status by date selected (not date of service)"
    Header3 = "based on cases statused from " & FromDate & " to " & ToDate

    
    
End Function

Public Function Initialize()
    FromDate = Format(DateTime.DateAdd("d", -10, Now), "M/dd")
    ToDate = Format(Now, "M/dd ham/pm")
End Function

Private Function FieldVisibility(ByVal shtName As String, ByVal pvName As String)
    Dim i As Integer
    Dim PVFields As Variant
    Dim FldCount As Integer
    
    Select Case pvName
        Case "Completed"
            PVFields = Array("Status Set By", "Dept", "Time Stamp", _
                        "DoneStatus", "Count of AccountNumber")
            FldCount = 5
        Case "Touched"
            PVFields = Array("Status Set By", "Dept", "Time Stamp", _
                        "DoneStatus", "Count of AccountNumber")
            FldCount = 5
        Case "DoneDetail"
            PVFields = Array("Status Set By", "Dept", "Status", _
                        "DoneStatus", "Count of AccountNumber")
            FldCount = 5
        Case "UndoneDetail"
            PVFields = Array("Status Set By", "Dept", "Status", _
                        "DoneStatus", "Count of AccountNumber")
            FldCount = 5
        Case "ANRPayor"
            PVFields = Array("Primary Insurance", "Status Set By", "Dept", _
                        "Status", "Count of AccountNumber")
            FldCount = 5
        Case "Leadtime"
            PVFields = Array("Status Set By", "Dept", "DoneStatus", _
                        "Values", "Average of leadtime", "Average of DOS status date")
            FldCount = 6
        Case "StatAccount"
            PVFields = Array("AccountNumber", "Status Set By", "Dept", _
                        "DoneStatus")
            FldCount = 4
    End Select
    
    With Sheets(shtName).PivotTables(pvName)
        If .VisibleFields.Count <> FldCount Then
            For i = 1 To .VisibleFields.Count
                If .VisibleFields(i) <> PVFields(i - 1) Then
                    Debug.Print "Found It"
                End If
            Next
        Else
        
        End If
    End With
    
End Function
