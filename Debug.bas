Attribute VB_Name = "Debug"
'by Abhi on 18-May-2011 for caseid 1757 Added Events for GDSAuto
Public Function EventLog(ByVal pErrorMessage_String As String, Optional ByVal pAtNewLine_Boolean As Boolean = False)
On Error GoTo PENErr
Dim vFile As Long
Static vErrorMessage_String As String
    
    pErrorMessage_String = Replace(pErrorMessage_String, vbCrLf, "")
    vErrorMessage_String = vErrorMessage_String & pErrorMessage_String

    vFile = FreeFile
    Open App.Path & "\Logs\PenGDS Events.log" For Append Shared As #vFile
        If pAtNewLine_Boolean Then
            Print #vFile, ""
        End If
        Print #vFile, DateFormat(Date); " "; Format(time, "HH:MM:SS AMPM"); "(" & TimeFormat(time) & ")"; " : "; vErrorMessage_String
    Close #vFile
    vErrorMessage_String = ""
Exit Function

PENErr:
    Close #vFile
    Resume Next
    Exit Function
End Function

'by Abhi on 01-Nov-2012 for caseid 2368 Sabre pnr missing
Public Function EventLogSabre(ByVal pErrorMessage_String As String, Optional ByVal pAtNewLine_Boolean As Boolean = False)
On Error GoTo PENErr
Dim vFile As Long
Static vErrorMessage_String As String
    
    'by Abhi on 02-Feb-2017 for caseid 7154 Remove log of PenGDS Events Sabre
    Exit Function
    'by Abhi on 02-Feb-2017 for caseid 7154 Remove log of PenGDS Events Sabre
    
    pErrorMessage_String = Replace(pErrorMessage_String, vbCrLf, "")
    vErrorMessage_String = vErrorMessage_String & pErrorMessage_String

    vFile = FreeFile
    'by Abhi on 04-Mar-2013 for caseid 2958 PenGDS Sabre pnr file not refeshing
    'Open App.Path & "\PenGDS Events Sabre.log" For Append Shared As #vFile
    Open App.Path & "\Logs\PenGDS Events Sabre " & Format(Date, "YYYYMMDD") & ".log" For Append Shared As #vFile
        If pAtNewLine_Boolean Then
            Print #vFile, ""
        End If
        Print #vFile, DateFormat(Date); " "; Format(time, "HH:MM:SS AMPM"); "(" & TimeFormat(time) & ")"; " : "; vErrorMessage_String
    Close #vFile
    vErrorMessage_String = ""
Exit Function

PENErr:
    Close #vFile
    Resume Next
    Exit Function
End Function

'by Abhi on 18-May-2011 for caseid 1757 Added Events for GDSAuto
Public Function EventLogDelete()
    If Dir(App.Path & "\Logs\PenGDS Events.log") <> "" Then
        If FileLen(App.Path & "\Logs\PenGDS Events.log") > 102400000 Then '10 MB
            Kill App.Path & "\Logs\PenGDS Events.log"
        End If
    End If
    'by Abhi on 01-Nov-2012 for caseid 2368 Sabre pnr missing
    Call EventLogDeleteSabre
End Function

'by Abhi on 01-Nov-2012 for caseid 2368 Sabre pnr missing
Public Function EventLogDeleteSabre()
    'by Abhi on 04-Mar-2013 for caseid 2958 PenGDS Sabre pnr file not refeshing
    Exit Function
    
    If Dir(App.Path & "\Logs\PenGDS Events Sabre.log") <> "" Then
        If FileLen(App.Path & "\Logs\PenGDS Events Sabre.log") > 102400000 Then '10 MB
            Kill App.Path & "\Logs\PenGDS Events Sabre.log"
        End If
    End If
End Function

Public Function ShowFormName(ByVal pForm As Form) As Boolean
On Error GoTo PENErr
    Dim vFrm As New ShowFormNameForm
    Call vFrm.ShowME(pForm)

Exit Function

PENErr:
    MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title & " [ShowFormName]"
    Exit Function
End Function

Public Function GetFormName(ByVal pForm As Form) As String
    GetFormName = pForm.Name
End Function

