Attribute VB_Name = "DotNetModule"
Public Enum SendEMAILFORMAT_Enum
    SendEMAILFORMATAUTO = 0
    SendEMAILFORMATTEXT = 1
    SendEMAILFORMATHTML = 2
    SendEMAILFORMATRTF = 3
End Enum

Public Type SendEMAIL_Type
    TOEMAILs_String As String
    CCEMAILs_String As String
    BCCEMAILs_String As String
    SUBJECT_String As String
    BODY_String As String
    BODYFILENAME_String As String
    FORMAT_SendEMAILFORMAT_EnumOptional As SendEMAILFORMAT_Enum
    FROMEMAIL_String As String 'PEN Server 2 <support@penguininc.com>
    SMTPUSER_StringOptional As String 'support@penguininc.com
    SMTPPWD_StringOptional As String 'penguin1234
    SMTPSERVER_StringOptional As String 'mail.penguininc.com
    SMTPPORT_StringOptional As String '25
    SMTPTIMEOUTMin_LongOptional As Long '2
    'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS
    SMTPSSL_IntegerOptional As Integer '0
    Result_StringOptional As String
    'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS
End Type

Public Function CallDotNetSendSMS(ByVal tmcsmsEmail As String, ByVal tmcsmsPassword As String, ByVal TOMOBILENO As String, ByVal MESSAGE As String) As Boolean
Dim Q As New QueryString
Dim ret As String
Dim Success As Boolean

ShowProgress "Opening ... Please wait"
DoEvents
'"EMAILADD=tinky@penguininc.com;PWD=penguin****;TOMOBILENO=;BODY=;"
    Q.Add "EMAILADD", tmcsmsEmail
    Q.Add "PWD", tmcsmsPassword
    Q.Add "TOMOBILENO", TOMOBILENO
    Q.Add "BODY", Left(MESSAGE, 160)
    
    ret = Q.Query
    ret = """" & ret & """"
    FMain.MousePointer = MousePointerConstants.vbHourglass
        ShellAndWait """" & App.Path & "\Modules\SendSMS\SendSMS.exe" & """" & " " & ret & ">""" & App.Path & "\Modules\SendSMS\SendSMS.log""", vbHide
        DoEvents
        Open App.Path & "\Modules\SendSMS\SendSMS.log" For Input Shared As #1
            Do While Not EOF(1)
                Line Input #1, temp
                If InStr(1, temp, "Sent Successfully") > 0 Then
                    Success = True
                    Exit Do
                Else
                    Success = False
                End If
                DoEvents
            Loop
        Close #1
    FMain.MousePointer = MousePointerConstants.vbDefault
HideProgress
        If Success Then
            MsgBox "Sent Successfully", vbInformation, App.Title
        Else
            MsgBox temp, vbCritical, App.Title '"Backup failed"
        End If

CallDotNetSendSMS = True
End Function

'Public Function CallDotNetSendEMAIL(ByVal TOEmailAddresses As String, ByVal Subject As String, ByVal Body As String, ByVal BodyFileName As String, ByVal FROMEmail As String, Optional SMTPUSER As String = "support@penguininc.com", Optional SMTPPWD As String = "penguin1234", Optional SMTPSERVER As String = "mail.penguininc.com", Optional SMTPPORT As String = "25") As Boolean
'Dim Q As New QueryString
'Dim ret As String
'Dim Success As Boolean
'Dim cmdStringREM As String
'
'ShowProgress "Sending Email ... Please wait"
'DoEvents
''"TOEMAIL=abhi@penguininc.com;SUB=GDS Stopped;BODY=Please;BFILENAME=D:\Collation_error.rtf;SMTPUSER=support@penguininc.com;FROMEMAIL=hhhh@email.com;SMTPPWD=penguin1234;SMTPSERVER=mail.penguininc.com;SMTPPORT=25;
'    Q.Add "TOEMAIL", TOEmailAddresses
'    Q.Add "SUB", Subject
'    Q.Add "BODY", Body
'    Q.Add "BFILENAME", BodyFileName
'    Q.Add "SMTPUSER", SMTPUSER
'    Q.Add "FROMEMAIL", FROMEmail
'    Q.Add "SMTPPWD", SMTPPWD
'    Q.Add "SMTPSERVER", SMTPSERVER
'    Q.Add "SMTPPORT", SMTPPORT
'
'    ret = Q.Query
'    ret = """" & ret & """"
'    If InStr(1, ret, Chr(13), vbTextCompare) > 0 Then
'        MsgBox "Please remove new line charactor", vbCritical
'        FMain.MousePointer = MousePointerConstants.vbDefault
'        HideProgress
'        Exit Function
'    End If
'    FMain.MousePointer = MousePointerConstants.vbHourglass
'        Open App.Path & "\Modules\SendEMAIL\BAT.BAT" For Output Shared As #1
'            cmdString = """" & App.Path & "\Modules\SendEMAIL\SendEMAIL.exe" & """" & " " & ret
'            cmdString = cmdString & ">""" & App.Path & "\Modules\SendEMAIL\SendEMAIL.log"""
'            Print #1, cmdString
'            cmdStringREM = "REM " & cmdString
'        Close #1
'        ShellAndWait """" & App.Path & "\Modules\SendEMAIL\BAT.BAT" & """", vbHide
'        DoEvents
'        Open App.Path & "\Modules\SendEMAIL\BAT.BAT" For Output Shared As #1
'            Print #1, cmdStringREM
'        Close #1
'
'        Open App.Path & "\Modules\SendEMAIL\SendEMAIL.log" For Input Shared As #1
'            Do While Not EOF(1)
'                Line Input #1, temp
'                If InStr(1, temp, "Sent Successfully") > 0 Then
'                    Success = True
'                    Exit Do
'                Else
'                    Success = False
'                End If
'            Loop
'        Close #1
'    FMain.MousePointer = MousePointerConstants.vbDefault
'HideProgress
'        If Success Then
'            'MsgBox "Sent Successfully", vbInformation
'        Else
'            MsgBox temp, vbCritical '"Backup failed"
'        End If
'
'CallDotNetSendEMAIL = True
'End Function

Public Function CallDotNetSendEMAIL(ByRef pSendEMAIL_Type As SendEMAIL_Type) As Boolean
Dim Q As New QueryString
Dim ret As String
Dim Success As Boolean
Dim cmdString As String
Dim cmdStringREM As String
Dim vSendEMAILFORMAT_String As String
'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
Dim vBATFilename_String As String
Dim vLOGFilename_String As String
'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS
Dim vDefaultSMTPUSER_String As String
Dim vDefaultSMTPPWD_String As String
Dim vDefaultSMTPSERVER_String As String
Dim vDefaultSMTPPORT_String As String
Dim vDefaultSMTPSSL_Integer As Integer
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS

ShowProgress "Sending Error to Penguin Support... Please wait"
DoEvents
''"TOEMAIL=abhi@penguininc.com;SUB=GDS Stopped;BODY=Please;BFILENAME=D:\Collation_error.rtf;SMTPUSER=support@penguininc.com;FROMEMAIL=hello <hhhh@email.com>;SMTPPWD=penguin1234;SMTPSERVER=mail.penguininc.com;SMTPPORT=25;
'"TOEMAIL=abhi@penguininc.com;CCMAIL=biji@penguininc.com;SUB=System Restarted [PEN Server 2];BODY=;BFILENAME=H:\UK_Project\PenAIR\SQL Server\Modules\SendEMAIL\StartupShutdown\Body.txt;FORMAT=TEXT;SMTPUSER=support@penguininc.com;FROMEMAIL=PEN Server 2 <support@penguininc.com>;SMTPPWD=penguin1234;SMTPSERVER=mail.penguininc.com;SMTPPORT=25;"
        
    'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS
    If InStr(1, pSendEMAIL_Type.FROMEMAIL_String, "@", vbTextCompare) = 0 Or Trim(pSendEMAIL_Type.SMTPUSER_StringOptional) = "" Then
        vDefaultSMTPUSER_String = INIRead(INIPenAIR_String, "DefaultSMTP", "SMTPUSER", "noreply@penguininc.com")
        vDefaultSMTPPWD_String = INIRead(INIPenAIR_String, "DefaultSMTP", "SMTPPWD", "s@pth@m@s!")
        vDefaultSMTPSERVER_String = INIRead(INIPenAIR_String, "DefaultSMTP", "SMTPSERVER", "mail.penguininc.com")
        vDefaultSMTPPORT_String = INIRead(INIPenAIR_String, "DefaultSMTP", "SMTPPORT", "587")
        vDefaultSMTPSSL_Integer = Val(INIRead(INIPenAIR_String, "DefaultSMTP", "SMTPSSL", "1"))
    End If
    'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS
        
    If pSendEMAIL_Type.BODYFILENAME_String <> "" Then
        pSendEMAIL_Type.BODY_String = ""
    End If
    
    If InStr(1, pSendEMAIL_Type.FROMEMAIL_String, "@", vbTextCompare) = 0 Then
        'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS
        'pSendEMAIL_Type.FROMEMAIL_String = pSendEMAIL_Type.FROMEMAIL_String & " <support@penguininc.com>"
        pSendEMAIL_Type.FROMEMAIL_String = pSendEMAIL_Type.FROMEMAIL_String & " <" & vDefaultSMTPUSER_String & ">"
        'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS
    End If
    
    Select Case pSendEMAIL_Type.FORMAT_SendEMAILFORMAT_EnumOptional
        Case SendEMAILFORMATTEXT
            vSendEMAILFORMAT_String = "TEXT"
        Case SendEMAILFORMATHTML
            vSendEMAILFORMAT_String = "HTML"
        Case SendEMAILFORMATRTF
            vSendEMAILFORMAT_String = "RTF"
        Case Else
            vSendEMAILFORMAT_String = "AUTO"
    End Select
    
    'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS
    'If Trim(pSendEMAIL_Type.SMTPUSER_StringOptional) = "" Then
    '    pSendEMAIL_Type.SMTPUSER_StringOptional = "support@penguininc.com"
    'End If
    '
    'If Trim(pSendEMAIL_Type.SMTPPWD_StringOptional) = "" Then
    '    pSendEMAIL_Type.SMTPPWD_StringOptional = "penguin1234"
    'End If
    '
    'If Trim(pSendEMAIL_Type.SMTPSERVER_StringOptional) = "" Then
    '    pSendEMAIL_Type.SMTPSERVER_StringOptional = "mail.penguininc.com"
    'End If
    '
    'If Trim(pSendEMAIL_Type.SMTPPORT_StringOptional) = "" Then
    '    pSendEMAIL_Type.SMTPPORT_StringOptional = "25"
    'End If
    If Trim(pSendEMAIL_Type.SMTPUSER_StringOptional) = "" Then
        pSendEMAIL_Type.SMTPUSER_StringOptional = vDefaultSMTPUSER_String
        pSendEMAIL_Type.SMTPPWD_StringOptional = vDefaultSMTPPWD_String
        pSendEMAIL_Type.SMTPSERVER_StringOptional = vDefaultSMTPSERVER_String
        pSendEMAIL_Type.SMTPPORT_StringOptional = vDefaultSMTPPORT_String
        pSendEMAIL_Type.SMTPSSL_IntegerOptional = vDefaultSMTPSSL_Integer
    End If
    pSendEMAIL_Type.SMTPPWD_StringOptional = Replace(pSendEMAIL_Type.SMTPPWD_StringOptional, "%", "%%", 1, , vbTextCompare)
    'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS
    
    'If Trim(pSendEMAIL_Type.CCEMAILs_String) <> "" Then
    '    pSendEMAIL_Type.TOEMAILs_String = pSendEMAIL_Type.TOEMAILs_String & "," & pSendEMAIL_Type.CCEMAILs_String
    'End If
    
    If pSendEMAIL_Type.SMTPTIMEOUTMin_LongOptional = 0 Then
        'by Abhi on 26-Nov-2014 for caseid 4774 Sending emails from PenAIR as per Gmail Server timeouts recommend 5 mins
        'pSendEMAIL_Type.SMTPTIMEOUTMin_LongOptional = 2
        pSendEMAIL_Type.SMTPTIMEOUTMin_LongOptional = 5
        'by Abhi on 26-Nov-2014 for caseid 4774 Sending emails from PenAIR as per Gmail Server timeouts recommend 5 mins
    End If
    
    Q.Add "TOEMAIL", pSendEMAIL_Type.TOEMAILs_String
    Q.Add "CCEMAIL", pSendEMAIL_Type.CCEMAILs_String
    Q.Add "BCC", pSendEMAIL_Type.BCCEMAILs_String
    Q.Add "SUB", pSendEMAIL_Type.SUBJECT_String
    Q.Add "BODY", pSendEMAIL_Type.BODY_String
    Q.Add "BFILENAME", pSendEMAIL_Type.BODYFILENAME_String
    Q.Add "FORMAT", vSendEMAILFORMAT_String
    Q.Add "FROMEMAIL", pSendEMAIL_Type.FROMEMAIL_String
    Q.Add "SMTPUSER", pSendEMAIL_Type.SMTPUSER_StringOptional
    Q.Add "SMTPPWD", pSendEMAIL_Type.SMTPPWD_StringOptional
    Q.Add "SMTPSERVER", pSendEMAIL_Type.SMTPSERVER_StringOptional
    Q.Add "SMTPPORT", pSendEMAIL_Type.SMTPPORT_StringOptional
    Q.Add "SMTPTIMEOUTMin", str(pSendEMAIL_Type.SMTPTIMEOUTMin_LongOptional)
    'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS
    Q.Add "SMTPSSL", Trim(str(pSendEMAIL_Type.SMTPSSL_IntegerOptional))
    Q.Add "HelpURL", ""
    Q.Add "MSGIDADDLREF", App.EXEName & ".exe"
    'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com enable for email alerts in PenAIR and PenGDS
    
    ret = Q.Query
    ret = """" & ret & """"
    If InStr(1, ret, Chr(13), vbTextCompare) > 0 Then
        MsgBox "Please remove new line charactor", vbCritical, App.Title
        FMain.MousePointer = MousePointerConstants.vbDefault
        HideProgress
        Exit Function
    End If
    FMain.MousePointer = MousePointerConstants.vbHourglass
        'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
        'Open App.Path & "\Modules\SendEMAIL\BAT.BAT" For Output Shared As #1
        '    cmdString = """" & App.Path & "\Modules\SendEMAIL\SendEMAIL.exe" & """" & " " & ret
        '    cmdString = cmdString & ">""" & App.Path & "\Modules\SendEMAIL\SendEMAIL.log"""
        '    Print #1, cmdString
        '    cmdStringREM = "REM " & cmdString
        'Close #1
        'ShellAndWait """" & App.Path & "\Modules\SendEMAIL\BAT.BAT" & """", vbHide
        'DoEvents
        'Open App.Path & "\Modules\SendEMAIL\BAT.BAT" For Output Shared As #1
        '    Print #1, cmdStringREM
        'Close #1
        '
        'Open App.Path & "\Modules\SendEMAIL\SendEMAIL.log" For Input Shared As #1
        vBATFilename_String = PENTempName(TempNameTypeFile, ".bat")
        vLOGFilename_String = PENTempName(TempNameTypeFile, ".log")
        Open vBATFilename_String For Output Shared As #1
            cmdString = """" & App.Path & "\Modules\SendEMAIL\SendEMAIL.exe" & """" & " " & ret
            cmdString = cmdString & ">""" & vLOGFilename_String & """"
            Print #1, cmdString
            cmdStringREM = "REM " & cmdString
        Close #1
        ShellAndWait """" & vBATFilename_String & """", vbHide
        DoEvents
        Open vBATFilename_String For Output Shared As #1
            Print #1, cmdStringREM
        Close #1
        
        Open vLOGFilename_String For Input Shared As #1
        'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
            Do While Not EOF(1)
                Line Input #1, temp
                If InStr(1, temp, "Sent Successfully") > 0 Then
                    Success = True
                    Exit Do
                Else
                    Success = False
                End If
                DoEvents
            Loop
        Close #1
    FMain.MousePointer = MousePointerConstants.vbDefault
HideProgress
        If Success Then
            'MsgBox "Sent Successfully", vbInformation
        Else
            MsgBox temp, vbCritical, App.Title '"Backup failed"
        End If

CallDotNetSendEMAIL = Success
End Function


