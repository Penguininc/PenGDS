Attribute VB_Name = "MainModules"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public connstr As String
Public mPathComp As String

Public dbCompany As ADODB.Connection
Public rsCompany As New ADODB.Recordset

Public fsObj As New FileSystemObject

'Deleting the Files
Private Type SHFILEOPTSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As Long
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" _
  Alias "SHFileOperationA" (lpFileOp As SHFILEOPTSTRUCT) As Long

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40

Public mLogin As Boolean
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public vgoPath As New FileSystemObject
Public i As Long, iTimer As Long
Public DirName As String
Public Email As String
Public INIPenAIR_String As String
Public Startup_Boolean As Boolean
'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
Public INIPen_String As String
Public CONAME_String As String
Public Declare Function FindWindow Lib "user32" Alias _
         "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName _
         As Any) As Long
Public THandle As Long
'Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9
Public Const SW_SHOWNORMAL = 1
Public Const SW_HIDE = 0
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
  ByVal nCmdShow As Long) As Long
'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray

Public Sub Main()
    If App.PrevInstance = True Then
        End
    End If
    Startup_Boolean = True
'    If INIRead(App.Path & "\PenSoft.ini", "PenGDS", "GDS", "OFF") = "ON" Then
'        End
'    End If
    With FMain
        'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
        .cmdExit.Enabled = False
        .cmdApply.Enabled = False
        .cmdStart.Enabled = False
        .cmdStop.Enabled = False
        'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
        .ShowAllchameleonButton.Enabled = False
        .HideAllchameleonButton.Enabled = False
        'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
        DoEvents
        'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
        .Dir1.Path = App.Path
        For i = 0 To .Dir1.ListCount - 1
            DirName = .Dir1.List(i)
            INIPenAIR_String = DirName & "\PenAIR.ini"
            If UCase(Dir(INIPenAIR_String)) = "" Then
                INIPenAIR_String = DirName & "\PenSoft.ini"
            End If
            If Dir(INIPenAIR_String) <> "" Then
                'DirName = Replace(DirName, App.Path & "\", "")
                .ListView.ListItems.Add , , Replace(DirName, App.Path & "\", "")
                .ListView.ListItems(.ListView.ListItems.Count).ListSubItems.Add , , "Checking..."
    '            If INIRead(DirName & "\PenSoft.ini", "PenGDS", "GDS", "OFF") = "ON" Then
    '                .ListView.ListItems(.ListView.ListItems.Count).Checked = True
    '                '.ListView.ListItems(.ListView.ListItems.Count).Bold = True
    '                .ListView.ListItems(.ListView.ListItems.Count).ForeColor = &HC000&
    '                .ListView.ListItems(.ListView.ListItems.Count).ListSubItems.Add , , "Started"
    '                '.ListView.ListItems(.ListView.ListItems.Count).ListSubItems(.ListView.ListItems(.ListView.ListItems.Count).ListSubItems.Count).Bold = True
    '                .ListView.ListItems(.ListView.ListItems.Count).ListSubItems(.ListView.ListItems(.ListView.ListItems.Count).ListSubItems.Count).ForeColor = &HC000&
    '            Else
    '                .ListView.ListItems(.ListView.ListItems.Count).ListSubItems.Add , , "Stopped"
    '                '.ListView.ListItems(.ListView.ListItems.Count).ListSubItems(.ListView.ListItems(.ListView.ListItems.Count).ListSubItems.Count).Bold = True
    '                .ListView.ListItems(.ListView.ListItems.Count).ListSubItems(.ListView.ListItems(.ListView.ListItems.Count).ListSubItems.Count).ForeColor = &HC0&
    '            End If
                .ListView.ListItems(.ListView.ListItems.Count).ListSubItems.Add , , DirName
                'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
                INIPen_String = DirName & "\Pen.ini"
                CONAME_String = INIRead(INIPen_String, "General", "CONAME", "")
                If Trim(CONAME_String) <> "" Then
                    CONAME_String = "PenGDS Interface [" & CONAME_String & "]"
                End If
                .ListView.ListItems(.ListView.ListItems.Count).ListSubItems.Add , , CONAME_String
                'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
            End If
            'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
            '.ListView.ColumnHeaders(3).Width = 0
            'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
        Next
        'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
        .ListView.ColumnHeaders(3).Width = 10000
        .ListView.ColumnHeaders(4).Width = 10000
        'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
        For i = 1 To .ListView.ListItems.Count
            With .ListView.ListItems(i)
                If INIRead(App.Path & "\PenGDS Server.ini", "Clients", UCase(.Text), False) = True Then
                    .Checked = True
                End If
            End With
        Next
        .chkAutoStart = INIRead(App.Path & "\PenGDS Server.ini", "General", "AutoStart", 0)
        .Show
        DoEvents
    End With
    
    DoEvents
    If FMain.chkAutoStart = vbChecked Then
        FMain.cmdApply_Click
    End If
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    If FMain.ListView.ListItems.Count > 0 Then
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
        FMain.ListView.ListItems(1).Selected = False
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    End If
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
    Call FMain.tmrUpload_Timer
    DoEvents
    FMain.cmdExit.Enabled = True
    FMain.cmdApply.Enabled = True
    FMain.cmdStart.Enabled = True
    FMain.cmdStop.Enabled = True
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    FMain.ShowAllchameleonButton.Enabled = True
    FMain.HideAllchameleonButton.Enabled = True
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    DoEvents
    Startup_Boolean = False
    FMain.tmrUpload.Enabled = True
    'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
End Sub

Public Function INIRead(sINIFile As String, sSection As String, sKey As String, sDefault As String) As String
    Dim sTemp As String * 256
    Dim nLength As Integer
    sTemp = Space$(256)
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)
    INIRead = Left$(sTemp, nLength)
End Function
Public Sub INIWrite(sINIFile As String, sSection As String, sKey As String, sValue As String)
    Dim N As Integer
    Dim sTemp As String
    sTemp = sValue
    'Replace any CR/CF characters with spaces
    For N = 1 To Len(sValue)
        If Mid$(sValue, N, 1) = vbCr Or Mid$(sValue, N, 1) = vbLf Then Mid$(sValue, N) = " "
    Next N
    N = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)
End Sub


Public Function GetFileNameFromPath(Path As String) As String
Dim temp, Count As Long
temp = Split(Path, "\")
Count = UBound(temp)
If Count >= 0 Then
    GetFileNameFromPath = temp(Count)
End If
End Function

Public Function PathExists(ByVal vPath As String) As Boolean
On Error GoTo myErr
    DoEvents
    If Dir(vPath, vbDirectory) = "" Then
        PathExists = False
    Else
        PathExists = True
    End If

Exit Function
myErr:
On Error GoTo myErr2
    DoEvents
    FMain.File1.Path = vPath
    PathExists = True
Exit Function
myErr2:
    DoEvents
    PathExists = False
                                                    'On Error GoTo myErr
                                                    '    If Dir(vpath, vbDirectory) = "" Then
                                                    '        PathExists = False
                                                    '    Else
                                                    '        PathExists = True
                                                    '    End If
                                                    'Exit Function
                                                    'myErr:
                                                    '    PathExists = vgoPath.FolderExists(vpath)
End Function

'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
Public Function TimeFormat12HRS(ByVal pTime_String As String) As String
    TimeFormat12HRS = Format(pTime_String, "HH:MM:SS AMPM")
End Function
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file

'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
Public Function BringToFront(ByVal pWindowName_String As String, Optional ByVal pShowMsgBox_Boolean As Boolean = True)

    THandle = FindWindow(vbEmpty, pWindowName_String)
    FindWindowLabel = THandle
    If THandle = 0 Then
        If pShowMsgBox_Boolean = True Then
            MsgBox "'" & pWindowName_String & "' is not running", vbCritical, App.Title
        End If
        Exit Function
    End If
   
    'ret = ShowWindow(THandle, SW_SHOW)
    ret = ShowWindow(THandle, SW_SHOWNORMAL)
    ret = ShowWindow(THandle, SW_RESTORE)

End Function
'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray

'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
Public Function BringToBackorHide(ByVal pWindowName_String As String, Optional ByVal pShowMsgBox_Boolean As Boolean = True)

    THandle = FindWindow(vbEmpty, pWindowName_String)
    FindWindowLabel = THandle
    If THandle = 0 Then
        If pShowMsgBox_Boolean = True Then
            MsgBox "'" & pWindowName_String & "' is not running", vbCritical, App.Title
        End If
        Exit Function
    End If
   
    ret = ShowWindow(THandle, SW_HIDE)

End Function
'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray

