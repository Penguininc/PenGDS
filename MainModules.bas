Attribute VB_Name = "MainModules"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public connstr As String
Public mPathComp As String
Public CHEAD As String

Public dbCompany As ADODB.Connection
Public rsCompany As New ADODB.Recordset
'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
Public CnExcel As New ADODB.Connection

Public fsObj As New FileSystemObject
Public vgsBackend As String
Public vgsProvider As String
Public vgbStandardSecurity As Boolean
Public vgsServer As String
Public vgsUID As String
Public vgsPWD As String
Public vgsCompany As String
Public vglServers As Long
Public vgliSer As Long
Public vglDatabases As Long
Public vgsDatabase As String

Public rsSelect As New ADODB.Recordset

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

''KillProcess - Terminate any application

Private Type LUID
   lowpart As Long
   highpart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As _
    Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle _
    As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias _
    "LookupPrivilegeValueA" (ByVal lpSystemName As String, _
    ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal _
    TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
    NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
    PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As _
    Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As _
    Long, ByVal uExitCode As Long) As Long
''KillProcess - Terminate any application

Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Const SYNCHRONIZE = &H100000
Public Const INFINITE = -1&


Public vgoPath As New FileSystemObject

Public ClientName As String
Public Host As String
Public Port As String
Public JustConnected As Boolean
Public Monitor As Boolean
Public SendComplete As Boolean

'by Abhi on 20-Oct-2009 for Penguin Encryption
Public Enum EncryptDecrypt_Enum
    Encrypt = 0
    Decrypt = 1
End Enum
Public Encrypted_Boolean As Boolean
Public INIPenAIR_String As String
Public EncryptedINI_Boolean As Boolean
Public PNAME_String As String

Private Declare Function SetLocaleInfo Lib "kernel32" Alias _
    "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As _
    Long, ByVal lpLCData As String) As Boolean
Private Const LOCALE_SSHORTDATE = &H1F
Private Declare Function GetSystemDefaultLCID Lib "kernel32" _
    () As Long
Private Declare Function PostMessage Lib "user32" Alias _
    "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public PENFAREPNO_Long As Long
'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
Public PENLINEID_String As String
'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
Public LUFPNR_String As String

'by Abhi on 12-Mar-2012 for caseid 1652 PenGDS Permission denied added file closed checking
Public Enum FileStatus_Enum
    FileStatusClosed = 0
    FileStatusOpened = 1
    FileStatusNotFound = 2
End Enum

Public TextStream_TextStream As TextStream
Public ReadAll_String As String
'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
Public AnyDataChanged_Boolean As Boolean
 
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
 (ByVal hOpen As Long, ByVal sUrl As String, ByVal sHeaders As String, _
 ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Public Declare Function InternetReadFile Lib "wininet.dll" _
 (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
 lNumberOfBytesRead As Long) As Integer
 
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer
 
Public Declare Function GetDesktopWindow Lib "user32" () As Long
 
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

Public Enum HelpServiceURLType_Enum
    HelpServiceURLType_Functional
    HelpServiceURLType_Technical
End Enum

Public mUserLID As String
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
Public NoofPermissionDenied As Long
Public PermissionDenied As Long
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
'by Abhi on 01-Jul-2014 for caseid 4222 Exchange value picking logic-Hardcoded currency "GBP" should be replaced with company currency for Galileo
Public COMCID_String As String
'by Abhi on 01-Jul-2014 for caseid 4222 Exchange value picking logic-Hardcoded currency "GBP" should be replaced with company currency for Galileo
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
Public GDSDeadlockRETRY_Integer As Integer
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
Public PENErr_BeginTrans As Boolean
'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
Public SendERROREmailPreviousErrNumber_String As String
Public SendERROREmailPreviousMess_String As String
Public SendERROREmailStart
'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
Public PENAIRTKTPNO_Long As Long
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
'by Abhi on 11-Jul-2015 for caseid 5393 PenAIR Run time error 76 path not found
Public UsersTemporaryFolder_String As String
Public PenEMAILTemporaryFolder_String As String
Public PenEMAILFileSystemObject As New FileSystemObject
Public Enum TempNameTypeEnum
    TempNameTypeFile = VbFileAttribute.vbNormal
    TempNameTypeDirectory = VbFileAttribute.vbDirectory
End Enum
'by Abhi on 11-Jul-2015 for caseid 5393 PenAIR Run time error 76 path not found
'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
Public ErrDetails_String As String
'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
Public GIT_CUSTOMERUSERCODE_String As String
'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
Public GIT_PENWAIT_String As String
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
Public GIT_PENLINEBRID_String As String
'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_SETTINGCHANGE = &H1A
Private Const LOCALE_SYSTEM_DEFAULT = &H400
Public DateSyDefault As String     '*   regional Setting

Public Const scUserAgent = "VB OpenUrl"
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_FLAG_RELOAD = &H80000000
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
Public ELcls As New ClassFormSeting
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
Public mFormName As String
Public mTxtCallReturn As String
Public mSelectedDate_String As String
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date

Public Sub Main()
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
Dim vvlAns As VbMsgBoxResult
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
    If App.PrevInstance = True Then
        End
    End If
    
    'by Abhi on 11-Jul-2015 for caseid 5393 PenAIR Run time error 76 path not found
    Call CreateUsersTemporaryFolder(True)
    'by Abhi on 11-Jul-2015 for caseid 5393 PenAIR Run time error 76 path not found
    
    If str(Date) <> Format(Date, "DD/MMM/YYYY") Or InStr(str(Date), "/") = 0 Then
        vvlAns = MsgBox("Date format is not compatible. Do you want to change?" & vbCrLf & vbCrLf & "Current date format is " & Date & vbCrLf & App.ProductName & " date format is " & Format(Date, "dd") & "/" & Format(Date, "MMM") & "/" & Format(Date, "yyyy"), vbOKCancel + vbCritical, App.Title)
        If vvlAns = vbOK Then
            Call ChangeDateRegionalSettings
        Else
            End
        End If
    End If
    
    PNAME_String = INIRead(App.Path & "\Pen.ini", "General", "PNAME", "PenAIR") '"PenAIR"
    INIPenAIR_String = App.Path & "\" & PNAME_String & ".ini"
    EncryptedINI_Boolean = True
    If UCase(Dir(INIPenAIR_String)) = "" Then
        INIPenAIR_String = App.Path & "\PenSoft.ini"
        EncryptedINI_Boolean = False
    End If
    
    If (UCase(Dir(App.Path & "\EndSQL.txt")) = UCase("EndSQL.txt") Or UCase(Dir(App.Path & "\EndGDS.txt")) = UCase("EndGDS.txt")) Then
        'by Abhi on 22-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module
        'Open (App.Path & "\ENDGDS.TXT") For Random As #1
        'by Abhi on 25-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module is Append
        'Open (App.Path & "\ENDGDS.TXT") For Output Shared As #1
        Open (App.Path & "\ENDGDS.TXT") For Append Shared As #1
            'by Abhi on 22-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module
            Print #1, "Updating... MainModules"
        Close #1
        INIWrite INIPenAIR_String, "PenGDS", "GDS", "OFF"
        End
    End If
    Monitor = INIRead(INIPenAIR_String, "PenGDS", "Monitor", 1)
    ClientName = INIRead(INIPenAIR_String, "PenGDS", "Name", "")
    Host = INIRead(INIPenAIR_String, "PenGDS", "Host", "penmonitor.dyndns.org")
    Port = INIRead(INIPenAIR_String, "PenGDS", "Port", "6230")
    vgsDatabase = INIRead(INIPenAIR_String, "General", "Database", "PENDEMO")
    
    mUserLID = "EN"
    
    If Trim(ClientName) = "" And Monitor = True Then
        FRegisterForm.Show 1
        If Trim(ClientName) = "" Then
            End
        End If
    End If
    
    If INIRead(INIPenAIR_String, "PenGDS", "GDS", "OFF") = "ON" Then
        'End
        If fsObj.FileExists(App.Path & "\_UploadingSQL_") = True Then
            'MsgBox "PenGDS is already running...", vbInformation, App.Title & " [" & vgsDatabase & "]"
            MsgBox App.ProductName & " is already running...", vbInformation, App.Title & " [" & vgsDatabase & "]"
            End
        End If
        'ShowProgress "PenGDS is preparing to open ... Please wait"
        ShowProgress App.ProductName & " is preparing to open ... Please wait"
        'by Abhi on 22-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module
        'Open (App.Path & "\ENDGDS.TXT") For Random As #1
        'by Abhi on 25-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module is Append
        'Open (App.Path & "\ENDGDS.TXT") For Output Shared As #1
        Open (App.Path & "\ENDGDS.TXT") For Append Shared As #1
            'by Abhi on 22-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module
            Print #1, "PenGDS is preparing to open ... Please wait... MainModules"
        Close #1
        Sleep 5000
        fsObj.DeleteFile (App.Path & "\ENDGDS.TXT"), True
        HideProgress
    End If
    INIWrite INIPenAIR_String, "PenGDS", "GDS", "ON"
    
    'by Abhi on 23-Sep-2013 for caseid 3393 All the error log and event logs should be in the directory "Logs"
    If Dir(App.Path & "\Logs", vbDirectory) = "" Then
        MkDir App.Path & "\Logs"
    End If
    
    'by Abhi on 21-Jan-2013 for caseid 2831 PenGDS File not found log
    If Dir(App.Path & "\PenAIR Events.log") <> "" Then
        Name App.Path & "\PenAIR Events.log" As App.Path & "\PenGDS Events.log"
    End If
    'by Abhi on 21-Jan-2013 for caseid 2831 PenGDS File not found log
    If Dir(App.Path & "\PenAIR Events Sabre.log") <> "" Then
        Name App.Path & "\PenAIR Events Sabre.log" As App.Path & "\PenGDS Events Sabre.log"
    End If
    
    Set dbCompany = New ADODB.Connection
    vgsBackend = INIRead(INIPenAIR_String, "General", "Backend", "SQL Server")
    vgsProvider = INIRead(INIPenAIR_String, "General", "Provider", "")
    vgsServer = INIRead(INIPenAIR_String, "General", "Server", "(Local)")
    vgsCompany = INIRead(INIPenAIR_String, "General", "Company", "PENCOMPANY")
    vgbStandardSecurity = Val(INIRead(INIPenAIR_String, "General", "StandardSecurity", "1"))
    vgsUID = INIRead(INIPenAIR_String, "General", "UID", "sa")
    'by Abhi on 20-Oct-2009 for Penguin Encryption
    If EncryptedINI_Boolean Then
        vgsUID = PENEncryptDecrypt(Decrypt, LCase(PNAME_String), vgsUID)
    End If
    vgsPWD = INIRead(INIPenAIR_String, "General", "PWD", "sa")
    'by Abhi on 20-Oct-2009 for Penguin Encryption
    If EncryptedINI_Boolean Then
        vgsPWD = PENEncryptDecrypt(Decrypt, LCase(PNAME_String), vgsPWD)
    End If
    vglServers = Val(INIRead(INIPenAIR_String, "General", "Servers", "0"))
    vglDatabases = Val(INIRead(INIPenAIR_String, "General", "Databases", "0"))
    
    If vglServers > 0 Then
        FServers.Caption = "PenAIR Servers"
        FServers.lstServers.Clear
        FServers.lstServers.AddItem vgsServer
        For vgliSer = 1 To vglServers
             FServers.lstServers.AddItem INIRead(INIPenAIR_String, "General", "Server" & vgliSer, "")
            DoEvents
        Next
        FServers.lstServers.ListIndex = 0
        FServers.Show 1
    End If
    
    If vglDatabases > 0 Then
        FServers.Caption = "PenAIR Databases"
        FServers.lstServers.Clear
        FServers.lstServers.AddItem vgsDatabase
        For vgliSer = 1 To vglDatabases
            FServers.lstServers.AddItem INIRead(INIPenAIR_String, "General", "Database" & vgliSer, "")
            DoEvents
        Next
        FServers.lstServers.ListIndex = 0
        FServers.Show 1
    End If
    
    'by Abhi on 16-Mar-2010 for SQL Server 2005
    If vgsBackend = "SQL Server 2005" Then
        vgsBackend = "SQL Native Client"
    'by Abhi on 31-May-2012 for caseid 2290 SQL Server 2008
    ElseIf vgsBackend = "SQL Server 2008" Then
        vgsBackend = "SQL Server Native Client 10.0"
    'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
    ElseIf vgsBackend = "SQL Server 2012" Then
        vgsBackend = "SQL Server Native Client 11.0"
    'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
    End If
    
    
    'mPathComp = vgsCompany 'App.Path & "\Company.dat"
    mPathComp = vgsDatabase '"PENCOMPANY"
    
    With dbCompany
        .CursorLocation = adUseClient
        '.CommandTimeout = 4
        'by Abhi on 07-Dec-2009 for Timeout Expired on Ariana
        .CommandTimeout = 300 '30
    End With
    
    'DBEngine.RegisterDatabase "PenSoftCompany", "Microsoft Access Driver (*.mdb)", True, "DBQ=" & mPathComp & ";PWD=penac"
    'DBEngine.RegisterDatabase "PenSoftCompany", vgsBackend, True, "Description=PenSoftCompany" & vbCr & "SERVER=PenAIR" & vbCr & "DATABASE=" & mPathComp & vbCr & "Network=DBMSSOCN" & vbCr & "Address=" & vgsServer '& vbCr '& "Trusted_Connection=Yes" & vbCr & "UID=" & vgsUID & vbCr & "PWD=" & vgsPWD & vbCr
    If Trim(vgsProvider) = "" Then
        DBEngine.RegisterDatabase "PenSoftCompany", "SQL Server", True, "Description=PenSoftCompany" & vbCr & "SERVER=" & vgsServer & vbCr & "DATABASE=" & mPathComp
        dbCompany.Open "PenSoftCompany", vgsUID, vgsPWD
    Else
        dbCompany.Open GetConnectionString(vgsBackend, vgsProvider, vgsServer, vgsUID, vgsPWD, mPathComp, vgbStandardSecurity)
    End If
    
    'by Abhi on 20-Oct-2009 for Penguin Encryption
    If PenguinIntegrity() = False Then
        MsgBox "Penguin Integrity failed. Please contact Penguin Support.", vbCritical, App.Title & " [Penguin Integrity]"
        End
    End If
    
    
    If rsCompany.State = 1 Then rsCompany.Close
    'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsCompany.Open "SELECT * FROM PEN0000 WHERE CDEFAULT = 'T' ", dbCompany, adOpenDynamic, adLockBatchOptimistic, 1
    rsCompany.Open "SELECT CHEAD FROM PEN0000 WITH (NOLOCK) WHERE CDEFAULT = 'T' ", dbCompany, adOpenForwardOnly, adLockReadOnly, 1
    If Not rsCompany.EOF Then
        'mPathComp = rsCompany!CPATH
        mPathComp = vgsDatabase
        CHEAD = SkipNull(rsCompany!CHEAD)
        FMain.Caption = FMain.Caption & " [" & CHEAD & "]"
        FMain.TrayArea1.ToolTip = FMain.Caption
        App.Title = FMain.Caption
    End If
    
    'commented by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'ConnectDataBaseAccounts (mPathComp)
    'commented by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    If rsSelect.State = 1 Then rsSelect.Close
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    'rsSelect.Open "Select SABRESTATUS,STATUS,AUTOSTART,UPLOADDIRNAME,DESTDIRNAME,SABREUPLOADDIRNAME,SABREDESTDIRNAME,AMDUPLOADDIRNAME,AMDDESTDIRNAME,AMDSTATUS,WSPUPLOADDIRNAME,WSPDESTDIRNAME,WSPSTATUS From [File]", dbCompany, adOpenForwardOnly, adLockReadOnly
    'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsSelect.Open "Select SABRESTATUS,STATUS,AUTOSTART,UPLOADDIRNAME,DESTDIRNAME,SABREUPLOADDIRNAME,SABREDESTDIRNAME,AMDUPLOADDIRNAME,AMDDESTDIRNAME,AMDSTATUS,WSPUPLOADDIRNAME,WSPDESTDIRNAME,WSPSTATUS,LegacyEnable,LegacySourcePath,LegacyDestinationPath From [File]", dbCompany, adOpenForwardOnly, adLockReadOnly
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    'rsSelect.Open "Select SABRESTATUS,STATUS,AUTOSTART,UPLOADDIRNAME,DESTDIRNAME,SABREUPLOADDIRNAME,SABREDESTDIRNAME,AMDUPLOADDIRNAME,AMDDESTDIRNAME,AMDSTATUS,WSPUPLOADDIRNAME,WSPDESTDIRNAME,WSPSTATUS,LegacyEnable,LegacySourcePath,LegacyDestinationPath From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
    rsSelect.Open "Select SABRESTATUS,STATUS,AUTOSTART,UPLOADDIRNAME,DESTDIRNAME,SABREUPLOADDIRNAME,SABREDESTDIRNAME,AMDUPLOADDIRNAME,AMDDESTDIRNAME,AMDSTATUS,WSPUPLOADDIRNAME,WSPDESTDIRNAME,WSPSTATUS,LegacyEnable,LegacySourcePath,LegacyDestinationPath,OutOfBookingsDateAmadeus From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    If rsSelect.EOF = False Then
        FMain.txtGalilieoSource = SkipNull(rsSelect!UPLOADDIRNAME)
        FMain.txtGalilieoDest = SkipNull(rsSelect!DESTDIRNAME)
        FMain.chkGalilieo = Val(SkipNull(rsSelect!Status))
        FMain.txtGalileoExt = INIRead(INIPenAIR_String, "PenGDS", "GalileoExt", "*.mir")
        
        FMain.txtSource = SkipNull(rsSelect!SABREUPLOADDIRNAME)
        FMain.txtDest = SkipNull(rsSelect!SABREDESTDIRNAME)
        FMain.chkSabre = Val(SkipNull(rsSelect!SABRESTATUS))
        FMain.chkSabreIncludeItineraryOnly = Val(INIRead(INIPenAIR_String, "PenGDS", "SabreIncludeItineraryOnly", "0"))
        FMain.txtSabreExt = INIRead(INIPenAIR_String, "PenGDS", "SabreExt", "*.fil;*.pnr")
        
        FMain.txtWorldspanSource = SkipNull(rsSelect!WSPUPLOADDIRNAME)
        FMain.txtWorldspanDest = SkipNull(rsSelect!WSPDESTDIRNAME)
        FMain.chkWorldspan = Val(SkipNull(rsSelect!WSPSTATUS))
        FMain.txtWorldspanExt = INIRead(INIPenAIR_String, "PenGDS", "WorldspanExt", "*.prt")
        
        'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
        FMain.OutOfBookingsDateAmadeusText.Text = DateFormat1900toBlank(SkipNull(rsSelect.Fields("OutOfBookingsDateAmadeus")))
        'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
        FMain.txtAmadeusSource = SkipNull(rsSelect!AMDUPLOADDIRNAME)
        FMain.txtAmadeusDest = SkipNull(rsSelect!AMDDESTDIRNAME)
        FMain.chkAmadeus = Val(SkipNull(rsSelect!AMDSTATUS))
        FMain.txtAmadeusExt = INIRead(INIPenAIR_String, "PenGDS", "AmadeusExt", "*.air;*.txt")
        
        'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
        FMain.LegacySourcePathText = SkipNull(rsSelect!LegacySourcePath)
        FMain.LegacyDestinationPathText = SkipNull(rsSelect!LegacyDestinationPath)
        FMain.LegacyEnableCheck = Val(SkipNull(rsSelect!LegacyEnable))
        FMain.LegacyExtText = INIRead(INIPenAIR_String, "PenGDS", "LegacyExt", "*.xls")
        
        FMain.chkAutoStart = Val(SkipNull(rsSelect!AutoStart))
    End If
    
    'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    'by Abhi on 25-Mar-2013 for caseid 2786 PenGDS Sabre PNR missing - SabreM1 row not found. we need to skip this file
    If Trim(FMain.txtSource) <> "" Then
        If PathExists(FMain.txtSource) = True Then
            'If Dir(FMain.txtSource & "\Error M1 Missing\", vbDirectory) = "" Then
            If PathExists(FMain.txtSource & "\Error M1 Missing\") = False Then
                MkDir FMain.txtSource & "\Error M1 Missing\"
            End If
        End If
    End If
    'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    If Trim(FMain.OutOfBookingsDateAmadeusText.Text) <> "" Then
        If Trim(FMain.txtAmadeusSource) <> "" Then
            If PathExists(FMain.txtAmadeusSource) = True Then
                If PathExists(FMain.txtAmadeusSource & "\Outofbookings\") = False Then
                    MkDir FMain.txtAmadeusSource & "\Outofbookings\"
                End If
            End If
        End If
    End If
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    Call getPENLINEID
    
    'by Abhi on 01-Jul-2014 for caseid 4222 Exchange value picking logic-Hardcoded currency "GBP" should be replaced with company currency for Galileo
    COMCID_String = getFromFileTable("COMCID")
    'by Abhi on 01-Jul-2014 for caseid 4222 Exchange value picking logic-Hardcoded currency "GBP" should be replaced with company currency for Galileo
    
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    AnyDataChanged_Boolean = False
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    
    If Command$ = "vbMinimized" And (FMain.chkAutoStart <> vbUnchecked And FMain.chkEnableAll <> vbUnchecked) Then
        FMain.Hide
    Else
        FMain.Show
    End If
    DoEvents
    If FMain.chkAutoStart = vbChecked Then
        FMain.Hide
        FMain.cmdStart_Click
    End If
End Sub
'Public Function mSeqNumberGen(Field As String)
''by Abhi on 15-Jul-2010 for caseid 1289 Number genarating issue
''On Error GoTo myerror
'On Error GoTo PENErr
'Dim PENErr_Number As String, PENErr_Description As String
'    Dim vlno As Long
'    'by Abhi on 05-Apr-2010 for caseid 1289 Number genarating issue
'    Dim vNextNumber_Long As Long
'    Dim rsUpdateCode As New ADODB.Recordset
'    rsUpdateCode.CursorLocation = adUseServer
'abc:
'        'Query = "Select " & Field & " as NextNumber from PEN0001 where yearid =(select Max(cast(yearid as bigint)) from PEN0001) "
'        Query = "Select " & Field & " as NextNumber from PEN0001 where YRCLOSEDYN =0  "
'        If rsUpdateCode.State = 1 Then rsUpdateCode.Close
'        'by Abhi on 12-Nov-2010 for caseid 1546 PenGDS Optimistic concurrency check failed
'        'rsUpdateCode.Open Query, dbCompany, adOpenDynamic, adLockOptimistic
'        rsUpdateCode.Open Query, dbCompany, adOpenDynamic, adLockPessimistic
'        If rsUpdateCode.EOF = False Then
'            vlno = Val(IIf(IsNull(rsUpdateCode!NextNumber), 1, rsUpdateCode!NextNumber))
'            'rsUpdateCode.Fields("NextNumber") = vlno + 1
'            'by Abhi on 05-Apr-2010 for caseid 1289 Number genarating issue
'            vNextNumber_Long = vlno + 1
'            rsUpdateCode.Fields("NextNumber") = vNextNumber_Long
'            rsUpdateCode.Update
'            DoEvents
'            'SendERROR "ERROR in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "The microsoft database engine stopped the process because you and another user are attempting to change the same data at the same time"
'            mSeqNumberGen = vlno
'        End If
'Exit Function
''by Abhi on 15-Jul-2010 for caseid 1289 Number genarating issue
''myerror:
'PENErr:
'    PENErr_Number = Err.Number
'    PENErr_Description = Err.Description
'    'by Abhi on 15-Jul-2010 for caseid 1289 Number genarating issue
'    'If Err <> 0 Then
'    If PENErr_Number <> 0 Then
'        rsUpdateCode.CancelUpdate
'        If PENErr_Number <> -2147467259 Then 'Deadlock
'            FMain.WindowState = vbNormal
'            FMain.Show
'            FMain.SendStatus "Error"
'            'SendERROR "ERROR in PenGDS[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & PENErr_Number & " - " & PENErr_Description & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & ")"
'            SendERROR "ERROR in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & PENErr_Number & " - " & PENErr_Description & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & ")"
'            ret = MsgBox("Error: " & PENErr_Number & vbCrLf & PENErr_Description, vbOKCancel + vbCritical, App.Title & " [mSeqNumberGen]")
'            If ret = vbCancel Then
'                FMain.cmdStop_Click
'                Exit Function
'            End If
'        End If
'        'rsUpdateCode.CancelUpdate
'        GoTo abc
'    End If
'End Function

'by Abhi on 04-Jan-2010 for caseid 1582 first free numbers sequence number genaration
Public Function PENFirstFreeNumber(ByVal pType_String As String) As Long
Dim vFirstFreeNumber_Long As Long, vFFN_NEXTNO_Long As Long
Dim vSQL_String As String
Dim vRecordset As New ADODB.Recordset

    vSQL_String = "" _
        & "SELECT     FFN_NEXTNO " _
        & "From PENFFN1 " _
        & "WHERE     (FFN_TYPE = N'" & SkipChars(pType_String) & "')"
    If vRecordset.State = 1 Then vRecordset.Close
    vRecordset.Open vSQL_String, dbCompany, adOpenForwardOnly, adLockReadOnly
    If vRecordset.EOF = False Then
        vFirstFreeNumber_Long = Val(SkipNull(vRecordset.Fields("FFN_NEXTNO"), 1))
    End If
    vFFN_NEXTNO_Long = vFirstFreeNumber_Long + 1
    vSQL_String = "" _
        & "UPDATE    PENFFN1 " _
        & "Set FFN_NEXTNO = " & vFFN_NEXTNO_Long & " " _
        & "WHERE     (FFN_TYPE = N'" & SkipChars(pType_String) & "')"
    dbCompany.Execute vSQL_String

PENFirstFreeNumber = vFirstFreeNumber_Long
Exit Function
End Function

''''by Abhi on 04-Jan-2010 for caseid 1582 first free numbers sequence number genaration
'''Public Function PENFirstFreeNumber(ByVal pTYPE_String As String) As Long
'''On Error GoTo PENErr
'''Dim PENErr_Number As String, PENErr_Description As String
'''Dim vFirstFreeNumber_Long As Long, vFFN_NEXTNO_Long As Long
'''Dim vSQL_String As String
'''Dim vRecordset As New ADODB.Recordset
'''
''''by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
'''DeadlockRETRY:
'''    vSQL_String = "" _
'''        & "SELECT     FFN_NEXTNO1 " _
'''        & "From PENFFN1 " _
'''        & "WHERE     (FFN_TYPE = N'" & SkipChars(pTYPE_String) & "')"
'''    If vRecordset.State = 1 Then vRecordset.Close
'''    vRecordset.Open vSQL_String, dbCompany, adOpenDynamic, adLockPessimistic
'''    If vRecordset.EOF = False Then
'''        vFirstFreeNumber_Long = Val(SkipNull(vRecordset.Fields("FFN_NEXTNO1"), 1))
'''        vFFN_NEXTNO_Long = vFirstFreeNumber_Long + 1
'''        vRecordset.Fields("FFN_NEXTNO") = vFFN_NEXTNO_Long
'''        vRecordset.Update
'''        DoEvents
'''        'SendERROR "ERROR in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "The microsoft database engine stopped the process because you and another user are attempting to change the same data at the same time"
'''        PENFirstFreeNumber = vFirstFreeNumber_Long
'''    End If
'''Exit Function
'''
'''PENErr:
'''    PENErr_Number = Err.Number
'''    PENErr_Description = Err.Description
'''    'by Abhi on 24-Jul-2009 for Deadlock
'''    If PENErr_Number = -2147467259 Then 'Deadlock
'''        Debug.Print "Deadlock"
'''        'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
'''        Sleep 5
'''        GoTo DeadlockRETRY
'''    End If
'''    'MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " (PENFirstFreeNumber)"
'''    If PENErr_Number <> 0 Then
'''        vRecordset.CancelUpdate
'''        FMain.WindowState = vbNormal
'''        FMain.Show
'''        FMain.SendStatus "Error"
'''        SendERROR "ERROR in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & PENErr_Number & " - " & PENErr_Description & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & ")"
'''        ret = MsgBox("Error: " & PENErr_Number & vbCrLf & PENErr_Description, vbOKCancel + vbCritical, App.Title & " (PENFirstFreeNumber)")
'''        If ret = vbCancel Then
'''            FMain.cmdStop_Click
'''            Exit Function
'''        End If
'''    End If
'''End Function

Public Sub ConnectDataBaseAccounts(mPath)
Set dbCompany = New ADODB.Connection
dbCompany.CursorLocation = adUseClient
'by Abhi on 07-Dec-2009 for Timeout Expired on Ariana
dbCompany.CommandTimeout = 300 '30
mPath = mPath
'DBEngine.RegisterDatabase "PenSoft", "SQL Server", True, "Description=PenSoft" & vbCr & "SERVER=" & vgsServer & vbCr & "DATABASE=" & mPath
'DBEngine.RegisterDatabase "PenSoft", vgsBackend, True, "Description=PenSoft" & vbCr & "SERVER=PenAIR" & vbCr & "DATABASE=" & mPath & vbCr & "Network=DBMSSOCN" & vbCr & "Address=" & vgsServer
If Trim(vgsProvider) = "" Then
    DBEngine.RegisterDatabase "PenSoft", vgsBackend, True, "Description=PenSoft" & vbCr & "SERVER=" & vgsServer & vbCr & "DATABASE=" & mPath
    dbCompany.Open "PenSoft", vgsUID, vgsPWD
Else
    dbCompany.Open GetConnectionString(vgsBackend, vgsProvider, vgsServer, vgsUID, vgsPWD, mPath, vgbStandardSecurity)
End If
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
        DoEvents
    Next N
    N = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)
End Sub

Public Function SkipNull(value, Optional default = Empty)
    Dim A
    A = IIf(IsNull(value) = True, default, value)
    SkipNull = A
End Function

'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
Public Function RemoveQuotes(pValue) As String
Dim vData_String As String
Dim vLen_Long As Long
    
    vData_String = SkipNull(pValue)
    vData_String = Trim(vData_String)
    vLen_Long = Len(vData_String)
    If Left(vData_String, 1) = "'" Or Left(vData_String, 1) = """" Then
        vData_String = Right(vData_String, vLen_Long - 1)
    End If
    vLen_Long = Len(vData_String)
    If Right(vData_String, 1) = "'" Or Right(vData_String, 1) = """" Then
        vData_String = Left(vData_String, vLen_Long - 1)
    End If
RemoveQuotes = vData_String
End Function

Public Sub DeleteFileToRecycleBin(FileName As String)

'Kill fileName
  Dim fop As SHFILEOPTSTRUCT

  With fop
    .wFunc = FO_DELETE
    .pFrom = FileName
    .fFlags = FOF_ALLOWUNDO
  End With

  SHFileOperation fop

End Sub


Public Function GetFileNameFromPath(Path As String) As String
Dim temp, Count As Long
temp = Split(Path, "\")
Count = UBound(temp)
If Count >= 0 Then
    GetFileNameFromPath = temp(Count)
End If
End Function
Public Function SkipNegative(value, Optional default = 0)
    SkipNegative = IIf(value < 0, default, value)
End Function

Public Function PathExists(ByVal vPath As String) As Boolean
If Val(INIRead(INIPenAIR_String, "PenGDS", "SpecialNetwork", "0")) = 1 Then
    PathExists = PathExistsSpecial(vPath)
    Exit Function
End If

On Error GoTo myErr1
    DoEvents
    If Dir(vPath, vbDirectory) = "" Then
        PathExists = False
    Else
        PathExists = True
    End If

Exit Function
myErr1:
'by Abhi on 15-Jan-2015 for caseid 4893 PenGDS shared path is getting Error: 76 - Path not found
'On Error GoTo myErr2
'    DoEvents
'    FMain.File1.Path = vPath
'    PathExists = True
'Exit Function
'myErr2:
'    DoEvents
'    PathExists = False
    DoEvents
    PathExists = PathExistsSpecial(vPath)
Exit Function
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

Public Function PathExistsSpecial(ByVal vPath As String) As Boolean
On Error GoTo myErr1
    DoEvents
    FMain.File1.Path = vPath
    PathExistsSpecial = True
Exit Function
myErr1:
    DoEvents
    PathExistsSpecial = False
Exit Function
End Function


Public Function IfExistsinTargetRename(vPath As String, ByVal vFname As String, vTargetPath As String) As String
'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
'On Error GoTo MyErr
On Error GoTo PENErr
Dim ErrNumber As String
Dim ErrDescription As String
'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available

Dim FileCount As Long
Dim vFname_No As String
'by Abhi on 12-Mar-2012 for caseid 1652 PenGDS Permission denied added file closed checking
Dim vFileTitle_String As String
Dim vFileExtension_String As String

    'by Abhi on 03-Mar-2016 for caseid 6102 Error 0 Warning(IfExistsinTargetRename) in PenGDS - Worldspan
    FMain.stbUpload.Panels(1).Text = "Checking FileExists..."
    'Dim FileObj As New FileSystemObject
    'FileObj.DeleteFile vPath & "\" & vFname, True
    If Dir(vPath & "\" & vFname) = "" Then
        IfExistsinTargetRename = ""
        Exit Function
    End If
    'by Abhi on 03-Mar-2016 for caseid 6102 Error 0 Warning(IfExistsinTargetRename) in PenGDS - Worldspan
    'by Abhi on 12-Mar-2012 for caseid 1652 PenGDS Permission denied added file closed checking
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    'FMain.stbUpload.Panels(1).Text = "Waiting for file available..."
    FMain.stbUpload.Panels(1).Text = "File is in use, waiting..."
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    'Call Wait4FileAvailable(vPath & "\" & vFname)
    'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
    'If Wait4FileAvailable(vPath & "\" & vFname) = True Then 'True means timed out, False means file FileStatus is Closed and can read the file
    If Wait4FileAvailable(vPath & "\" & vFname) = False Then 'False means timed out/error, True means file FileStatus is Closed and can read the file
    'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
        GoTo PENErr
    End If
    'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    'by Abhi on 07-Oct-2010 for caseid 1516 PenGDS Amadeus slow reading
    FMain.stbUpload.Panels(1).Text = "Checking duplicate filename..."
    FileCount = 1
    vFname_No = vFname
    'by Abhi on 12-Mar-2012 for caseid 1652 PenGDS Permission denied added file closed checking
    Call SplitFileTitleandExtension(vFname, vFileTitle_String, vFileExtension_String)
    
    'by Abhi on 07-Jul-2014 for caseid 4258 Change in logic for file name checking when moving to target or error files
    'Do While Dir(vTargetPath & "\" & vFname_No) = vFname_No
    '    FileCount = FileCount + 1
    '    'by Abhi on 12-Mar-2012 for caseid 1652 PenGDS Permission denied added file closed checking
    '    'vFname_No = Left(vFname, Len(vFname) - 4) & "_" & Format(FileCount, "00") & Right(vFname, 4)
    '    vFname_No = vFileTitle_String & "_" & Format(FileCount, "000") & vFileExtension_String
    '    FMain.stbUpload.Panels(3).Text = vFname_No
    '    If Dir(vTargetPath & "\" & vFname_No) = "" Then
    '        Name vPath & "\" & vFname As vPath & "\" & vFname_No
    '    End If
    '    DoEvents
    'Loop
    'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
    If Dir(vPath & "\" & vFname) = vFname Then
    'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
        If Dir(vTargetPath & "\" & vFname_No) = vFname_No Then
            vFname_No = vFileTitle_String & "_" & Format(Now, "YYYYMMDDHHMMSS") & vFileExtension_String
            'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
            'FMain.stbUpload.Panels(3).Text = vFname_No
            'Name vPath & "\" & vFname As vPath & "\" & vFname_No
            Name vPath & "\" & vFname As vPath & "\" & vFname_No
            FMain.stbUpload.Panels(3).Text = vFname_No
            'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
            DoEvents
        End If
    'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
    Else
        vFname_No = ""
    End If
    'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
    'by Abhi on 07-Jul-2014 for caseid 4258 Change in logic for file name checking when moving to target or error files
IfExistsinTargetRename = vFname_No
Exit Function
'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
'MyErr:
'
'    Sleep 500
'    Resume
PENErr:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    
    FMain.cmdStop_Click
    NoofPermissionDenied = 0
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    If ErrDetails_String = "FMain.fForcedStop_Boolean = True" Then
        ErrDetails_String = ""
        Exit Function
    End If
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    'SendERROR "Warning(IfExistsinTargetRename) in PenGDS[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & "). PenGDS is automatically Resumed."
    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
    'SendERROR "Warning(IfExistsinTargetRename) in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed."
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'SendERROR "Warning(IfExistsinTargetRename) in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed.", ErrNumber
    SendERROR "Warning(IfExistsinTargetRename) in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ErrDetails_String & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed.", ErrNumber
    ErrDetails_String = ""
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
    FMain.cmdStart_Click
'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
End Function

Public Function StringOccurs(ByVal MainString As String, ByVal SubString As String, Optional ByVal flag As Boolean = True) As Long
    Dim p As Integer
    Dim cnt As Integer, pos As Integer
    
    MainString = LCase(MainString)
    SubString = LCase(SubString)
    
    p = 0
    Do While p < Len(MainString)
        If flag = False Then
            p = InStr(p + 1, MainString, SubString, vbTextCompare)
            If p > 0 Then
                cnt = cnt + 1
            Else
                Exit Do
            End If
        Else
            If pos = 0 Then
                p = InStr(p + 1, MainString, SubString & " ", vbTextCompare)
            Else
                p = InStr(p + 1, MainString, " " & SubString & " ", vbTextCompare)
            End If
            
            pos = pos + 1
            
            If p > 0 Then
                cnt = cnt + 1
            Else
                Exit Do
            End If
        End If
        DoEvents
    Loop
    
    If Right(MainString, Len(SubString) + 1) = (" " & SubString) Then
        cnt = cnt + 1
    End If
    StringOccurs = cnt
End Function

Public Function ShowProgress(Optional ByVal vMessage As String)
On Error Resume Next
    If Trim(vMessage) = "" Then
        'vMessage = "PenAIR progressing... please wait"
        vMessage = App.ProductName & " progressing... please wait"
    End If
    FrmProgress.lblMessage = vMessage
    FrmProgress.Show
    DoEvents
End Function

Public Function HideProgress()
    Unload FrmProgress
    DoEvents
End Function


Public Sub ShellAndWait(ByVal program_name As String, ByVal window_style As VbAppWinStyle)
Dim process_id As Long
Dim process_handle As Long

    ' Start the program.
    On Error GoTo ShellError
    process_id = Shell(program_name, window_style)
    On Error GoTo 0

    ' Hide.
    DoEvents

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    ' Reappear.
    Exit Sub

ShellError:
    MsgBox "Error starting task " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
         App.Title
End Sub

'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
'Public Function SendERROREmail(ByVal pSub As String, ByVal pMess As String) As Boolean
Public Function SendERROREmail(ByVal pSub As String, ByVal pMess As String, ByVal pErrNumber_String As String) As Boolean
'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
Dim vSendEMAIL As Boolean
Dim vSendEMAILTOEmailAddresses
Dim vSuccess As Boolean
Dim vSendEMAIL_Type As SendEMAIL_Type
'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
Dim vEnd, vDuration
'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email

    
    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
    'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
    If Trim(pErrNumber_String) <> "" Then
    'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
        If SendERROREmailStart = "" Then
            SendERROREmailStart = Now
        End If
        vEnd = Now
        vDuration = DateDiff("n", SendERROREmailStart, vEnd)
        If pErrNumber_String = SendERROREmailPreviousErrNumber_String Then
            If Val(vDuration) <= 5 Then '5 mins
                If Trim(SendERROREmailPreviousMess_String) <> "" Then
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                    'SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & vbCrLf
                    SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & vbCrLf & vbCrLf
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                End If
                'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
                'SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & pMess
                SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & DateFormat(Date) & " " & Format(time, "HH:MM:SS AMPM") & "(" & TimeFormat(time) & ")" & " : " & pMess
                'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
                FMain.SSTab1.TabPicture(0) = FMain.ImageError.Picture
                Exit Function
            Else
                SendERROREmailStart = Now
                If Trim(SendERROREmailPreviousMess_String) <> "" Then
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                    'SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & vbCrLf
                    SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & vbCrLf & vbCrLf
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                End If
                'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
                'SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & pMess
                SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & DateFormat(Date) & " " & Format(time, "HH:MM:SS AMPM") & "(" & TimeFormat(time) & ")" & " : " & pMess
                'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
            End If
        Else
            SendERROREmailStart = Now
            'If pErrNumber_String <> "SendERROREmailBeforeExit" Then
                If Trim(SendERROREmailPreviousMess_String) <> "" Then
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                    'SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & vbCrLf
                    SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & vbCrLf & vbCrLf
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                End If
                'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
                'SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & pMess
                SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & DateFormat(Date) & " " & Format(time, "HH:MM:SS AMPM") & "(" & TimeFormat(time) & ")" & " : " & pMess
                'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
            'End If
        End If
        SendERROREmailPreviousErrNumber_String = pErrNumber_String
        FMain.SSTab1.TabPicture(0) = LoadPicture()
    'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
    Else
        'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
        'SendERROREmailPreviousMess_String = pMess
        SendERROREmailPreviousMess_String = DateFormat(Date) & " " & Format(time, "HH:MM:SS AMPM") & "(" & TimeFormat(time) & ")" & " : " & pMess
        'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
    End If
    'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
    
    vSendEMAIL = Val(INIRead(INIPenAIR_String, "PenGDS", "SendEMAIL", "0"))
    vSendEMAILTOEmailAddresses = INIRead(INIPenAIR_String, "PenGDS", "SendEMAILTOEmailAddresses", "")
    'by Abhi on 25-Mar-2013 for caseid 2786 PenGDS Sabre PNR missing - SabreM1 row not found. we need to skip this file
    If InStr(1, pMess, "-2147220991", vbTextCompare) Then
        vSendEMAILTOEmailAddresses = "abhi@penguininc.com"
    End If
    
    vSendEMAIL_Type.TOEMAILs_String = vSendEMAILTOEmailAddresses
    vSendEMAIL_Type.CCEMAILs_String = ""
    vSendEMAIL_Type.BCCEMAILs_String = ""
    vSendEMAIL_Type.SUBJECT_String = pSub
    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
    'vSendEMAIL_Type.BODY_String = pMess
    'vSendEMAIL_Type.BODYFILENAME_String = ""
    vSendEMAIL_Type.BODY_String = ""
    vSendEMAIL_Type.BODYFILENAME_String = GetEMAILBodyFileName(SendERROREmailPreviousMess_String)
    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
    vSendEMAIL_Type.FORMAT_SendEMAILFORMAT_EnumOptional = SendEMAILFORMATAUTO
    'vSendEMAIL_Type.FROMEMAIL_String = "PenGDS"
    vSendEMAIL_Type.FROMEMAIL_String = App.ProductName
    
    If vSendEMAIL = True And vSendEMAILTOEmailAddresses <> "" Then
        'vSuccess = CallDotNetSendEMAIL(vSendEMAILTOEmailAddresses, pSub, pMess, "", "PenGDS <support@penguininc.com>")
        vSuccess = CallDotNetSendEMAIL(vSendEMAIL_Type)
    End If

    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
    Kill vSendEMAIL_Type.BODYFILENAME_String
    SendERROREmailPreviousMess_String = ""
    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email

SendERROREmail = vSuccess
End Function
Public Function SendERRORSms(ByVal pSub As String, Optional ByVal pMess As String)
    'CallDotNetSendEMAIL "abhi@penguininc.com", vSub, vMess, "", "support@penguininc.com"
End Function

'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
'Public Function SendERROR(ByVal pSub As String, Optional ByVal pMess As String) As Boolean
Public Function SendERROR(ByVal pSub As String, ByVal pMess As String, ByVal pErrNumber_String As String) As Boolean
'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
Dim vSuccess As Boolean
    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
    'vSuccess = SendERROREmail(pSub, pMess)
    vSuccess = SendERROREmail(pSub, pMess, pErrNumber_String)
    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
    'SendERRORSms

SendERROR = vSuccess
End Function

Public Function CallPenSEARCH(ByVal pGDS As String)
Dim vShellCommandPara As String

    'vShellCommandPara = """" & vShellCommandPara & """"
    vShellCommandPara = pGDS
    Shell """" & App.Path & "\PenSEARCH.exe" & """" & " " & vShellCommandPara, vbNormalFocus

End Function

Public Function GetConnectionString(ByVal pBackend_String As String, ByVal pProvider_String As String, ByVal pServer_String As String, ByVal pUID_String As String, ByVal pPWD_String As String, ByVal pDatabase_String As String, ByVal pStandardSecurity_Boolean As Boolean) As String
Dim vConnectionString_String As String

'    Select Case Trim(pBackend_String)
'        Case "SQL Server"
'            Select Case Trim(pProvider_String)
'                Case "sqloledb", "SQLNCLI", "SQLNCLI10"
'                    pProvider_String = "sqloledb"
'                Case "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}"
'                    pProvider_String = "{SQL Server}"
'            End Select
'        Case "SQL Server 2005"
'            Select Case Trim(pProvider_String)
'                Case "sqloledb", "SQLNCLI", "SQLNCLI10"
'                    pProvider_String = "SQLNCLI"
'                Case "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}"
'                    pProvider_String = "{SQL Native Client}"
'            End Select
'        Case "SQL Server 2008"
'            Select Case Trim(pProvider_String)
'                Case "sqloledb", "SQLNCLI", "SQLNCLI10"
'                    pProvider_String = "SQLNCLI10"
'                Case "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}"
'                    pProvider_String = "{SQL Server Native Client 10.0}"
'            End Select
'    End Select
    
    vConnectionString_String = ""
    'Provider/Driver
    Select Case Trim(pProvider_String)
        'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
        'Case "sqloledb", "SQLNCLI", "SQLNCLI10"
        Case "sqloledb", "SQLNCLI", "SQLNCLI10", "SQLNCLI11"
        'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
            vConnectionString_String = vConnectionString_String & "Provider=" & pProvider_String & ";"
        'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
        'Case "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}"
        Case "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}", "{SQL Server Native Client 11.0}"
        'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
            vConnectionString_String = vConnectionString_String & "Driver=" & pProvider_String & ";"
    End Select
    'Data Source/Server
    Select Case Trim(pProvider_String)
        Case "sqloledb"
            vConnectionString_String = vConnectionString_String & "Data Source=" & pServer_String & ";"
        'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
        'Case "SQLNCLI", "SQLNCLI10", "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}"
        Case "SQLNCLI", "SQLNCLI10", "SQLNCLI11", "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}", "{SQL Server Native Client 11.0}"
        'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
            vConnectionString_String = vConnectionString_String & "Server=" & pServer_String & ";"
    End Select
    'Initial Catalog/Database
    Select Case Trim(pProvider_String)
        Case "sqloledb"
            vConnectionString_String = vConnectionString_String & "Initial Catalog=" & pDatabase_String & ";"
        'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
        'Case "SQLNCLI", "SQLNCLI10", "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}"
        Case "SQLNCLI", "SQLNCLI10", "SQLNCLI11", "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}", "{SQL Server Native Client 11.0}"
        'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
            vConnectionString_String = vConnectionString_String & "Database=" & pDatabase_String & ";"
    End Select
    If pStandardSecurity_Boolean = True Then
        'User ID/Uid
        Select Case Trim(pProvider_String)
            Case "sqloledb"
                vConnectionString_String = vConnectionString_String & "User ID=" & pUID_String & ";"
            'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
            'Case "SQLNCLI", "SQLNCLI10", "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}"
            Case "SQLNCLI", "SQLNCLI10", "SQLNCLI11", "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}", "{SQL Server Native Client 11.0}"
            'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
                vConnectionString_String = vConnectionString_String & "Uid=" & pUID_String & ";"
        End Select
        'Password/Pwd
        Select Case Trim(pProvider_String)
            Case "sqloledb"
                vConnectionString_String = vConnectionString_String & "Password=" & pPWD_String & ";"
            'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
            'Case "SQLNCLI", "SQLNCLI10", "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}"
            Case "SQLNCLI", "SQLNCLI10", "SQLNCLI11", "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}", "{SQL Server Native Client 11.0}"
            'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
                vConnectionString_String = vConnectionString_String & "Pwd=" & pPWD_String & ";"
        End Select
    Else
        'Integrated Security=SSPI/Trusted_Connection=Yes/Trusted_Connection=yes
        Select Case Trim(pProvider_String)
            Case "sqloledb"
                vConnectionString_String = vConnectionString_String & "Integrated Security=SSPI;"
            'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
            'Case "{SQL Server}"
            '    vConnectionString_String = vConnectionString_String & "Trusted_Connection=Yes;"
            'Case "SQLNCLI", "SQLNCLI10", "{SQL Native Client}", "{SQL Server Native Client 10.0}"
            Case "SQLNCLI", "SQLNCLI10", "SQLNCLI11", "{SQL Server}", "{SQL Native Client}", "{SQL Server Native Client 10.0}", "{SQL Server Native Client 11.0}"
            'by Abhi on 07-Jan-2015 for caseid 4876 SQL Server 2012
                vConnectionString_String = vConnectionString_String & "Trusted_Connection=yes;"
        End Select
    End If
    
'    Select Case Trim(pProvider_String)
'    'SQL Server
'        Case "sqloledb" 'Microsoft OLE DB Provider for SQL Server
'            If pStandardSecurity_Boolean = True Then
'                vConnectionString_String = _
'                        "Provider=" & pProvider_String & ";" & _
'                        "Data Source=" & pServer_String & ";" & _
'                        "Initial Catalog=" & pDatabase_String & ";" & _
'                        "User ID=" & pUID_String & ";" & _
'                        "Password=" & pPWD_String & ";"
'            Else
'                vConnectionString_String = _
'                        "Provider=" & pProvider_String & ";" & _
'                        "Data Source=" & pServer_String & ";" & _
'                        "Initial Catalog=" & pDatabase_String & ";" & _
'                        "Integrated Security=SSPI;"
'            End If
'        Case "{SQL Server}" 'Microsoft SQL Server ODBC Driver
'            If pStandardSecurity_Boolean = True Then
'                vConnectionString_String = _
'                        "Driver=" & pProvider_String & ";" & _
'                        "Server=" & pServer_String & ";" & _
'                        "Database=" & pDatabase_String & ";" & _
'                        "Uid=" & pUID_String & ";" & _
'                        "Pwd=" & pPWD_String & ";"
'            Else
'                vConnectionString_String = _
'                        "Driver=" & pProvider_String & ";" & _
'                        "Server=" & pServer_String & ";" & _
'                        "Database=" & pDatabase_String & ";" & _
'                        "Trusted_Connection=Yes;"
'            End If
'    'SQL Server 2005
'        Case "SQLNCLI" 'SQL Native Client 9.0 OLE DB provider
'            If pStandardSecurity_Boolean = True Then
'                vConnectionString_String = _
'                        "Provider=" & pProvider_String & ";" & _
'                        "Server=" & pServer_String & ";" & _
'                        "Database=" & pDatabase_String & ";" & _
'                        "Uid=" & pUID_String & ";" & _
'                        "Pwd=" & pPWD_String & ";"
'            Else
'                vConnectionString_String = _
'                        "Provider=" & pProvider_String & ";" & _
'                        "Server=" & pServer_String & ";" & _
'                        "Database=" & pDatabase_String & ";" & _
'                        "Trusted_Connection=yes;"
'            End If
'        Case "{SQL Native Client}" 'SQL Native Client 9.0 ODBC Driver
'            If pStandardSecurity_Boolean = True Then
'                vConnectionString_String = _
'                        "Driver=" & pProvider_String & ";" & _
'                        "Server=" & pServer_String & ";" & _
'                        "Database=" & pDatabase_String & ";" & _
'                        "Uid=" & pUID_String & ";" & _
'                        "Pwd=" & pPWD_String & ";"
'            Else
'                vConnectionString_String = _
'                        "Driver=" & pProvider_String & ";" & _
'                        "Server=" & pServer_String & ";" & _
'                        "Database=" & pDatabase_String & ";" & _
'                        "Trusted_Connection=yes;"
'            End If
'    'SQL Server 2008
'        Case "SQLNCLI10" 'SQL Server Native Client 10.0 OLE DB Provider
'            If pStandardSecurity_Boolean = True Then
'                vConnectionString_String = _
'                        "Provider=" & pProvider_String & ";" & _
'                        "Server=" & pServer_String & ";" & _
'                        "Database=" & pDatabase_String & ";" & _
'                        "Uid=" & pUID_String & ";" & _
'                        "Pwd=" & pPWD_String & ";"
'            Else
'                vConnectionString_String = _
'                        "Provider=" & pProvider_String & ";" & _
'                        "Server=" & pServer_String & ";" & _
'                        "Database=" & pDatabase_String & ";" & _
'                        "Trusted_Connection=yes;"
'            End If
'        Case "{SQL Server Native Client 10.0}" 'SQL Server Native Client 10.0 ODBC Driver
'            If pStandardSecurity_Boolean = True Then
'                vConnectionString_String = _
'                        "Driver=" & pProvider_String & ";" & _
'                        "Server=" & pServer_String & ";" & _
'                        "Database=" & pDatabase_String & ";" & _
'                        "Uid=" & pUID_String & ";" & _
'                        "Pwd=" & pPWD_String & ";"
'            Else
'                vConnectionString_String = _
'                        "Driver=" & pProvider_String & ";" & _
'                        "Server=" & pServer_String & ";" & _
'                        "Database=" & pDatabase_String & ";" & _
'                        "Trusted_Connection=yes;"
'            End If
'    End Select

GetConnectionString = vConnectionString_String
End Function

Public Function PENEncryptDecrypt(ByVal pEncryptDecrypt_Enum As EncryptDecrypt_Enum, ByVal pPublicKey_String As String, ByVal pData_String As String) As String
Dim vCrypt_String As String
Dim objCrypt      As New vbCrypt.EncryptionTools
    
    Select Case pEncryptDecrypt_Enum
        Case Encrypt
            vCrypt_String = objCrypt.Encrypt(pPublicKey_String, pData_String)
            vCrypt_String = objCrypt.URLEncodeBinaryData(vCrypt_String)
        Case Decrypt
            vCrypt_String = objCrypt.URLDecodeBinaryData(pData_String)
            vCrypt_String = objCrypt.Decrypt(pPublicKey_String, vCrypt_String)
    End Select
PENEncryptDecrypt = vCrypt_String
End Function

Public Function PenguinIntegrity() As Boolean
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
Dim vSQL_String          As String
Dim rsSelect             As New ADODB.Recordset
Dim vDATABASENAME_String As String
Dim vUPWD_String         As String
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
Dim DeadlockRETRY_Integer As Integer
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    'ShowProgress "PenAIR is checking  Integrity... Please wait"
    ShowProgress App.ProductName & " is checking  Integrity... Please wait"
'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
DeadlockRETRY:
    'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'vSQL_String = "" _
        & "SELECT     DATABASENAME " _
        & "From Pen "
    vSQL_String = "" _
        & "SELECT     DATABASENAME " _
        & "From Pen WITH (NOLOCK) "
    rsSelect.Open vSQL_String, dbCompany, adOpenForwardOnly, adLockReadOnly
    If rsSelect.EOF = False Then
        vDATABASENAME_String = SkipNull(rsSelect.Fields("DATABASENAME"))
    Else
        vDATABASENAME_String = 0
    End If
    
    If Trim(vDATABASENAME_String) = Trim(vgsDatabase) Then
        Encrypted_Boolean = True
    End If
    If rsSelect.State = 1 Then rsSelect.Close
    Set rsSelect = Nothing
    HideProgress
    Screen.MousePointer = vbNormal
PenguinIntegrity = True
Exit Function
PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    'by Abhi on 24-Jul-2009 for Deadlock
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    'If PENErr_Number = -2147467259 Then 'Deadlock
    If (PENErr_Number = -2147467259 Or PENErr_Number = -2147217871) And DeadlockRETRY_Integer < 3 Then '-2147467259 Deadlock, -2147217871 Query timeout expired
        DeadlockRETRY_Integer = DeadlockRETRY_Integer + 1
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
        Debug.Print "Deadlock"
        'commented by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
        'Resume
        'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
        Sleep 5
        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
        'GoTo DeadlockRETRY
        Resume DeadlockRETRY
        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    End If
    If rsSelect.State = 1 Then rsSelect.Close
    Set rsSelect = Nothing
    HideProgress
    Screen.MousePointer = vbNormal
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " [Penguin Integrity]"
    PenguinIntegrity = False
    Exit Function
End Function


Public Sub ChangeDateRegionalSettings()
Dim dwLCID As Long
DateSyDefault = GetLocaleString(LOCALE_SSHORTDATE)
dwLCID = GetSystemDefaultLCID()
         If SetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, "dd/MMM/yyyy") _
            = False Then
            MsgBox "Failed", vbCritical, App.Title
            Exit Sub
         End If
         PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
End Sub


Private Function GetLocaleString(ByVal lLocaleNum As Long) As String
'
' Generic routine to get the locale string from the Operating system.
'
    Dim lBuffSize As String
    Dim sBuffer As String
    Dim lRet As Long
'
' Create a string buffer large enough to hold the returned value, 256 should
' be more than enough
'
    lBuffSize = 256
    sBuffer = String$(lBuffSize, vbNullChar)
'
' Get the information from the registry
'
    lRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, lLocaleNum, sBuffer, lBuffSize)
'
' If lRet > 0 then success - lret is the size of the string returned
'
    If lRet > 0 Then
        GetLocaleString = Left$(sBuffer, lRet - 1)
    End If
    
End Function
Public Sub ChangeDateRegionalSettingsDefault()
Dim dwLCID As Long
'DateSyDefault = GetLocaleString(LOCALE_SSHORTDATE)
dwLCID = GetSystemDefaultLCID()
         If SetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, DateSyDefault) _
            = False Then
            'MsgBox "Failed"
            Exit Sub
         End If
         PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
End Sub


Public Sub SelectText(txt As TextBox)
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
End Sub


'by Abhi on 28-Oct-2010 for caseid 1531 PenGDS should take *CHD in passenger name as Child
'Public Function GetPassengerTypefromShort(ByVal pPassType_String As String, Optional ByVal pIfnotFoundSetBlank_Boolean As Boolean = False)
Public Function GetPassengerTypefromShort(ByVal pPassType_String As String, Optional ByVal pIfnotFoundSetBlank_Boolean As Boolean = False, Optional ByVal pPassenger_String As String)
Dim vPassType_String As String
'by Abhi on 12-Jul-2017 for caseid 7618 Pax type B17 picking issue from Amadeus file
Dim vPassTypeFirst_String As String, vPassTypeRest_String As String
'by Abhi on 12-Jul-2017 for caseid 7618 Pax type B17 picking issue from Amadeus file
    
    'by Abhi on 28-Oct-2010 for caseid 1531 PenGDS should take *CHD in passenger name as Child
    If InStr(1, pPassenger_String, "*CHD", vbTextCompare) > 0 Then
        pPassType_String = "CHD"
    End If
    
    Select Case UCase(Trim(pPassType_String))
        'by Abhi on 20-Jul-2010 for caseid 1428 new passenger type
        'Case "ADT"
        'by Abhi on 04-Nov-2010 for caseid 1428 new passenger type VFR as Adult
        'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
        'Case "ADT", "JCB", "ITX", "VFR"
        'by Abhi on 16-May-2017 for caseid 7439 Amadeus fare picking issue due to invalid value in Amdline1 and AmdlineK
        'Case "ADT", "JCB", "ITX", "VFR", ""
        'by Abhi on 24-Jul-2019 for caseid 10548 New Pax type from GDS files TIM TNN TNF
        'Case "ADT", "JCB", "ITX", "VFR", "", "IIT"
        Case "ADT", "JCB", "ITX", "VFR", "", "IIT", "TIM"
        'by Abhi on 24-Jul-2019 for caseid 10548 New Pax type from GDS files TIM TNN TNF
        'by Abhi on 16-May-2017 for caseid 7439 Amadeus fare picking issue due to invalid value in Amdline1 and AmdlineK
        'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
            vPassType_String = "Adult"
        'by Abhi on 20-Jul-2010 for caseid 1428 new passenger type
        'Case "CNN", "CHD"
        'by Abhi on 24-Jul-2019 for caseid 10548 New Pax type from GDS files TIM TNN TNF
        'Case "CNN", "CHD", "JNN", "INN"
        Case "CNN", "CHD", "JNN", "INN", "TNN"
        'by Abhi on 24-Jul-2019 for caseid 10548 New Pax type from GDS files TIM TNN TNF
            vPassType_String = "Child"
        'by Abhi on 27-Feb-2014 for caseid 3780 Passenger type JNF from gal file should be treated as Infant
        'Case "INF"
        'by Abhi on 24-Jul-2019 for caseid 10548 New Pax type from GDS files TIM TNN TNF
        'Case "INF", "JNF"
        Case "INF", "JNF", "TNF"
        'by Abhi on 24-Jul-2019 for caseid 10548 New Pax type from GDS files TIM TNN TNF
        'by Abhi on 27-Feb-2014 for caseid 3780 Passenger type JNF from gal file should be treated as Infant
            vPassType_String = "Infant"
        'by Abhi on 04-Mar-2016 for caseid 6108 New pax type Youth from pnr Files-GDS Galileo
        'by Abhi on 11-Oct-2017 for caseid 7257 Youth tag for worldspan GBE
        'Case "YTH"
        Case "YTH", "GBE"
        'by Abhi on 11-Oct-2017 for caseid 7257 Youth tag for worldspan GBE
            vPassType_String = "Youth"
        'by Abhi on 04-Mar-2016 for caseid 6108 New pax type Youth from pnr Files-GDS Galileo
        Case Else
            'by Abhi on 12-Jul-2017 for caseid 7618 Pax type B17 picking issue from Amadeus file
            ''by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
            ''If pIfnotFoundSetBlank_Boolean = False Then
            ''    vPassType_String = pPassType_String
            ''Else
            ''    vPassType_String = ""
            ''End If
            ''by Abhi on 24-Mar-2014 for caseid 3839 Passenger type -GDS File "CHD" "CNN", "JNN", "INN" (where NN is numberic) will be as "Child"
            ''If Left(UCase(Trim(pPassType_String)), 1) = "J" And IsNumeric(mID(UCase(Trim(pPassType_String)), 2, 2)) = True Then 'JNN means "J09"
            ''by Abhi on 04-Mar-2016 for caseid 6108 New pax type Youth from pnr Files-GDS Galileo
            ''If (Left(UCase(Trim(pPassType_String)), 1) = "J" Or Left(UCase(Trim(pPassType_String)), 1) = "C" Or Left(UCase(Trim(pPassType_String)), 1) = "I") And IsNumeric(Mid(UCase(Trim(pPassType_String)), 2, 2)) = True Then 'JNN means "J09"
            ''by Abhi on 16-May-2017 for caseid 7439 Amadeus fare picking issue due to invalid value in Amdline1 and AmdlineK
            ''If (Left(UCase(Trim(pPassType_String)), 1) = "J" Or Left(UCase(Trim(pPassType_String)), 1) = "C" Or Left(UCase(Trim(pPassType_String)), 1) = "I") And Val(Mid(UCase(Trim(pPassType_String)), 2, 2)) >= 12 And Val(Mid(UCase(Trim(pPassType_String)), 2, 2)) <= 16 Then  'Youth means "C12-C16"
            'If ((Left(UCase(Trim(pPassType_String)), 1) = "J" Or Left(UCase(Trim(pPassType_String)), 1) = "C" Or Left(UCase(Trim(pPassType_String)), 1) = "I" Or Left(UCase(Trim(pPassType_String)), 1) = "B") And Val(Mid(UCase(Trim(pPassType_String)), 2, 2)) >= 12 And Val(Mid(UCase(Trim(pPassType_String)), 2, 2)) <= 16) Or (Val(Trim(pPassType_String)) >= 12 And Val(Trim(pPassType_String)) <= 16) Then  'Youth means "C12-C16"
            ''by Abhi on 16-May-2017 for caseid 7439 Amadeus fare picking issue due to invalid value in Amdline1 and AmdlineK
            '    vPassType_String = "Youth"
            ''by Abhi on 16-May-2017 for caseid 7439 Amadeus fare picking issue due to invalid value in Amdline1 and AmdlineK
            ''ElseIf (Left(UCase(Trim(pPassType_String)), 1) = "J" Or Left(UCase(Trim(pPassType_String)), 1) = "C" Or Left(UCase(Trim(pPassType_String)), 1) = "I") And IsNumeric(Mid(UCase(Trim(pPassType_String)), 2, 2)) = True Then 'JNN means "J09"
            'ElseIf ((Left(UCase(Trim(pPassType_String)), 1) = "J" Or Left(UCase(Trim(pPassType_String)), 1) = "C" Or Left(UCase(Trim(pPassType_String)), 1) = "I" Or Left(UCase(Trim(pPassType_String)), 1) = "B") And IsNumeric(Mid(UCase(Trim(pPassType_String)), 2, 2)) = True) Or (IsNumeric(Trim(pPassType_String)) = True) Then 'JNN means "J09"
            ''by Abhi on 16-May-2017 for caseid 7439 Amadeus fare picking issue due to invalid value in Amdline1 and AmdlineK
            ''by Abhi on 04-Mar-2016 for caseid 6108 New pax type Youth from pnr Files-GDS Galileo
            '    vPassType_String = "Child"
            ''by Abhi on 24-Mar-2014 for caseid 3839 Passenger type -GDS File "CHD" "CNN", "JNN", "INN" (where NN is numberic) will be as "Child"
            'Else
            '    If pIfnotFoundSetBlank_Boolean = False Then
            '        vPassType_String = pPassType_String
            '    Else
            '        vPassType_String = ""
            '    End If
            'End If
            
            vPassTypeFirst_String = Left(UCase(Trim(pPassType_String)), 1)
            vPassTypeRest_String = Mid(UCase(Trim(pPassType_String)), 2, 2)
            
            If ( _
                (vPassTypeFirst_String = "J" Or vPassTypeFirst_String = "C" Or vPassTypeFirst_String = "I" Or vPassTypeFirst_String = "B" Or vPassTypeFirst_String = "T") _
                And _
                IsNumeric(Trim(vPassTypeRest_String)) = True And Val(vPassTypeRest_String) >= 12 And Val(vPassTypeRest_String) <= 16 _
               ) _
               Or _
               ( _
               IsNumeric(Trim(pPassType_String)) = True And Val(Trim(pPassType_String)) >= 12 And Val(Trim(pPassType_String)) <= 16 _
               ) Then  'Youth means "C12-C16"
                    vPassType_String = "Youth"
            ElseIf ( _
                    (vPassTypeFirst_String = "J" Or vPassTypeFirst_String = "C" Or vPassTypeFirst_String = "I" Or vPassTypeFirst_String = "B" Or vPassTypeFirst_String = "T") _
                    And _
                    IsNumeric(Trim(vPassTypeRest_String)) = True And Val(vPassTypeRest_String) >= 0 And Val(vPassTypeRest_String) <= 11 _
                   ) _
                   Or _
                   ( _
                    IsNumeric(Trim(pPassType_String)) = True And Val(Trim(pPassType_String)) >= 0 And Val(Trim(pPassType_String)) <= 11 _
                   ) Then 'JNN means "J00-J11"
                        vPassType_String = "Child"
            ElseIf ( _
                    (vPassTypeFirst_String = "J" Or vPassTypeFirst_String = "C" Or vPassTypeFirst_String = "I" Or vPassTypeFirst_String = "B" Or vPassTypeFirst_String = "T") _
                    And _
                    IsNumeric(Trim(vPassTypeRest_String)) = True And Val(vPassTypeRest_String) >= 17 _
                   ) _
                   Or _
                   ( _
                    IsNumeric(Trim(pPassType_String)) = True And Val(Trim(pPassType_String)) >= 17 _
                   ) Then 'JNN means "J17-J99..."
                        vPassType_String = "Adult"
            Else
                If pIfnotFoundSetBlank_Boolean = False Then
                    vPassType_String = pPassType_String
                Else
                    vPassType_String = ""
                End If
            End If
            
            'by Abhi on 12-Jul-2017 for caseid 7618 Pax type B17 picking issue from Amadeus file
            
            'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
    End Select

GetPassengerTypefromShort = vPassType_String
End Function

'by Abhi on 21-Apr-2010 for caseid 1320 NOFOLDER for in penline
Public Function isNOFOLDERExists(ByVal pFile_String As String) As Boolean
Dim fsObj As New FileSystemObject
Dim tsObj As TextStream
Dim vFileContent_String As String
Dim vNOFOLDERExists_Boolean As Boolean
    
    Set tsObj = fsObj.OpenTextFile(pFile_String, ForReading)
    If tsObj.AtEndOfStream <> True Then
        vFileContent_String = tsObj.ReadAll
    Else
        vFileContent_String = ""
    End If
    
    If InStr(1, vFileContent_String, "NOFOLDER", vbTextCompare) > 0 Then
        vNOFOLDERExists_Boolean = True
    End If
    
    'by Abhi on 04-Mar-2013 for caseid 2958 PenGDS Sabre pnr file not refeshing
    tsObj.Close
    Set tsObj = Nothing
isNOFOLDERExists = vNOFOLDERExists_Boolean
End Function

Function getFromFileTable(vlsFLD As String) As String
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
Dim rsSelect As New ADODB.Recordset
Dim vlsValue As String
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
Dim DeadlockRETRY_Integer As Integer
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
DeadlockRETRY:
    'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsSelect.Open "SELECT " & vlsFLD & " From [File] ;", dbCompany, adOpenForwardOnly, adLockReadOnly
    rsSelect.Open "SELECT " & vlsFLD & " From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
    If rsSelect.EOF = False Then
        vlsValue = SkipNull(rsSelect.Fields(vlsFLD))
    End If
    rsSelect.Close
    Set rsSelect = Nothing
    getFromFileTable = vlsValue
Exit Function

PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    'by Abhi on 24-Jul-2009 for Deadlock
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    'If PENErr_Number = -2147467259 Then 'Deadlock
    If (PENErr_Number = -2147467259 Or PENErr_Number = -2147217871) And DeadlockRETRY_Integer < 3 Then '-2147467259 Deadlock, -2147217871 Query timeout expired
        DeadlockRETRY_Integer = DeadlockRETRY_Integer + 1
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
        Debug.Print "Deadlock"
        'commented by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
        'Resume
        'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
        Sleep 5
        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
        'GoTo DeadlockRETRY
        Resume DeadlockRETRY
        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    End If
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " (getFromFileTable())"
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    Exit Function
End Function

Function getFromExecuted(ByVal pSQL As String, ByVal pFLD As String) As String
'by Abhi on 25-Apr-2013 for caseid 3057 Check Error handling is missed in bank receipt and Payment Module
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
Dim rsSelect As New ADODB.Recordset
Dim vValue As String
'by Abhi on 24-Dec-2014 for caseid 4848 2nd level Log to identify the issue of Refund reset not removing from Tacr
Dim DeadlockRETRY_Integer As Integer
'by Abhi on 24-Dec-2014 for caseid 4848 2nd level Log to identify the issue of Refund reset not removing from Tacr
    
    dbCompany.CommandTimeout = 300 '30 sec '''Caseid 834 binu 08/03/2010
    'rsSelect.Open pSQL, CnCompany, adOpenDynamic, adLockBatchOptimistic
'by Abhi on 25-Apr-2013 for caseid 3057 Check Error handling is missed in bank receipt and Payment Module
DeadlockRETRY:
    'by Abhi on 25-Apr-2013 for caseid 3057 Check Error handling is missed in bank receipt and Payment Module
    If rsSelect.State = 1 Then rsSelect.Close
    rsSelect.Open pSQL, dbCompany, adOpenForwardOnly, adLockReadOnly
    If rsSelect.EOF = False Then
        vValue = SkipNull(rsSelect.Fields(pFLD))
    End If
    'by Abhi on 25-Apr-2013 for caseid 3057 Check Error handling is missed in bank receipt and Payment Module
    If rsSelect.State = 1 Then rsSelect.Close
    Set rsSelect = Nothing
    getFromExecuted = vValue
Exit Function

PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    'by Abhi on 24-Jul-2009 for Deadlock
    'by Abhi on 24-Dec-2014 for caseid 4848 2nd level Log to identify the issue of Refund reset not removing from Tacr
    'If PENErr_Number = -2147467259 Then 'Deadlock
    'by Abhi on 30-Sep-2016 for caseid 6781 Dead lock error. Errr No.2147217900
    'If (PENErr_Number = -2147467259 Or PENErr_Number = -2147217871) And DeadlockRETRY_Integer < 3 Then '-2147467259 Deadlock, -2147217871 Query timeout expired
    'by Abhi on 27-Oct-2016 for caseid 6852 Deadlock error checking should check error description instead of error number
    'If (PENErr_Number = -2147467259 Or PENErr_Number = -2147217900 Or PENErr_Number = -2147217871) And DeadlockRETRY_Integer < 3 Then '-2147467259 Deadlock, -2147217900 Deadlock, -2147217871 Query timeout expired
    If (Deadlocked(PENErr_Description) = True Or PENErr_Number = -2147217871) And DeadlockRETRY_Integer < 3 Then '-2147217871 Query timeout expired
    'by Abhi on 27-Oct-2016 for caseid 6852 Deadlock error checking should check error description instead of error number
    'by Abhi on 30-Sep-2016 for caseid 6781 Dead lock error. Errr No.2147217900
        DeadlockRETRY_Integer = DeadlockRETRY_Integer + 1
    'by Abhi on 24-Dec-2014 for caseid 4848 2nd level Log to identify the issue of Refund reset not removing from Tacr
        'by Abhi on 25-Apr-2013 for caseid 3057 Check Error handling is missed in bank receipt and Payment Module
        Sleep 5
        'by Abhi on 16-Apr-2016 for caseid 6246 DeadlockRETRY Optimisation
        'GoTo DeadlockRETRY
        Resume DeadlockRETRY
        'by Abhi on 16-Apr-2016 for caseid 6246 DeadlockRETRY Optimisation
    End If
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " (getFromExecuted())"
    Exit Function
End Function

Public Function Deadlocked(ByVal pErrDescription_String As String) As Boolean
Dim vDeadlocked_Boolean As Boolean
'by Abhi on 28-Oct-2017 for caseid 7912 Deadlock between folder invoice and folderorder2cash
Dim vErrorLogLine_String As String
'by Abhi on 28-Oct-2017 for caseid 7912 Deadlock between folder invoice and folderorder2cash
    
    vDeadlocked_Boolean = Val(InStr(1, pErrDescription_String, "deadlocked", vbTextCompare))
    'by Abhi on 28-Oct-2017 for caseid 7912 Deadlock between folder invoice and folderorder2cash
    '"Transaction (Process ID 132) was deadlocked on lock resources with another process and has been chosen as the deadlock victim. Rerun the transaction."
    If vDeadlocked_Boolean = True Then
        vErrorLogLine_String = pErrDescription_String
        'Call ErrorLog("@Deadlocked@ " & vErrorLogLine_String)
    End If
    'by Abhi on 28-Oct-2017 for caseid 7912 Deadlock between folder invoice and folderorder2cash
Deadlocked = vDeadlocked_Boolean
End Function


'by Abhi on 06-Mar-2017 for caseid 7248 Special charcater in Currency code while saving PNR File
'Public Function SkipChars(TheString) As Variant
'On Error Resume Next
'SkipChars = Replace(TheString, "'", "''")
'End Function

'by Abhi on 01-Apr-2017 for caseid 7335 Folder invoice getting error String or binary data would be truncated due to skipchars
'Public Function SkipChars(TheString) As String
Public Function SkipChars(ByVal TheString As String) As String
'by Abhi on 01-Apr-2017 for caseid 7335 Folder invoice getting error String or binary data would be truncated due to skipchars
On Error Resume Next
    TheString = Replace(TheString, "'", "''")
SkipChars = TheString
End Function

'by Abhi on 01-Apr-2017 for caseid 7335 Folder invoice getting error String or binary data would be truncated due to skipchars
'Public Function SkipCharsNonPrintable(TheString) As String
Public Function SkipCharsNonPrintable(ByVal TheString As String) As String
'by Abhi on 01-Apr-2017 for caseid 7335 Folder invoice getting error String or binary data would be truncated due to skipchars
On Error Resume Next
    TheString = Replace(TheString, vbCrLf, "")
    TheString = Replace(TheString, vbCr, "")
    TheString = Replace(TheString, vbLf, "")
SkipCharsNonPrintable = TheString
End Function
'by Abhi on 06-Mar-2017 for caseid 7248 Special charcater in Currency code while saving PNR File


'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
Public Function getPENLINEID()
    PENLINEID_String = getFromFileTable("PENLINEID")
    If Trim(PENLINEID_String) <> "" Then
        PENLINEID_String = "P" & Trim(PENLINEID_String)
    End If
End Function

'by Abhi on 25-Sep-2010 for caseid 1505 will send a mail tp supoort
Public Function SendENDEmail() As Boolean
Dim vSendEMAIL As Boolean
Dim vSendEMAILTOEmailAddresses
Dim vSuccess As Boolean
Dim vSendEMAIL_Type As SendEMAIL_Type

    vSendEMAIL = Val(INIRead(INIPenAIR_String, "PenGDS", "SendEMAIL", "0"))
    vSendEMAILTOEmailAddresses = INIRead(INIPenAIR_String, "PenGDS", "SendEMAILTOEmailAddresses", "")
    
    vSendEMAIL_Type.TOEMAILs_String = vSendEMAILTOEmailAddresses
    vSendEMAIL_Type.CCEMAILs_String = ""
    vSendEMAIL_Type.BCCEMAILs_String = ""
    'vSendEMAIL_Type.SUBJECT_String = "ENDGDS.TXT in PenGDS[" & CHEAD & "] (" & vgsDatabase & ")"
    vSendEMAIL_Type.SUBJECT_String = "ENDGDS.TXT in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")"
    vSendEMAIL_Type.BODY_String = "Hi Support, 'ENDSQL.TXT' or 'ENDGDS.TXT' found. Please check ASAP."
    vSendEMAIL_Type.BODYFILENAME_String = ""
    vSendEMAIL_Type.FORMAT_SendEMAILFORMAT_EnumOptional = SendEMAILFORMATAUTO
    'vSendEMAIL_Type.FROMEMAIL_String = "PenGDS"
    vSendEMAIL_Type.FROMEMAIL_String = App.ProductName
    
    If vSendEMAIL = True And vSendEMAILTOEmailAddresses <> "" Then
        'vSuccess = CallDotNetSendEMAIL(vSendEMAILTOEmailAddresses, pSub, pMess, "", "PenGDS <support@penguininc.com>")
        vSuccess = CallDotNetSendEMAIL(vSendEMAIL_Type)
    End If

SendENDEmail = vSuccess
End Function

'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
Public Function DateFormat(ByVal pDATE_String As String) As String
    DateFormat = Format(pDATE_String, "DD/MMM/YYYY")
End Function

'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
Public Function TimeFormat(ByVal pTime_String As String) As String
    TimeFormat = Format(pTime_String, "HH:MM:SS")
End Function

'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
Public Function TimeFormat12HRS(ByVal pTime_String As String) As String
    TimeFormat12HRS = Format(pTime_String, "HH:MM:SS AMPM")
End Function

'by Abhi on 12-Mar-2012 for caseid 1652 PenGDS Permission denied added file closed checking
Public Function FileStatus(ByVal FileName As String) As FileStatus_Enum
    Dim intFile As Integer

    On Error Resume Next
    GetAttr FileName
    If Err.Number Then
        'FileStatus = vbUseDefault       'File doesn't exist or file server not available.
        FileStatus = FileStatusNotFound  'File doesn't exist or file server not available.
    Else
        Err.Clear
        intFile = FreeFile(0)
        Open FileName For Binary Lock Read Write As #intFile
        If Err.Number Then
            'FileStatus = vbFalse         'File already open.
            FileStatus = FileStatusOpened 'File already open.
        Else
            Close #intFile
            'FileStatus = vbTrue 'File available and not open by anyone.
            FileStatus = FileStatusClosed 'File available and not open by anyone.
        End If
    End If
End Function

'by Abhi on 12-Mar-2012 for caseid 1652 PenGDS Permission denied added file closed checking
'by Abhi on 12-Mar-2012 for caseid 1652 PenGDS Permission denied added file closed checking
'Public Function Wait4FileAvailable(ByVal pFileNamewPath_String As String)
'    Do While Not FileStatus(pFileNamewPath_String) = FileStatusClosed
'        FMain.stbUpload.Panels(1).Text = "Waiting for file available..."
'        DoEvents
'    Loop
'End Function
Public Function Wait4FileAvailable(ByVal pFileNamewPath_String As String) As Boolean
Dim vStart, vEnd, vDuration
Dim vTimedOut_Boolean As Boolean
'by Abhi on 03-Mar-2016 for caseid 6102 Error 0 Warning(IfExistsinTargetRename) in PenGDS - Worldspan
'Dim FileObj As New FileSystemObject
'by Abhi on 03-Mar-2016 for caseid 6102 Error 0 Warning(IfExistsinTargetRename) in PenGDS - Worldspan
    'testing by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    'Open pFileNamewPath_String For Binary Access Read Write Lock Read Write As #1
    'testing by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    vStart = Now
    'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
    If Dir(pFileNamewPath_String) = "" Then
        Call EventLog("#PenGDS# " & PadR(FMain.stbUpload.Panels(2).Text, 9) & " - " & PadR(LUFPNR_String, 6) & " - " & PadR(FMain.stbUpload.Panels(3).Text, 27) & ", Wait4FileAvailable Dir(pFileNamewPath_String) = """)
        vTimedOut_Boolean = False
        'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
        ErrDetails_String = "Cannot find the file."
        'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
        Exit Function
    End If
    'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
    
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    FMain.cmdStop.Enabled = True
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    
    Do While Not FileStatus(pFileNamewPath_String) = FileStatusClosed
        vEnd = Now
        vDuration = DateDiff("s", vStart, vEnd)
        'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
        'FMain.stbUpload.Panels(1).Text = "Waiting for file available..." & vDuration & " Sec(s)"
        FMain.stbUpload.Panels(1).Text = "File is in use, waiting..." & vDuration & " Sec(s)"
        'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
        DoEvents
        'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
        If Dir(pFileNamewPath_String) = "" Then
            Call EventLog("#PenGDS# " & PadR(FMain.stbUpload.Panels(2).Text, 9) & " - " & PadR(LUFPNR_String, 6) & " - " & PadR(FMain.stbUpload.Panels(3).Text, 27) & ", Wait4FileAvailable " & Val(vDuration) & " Sec(s) Dir(pFileNamewPath_String) = """)
            vTimedOut_Boolean = False
            'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
            ErrDetails_String = "Cannot find the file."
            'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
            Exit Function
        End If
        'by Abhi on 14-Mar-2016 for caseid 6151 Error 0 Warning(IfExistsinTargetRename) in PenGDS 2nd time
        If Val(vDuration) >= 90 Then
        'testing by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
        'If Val(vDuration) >= 10 Then
        'testing by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
            'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
            'vTimedOut_Boolean = True
            'Exit Do
            'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
            'Call EventLog("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & LUFPNR_String & " - " & FMain.stbUpload.Panels(3).Text & ", Wait4FileAvailable " & Val(vDuration) & " Sec(s)")
            Call EventLog("#PenGDS# " & PadR(FMain.stbUpload.Panels(2).Text, 9) & " - " & PadR(LUFPNR_String, 6) & " - " & PadR(FMain.stbUpload.Panels(3).Text, 27) & ", Wait4FileAvailable " & Val(vDuration) & " Sec(s)")
            'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
            vTimedOut_Boolean = False
            'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
            ErrDetails_String = "File is in use by another application and cannot be accessed."
            'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
            Exit Function
            'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
        End If
        'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
        If FMain.fForcedStop_Boolean = True Then
            vTimedOut_Boolean = False
            ErrDetails_String = "FMain.fForcedStop_Boolean = True"
            Exit Function
        End If
        'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    Loop
    
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    FMain.cmdStop.Enabled = False
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    
    'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
    vTimedOut_Boolean = True
    'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
    'testing by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    'Close #1
    'testing by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
    'Call EventLog("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & LUFPNR_String & " - " & FMain.stbUpload.Panels(3).Text & ", Wait4FileAvailable " & Val(vDuration) & " Sec(s)")
    Call EventLog("#PenGDS# " & PadR(FMain.stbUpload.Panels(2).Text, 9) & " - " & PadR(LUFPNR_String, 6) & " - " & PadR(FMain.stbUpload.Panels(3).Text, 27) & ", Wait4FileAvailable " & Val(vDuration) & " Sec(s), FileStatusClosed")
    'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
Wait4FileAvailable = vTimedOut_Boolean
End Function
'by Abhi on 12-Mar-2012 for caseid 1652 PenGDS Permission denied added file closed checking


'by Abhi on 12-Mar-2012 for caseid 1652 PenGDS Permission denied added file closed checking
Public Function SplitFileTitleandExtension(ByVal pFileName_String As String, ByRef pFileTitle_String As String, ByRef pFileExtension_String As String)
Dim temp
    temp = SplitForce(pFileName_String, ".", 2)
pFileTitle_String = temp(0)
If Trim(temp(1)) <> "" Then
    pFileExtension_String = "." & temp(1)
End If
End Function

'by Abhi on 12-Mar-2012 for caseid 1652 PenGDS Permission denied added file closed checking
Private Function SplitForce(Data, delimiter As String, minNos As Integer)
Dim temp() As String
Dim Nos As Integer
    temp = Split(Data, delimiter)
    Nos = UBound(temp) + 1
    If Nos < minNos Then
    ReDim Preserve temp(minNos)
    End If
SplitForce = temp
End Function

'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
Public Function ConnectExcelFile(vlsPath As String) As Boolean
On Error GoTo MyErr
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
Dim constr As String
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS

If Len(Trim(vlsPath)) > 0 Then
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vlsPath & ";Extended Properties=Excel 8.0;Persist Security Info=False"
    'constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vlsPath & ";Extended Properties=Excel 8.0;HDR=Yes;IMEX=1"
    If CnExcel.State = adStateOpen Then CnExcel.Close
    Set CnExcel = Nothing
    'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
    CnExcel.CursorLocation = adUseClient
    'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
    CnExcel.Open constr
    ConnectExcelFile = True
End If
Exit Function
MyErr:
    MsgBox Err.Description
    ConnectExcelFile = False
End Function

'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
Public Function MoneyFormat(ByVal pMoney_String As String) As String
    MoneyFormat = Format(Val(pMoney_String), "0.00")
End Function

'by Abhi on 01-Nov-2012 for caseid 2368 Sabre pnr missing
Public Function Wait4EOFSabre(ByVal pFileNamewPath_String As String)
    'by Abhi on 09-Nov-2012 for caseid 2368 Sabre pnr missing checking sequence changed
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    'FMain.stbUpload.Panels(1).Text = "Waiting for file available..."
    FMain.stbUpload.Panels(1).Text = "File is in use, waiting..."
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    Call Wait4FileAvailable(pFileNamewPath_String)
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    'FMain.stbUpload.Panels(1).Text = "Waiting for EOF..."
    FMain.stbUpload.Panels(1).Text = "File is in use, waiting for EOF..."
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    ReadAll_String = ""
    Do While InStr(1, ReadAll_String, "***EOM***", vbTextCompare) = 0
        Set TextStream_TextStream = fsObj.OpenTextFile(pFileNamewPath_String, ForReading)
        If TextStream_TextStream.AtEndOfStream <> True Then
            ReadAll_String = TextStream_TextStream.ReadAll
        Else
            ReadAll_String = ""
        End If
        Sleep 1000
        DoEvents
    Loop
    TextStream_TextStream.Close
End Function

'Function getFromTable(ByVal vlsTable As String, ByVal vlsFLD As String, ByVal vlsSerField As String, ByVal vlsSerData As String, Optional ByVal vlsSerConditions_wo_Where As String) As String
'On Error GoTo PENErr
'Dim PENErr_Number As String, PENErr_Description As String
'Dim rsSelect As New ADODB.Recordset
'Dim vlsValue As String
'DeadlockRETRY:
'    If rsSelect.State = 1 Then rsSelect.Close
'    If vlsSerConditions_wo_Where = "" Then
'        rsSelect.Open "SELECT     " & vlsFLD & " " _
'                        & "From " & vlsTable & " WITH (NOLOCK) " _
'                        & "WHERE     (" & vlsSerField & "= " & vlsSerData & ")", dbCompany, adOpenForwardOnly, adLockReadOnly
'    Else
'        rsSelect.Open "SELECT     " & vlsFLD & " " _
'                        & "From " & vlsTable & " WITH (NOLOCK) " _
'                        & "WHERE     (" & vlsSerConditions_wo_Where & ")", dbCompany, adOpenForwardOnly, adLockReadOnly
'    End If
'    If rsSelect.EOF = False Then
'        vlsValue = SkipNull(rsSelect.Fields(vlsFLD))
'    End If
'    rsSelect.Close
'    Set rsSelect = Nothing
'    getFromTable = vlsValue
'Exit Function
'
'PENErr:
'    PENErr_Number = Err.Number
'    PENErr_Description = Err.Description
'    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
'    'If PENErr_Number = -2147467259 Then 'Deadlock
'    If PENErr_Number = -2147467259 Or PENErr_Number = -2147217871 Then '-2147467259 Deadlock, -2147217871 Query timeout expired
'    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
'        Debug.Print "Deadlock"
'        Sleep 5
'        GoTo DeadlockRETRY
'    End If
'    Exit Function
'End Function


Public Function DateFormatBlankto1900(ByVal pDATE_String As String) As String
    If Trim(pDATE_String) = "" Then
        pDATE_String = "01/Jan/1900"
    End If
DateFormatBlankto1900 = DateFormat(pDATE_String)
End Function

'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
Public Function DateFormat1900toBlank(ByVal pDATE_String As String) As String
    If IsDate(pDATE_String) = True Then
        If DateFormat(pDATE_String) = "01/Jan/1900" Then
            pDATE_String = ""
        End If
    End If
    If Trim(pDATE_String) <> "" Then
        pDATE_String = DateFormat(pDATE_String)
    End If
DateFormat1900toBlank = pDATE_String
End Function
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date

Public Sub HelpService(ByVal pForm As Form, ByVal pHelpServiceURLType_Enum As HelpServiceURLType_Enum)
Dim vURL As String
Dim vURLTYPE_String As String
Dim vHelpId_String As String
Dim vPage_String As String
Dim vURLQueryString_String As String
Dim vHelpServiceURL_String As String

    ShowProgress "Opening ... Please wait"
    DoEvents
    vHelpServiceURL_String = INIRead(INIPenAIR_String, "General", "HelpServiceURL", "")
    If vHelpServiceURL_String <> "" Then
        vHelpId_String = "PenGDS." & pForm.Name
        If pHelpServiceURLType_Enum = HelpServiceURLType_Functional Then
            vURLTYPE_String = "F"
        Else
            vURLTYPE_String = "T"
        End If
        vURLQueryString_String = "?HelpId=" & vHelpId_String & "&LID=" & mUserLID & "&URLTYPE=" & vURLTYPE_String & ""
        vURL = vHelpServiceURL_String & vURLQueryString_String
        DoEvents
        vPage_String = InetOpenURL(vURL)
        DoEvents
        If vPage_String = "" Then
            vURLQueryString_String = "?HelpId=" & 0 & "&Language=" & mUserLID & "&URLTYPE=" & vURLTYPE_String & ""
            vURL = vHelpServiceURL_String & vURLQueryString_String
            vPage_String = InetOpenURL(vURL)
        End If
        DoEvents
        If vPage_String <> "" Then
            LaunchInNewwindow vPage_String
        End If
        DoEvents
        If vPage_String = "" Then
            MsgBox "Help not found for the screen '" & vHelpId_String & "'! Please contact Penguin.", vbCritical, App.Title & " (HelpService)"
        End If
    End If
    HideProgress
End Sub

Public Function InetOpenURL(ByVal sUrl As String) As String
    Dim hOpen               As Long
    Dim hOpenUrl            As Long
    Dim bDoLoop             As Boolean
    Dim bRet                As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim sBuffer             As String
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    
    'sUrl = "http://www.microsoft.com"
    
    hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    
    bDoLoop = True
    While bDoLoop
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend
    
    'Open "C:\Temp\log.txt" For Binary Access Write As #1
    'Put #1, , sBuffer
    'Close #1
    
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)

InetOpenURL = sBuffer
End Function

Public Function LaunchInNewwindow(sUrl As String) As Boolean
'Public Function LaunchInNewwindow(sURL As String, pSWEnum As SWEnum=SWEnum.SW_SHOWNORMAL) As Boolean
    Const SW_SHOWNORMAL = 1
    Dim lRetVal As Long
    Dim sTemp   As String
    Dim sBrowserExec As String
    
    sBrowserExec = GetBrowserExe        'get the exe
    sUrl = AddHTTP(sUrl)
    
    sBrowserExec = sBrowserExec & " toolbar=no"
    
    lRetVal = ShellExecute(GetDesktopWindow(), "open", sBrowserExec, sUrl, sTemp, SW_SHOWNORMAL)
    'lRetVal = ShellExecute(FMain.hWnd, "open", sUrl, "", sTemp, SW_SHOWNORMAL)           '1998/07/31 This works as well
    If lRetVal > 32 Then            ' OK
        LaunchInNewwindow = True
    End If
End Function

Private Function AddHTTP(sUrl As String) As String
' 2004/12/16 Function added by Larry Rebich using the DELL8500 while in Fort McDowell, AZ
' 2004/12/16 Add http:// is none
    Dim sTemp As String
    
    sTemp = sUrl
    If InStr(LCase$(sTemp), "https://") = 1 Then
        AddHTTP = sTemp
    ElseIf InStr(LCase$(sTemp), "http://") = 1 Then
        AddHTTP = sTemp
    Else
        AddHTTP = "http://" & sTemp
    End If

End Function

Public Function GetBrowserExe() As String
    Dim sFilename   As String
    Dim sBrowserExec As String * 255
    Dim lRetVal     As Long
    Dim iFN         As Integer
    Dim sTemp       As String
    
    sBrowserExec = Space(255)
    sFilename = App.Path & "\temphtm.HTM"
    
    iFN = FreeFile()                    ' Get unused file number
    
    Open sFilename For Output As #iFN   ' Create temp HTML file
    Print #iFN, "<HTML> <\HTML>"        ' Output text
    Close #iFN                          ' Close file
    
    ' Then find the application associated with it.
    lRetVal = FindExecutable(sFilename, sTemp, sBrowserExec)
    ' If an application return the name
    If lRetVal <= 32 Or IsEmpty(sBrowserExec) Then ' Error
    Else
        GetBrowserExe = Trim$(sBrowserExec)
    End If
    Kill sFilename  ' delete temp HTML file

End Function


'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
Public Function PENErr() As Boolean
Dim ErrNumber As String
Dim ErrDescription As String
Dim ErrDeadlockRETRY_Boolean As Boolean
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
Dim ret As VbMsgBoxResult
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
    
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    PermissionDenied = INIRead(INIPenAIR_String, "PenGDS", "PermissionDenied", "5") '5
    If Val(ErrNumber) = 70 And NoofPermissionDenied < PermissionDenied Then
        NoofPermissionDenied = NoofPermissionDenied + 1
        Sleep 500
        Resume
    'by Abhi on 24-Jul-2010 for caseid 1436 Deadlock in PenGDS
    'commented by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
    'ElseIf ErrNumber = -2147467259 Then 'Deadlock
    '    Debug.Print "Deadlock"
    '    Resume
    Else
        If Val(ErrNumber) = 0 Then
            'SendERROR "Warning in PenGDS[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ", PenGDS is automatically Resumed."
            'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
            'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ", " & App.ProductName & " is automatically Resumed."
            'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
            'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ", " & App.ProductName & " is automatically Resumed.", ErrNumber
            SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ErrDetails_String & ", " & App.ProductName & " is automatically Resumed.", ErrNumber
            ErrDetails_String = ""
            'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
            'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
            Resume
        Else
            'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
            'by Abhi on 26-Oct-2010 for caseid 1527 DeadlockRETRY
            'dbCompany.RollbackTrans
            'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
            If PENErr_BeginTrans = True Then
            'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
                If ErrNumber <> -2147168242 Then dbCompany.RollbackTrans
            'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
                PENErr_BeginTrans = False
            End If
            'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
            'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
            'If ErrNumber = -2147467259 Then 'Deadlock
            If (ErrNumber = -2147467259 Or ErrNumber = -2147217871) And GDSDeadlockRETRY_Integer < 3 Then '-2147467259 Deadlock, -2147217871 Query timeout expired
                GDSDeadlockRETRY_Integer = GDSDeadlockRETRY_Integer + 1
            'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
                Debug.Print "Deadlock"
                Sleep 5
                'GoTo DeadlockRETRY
                ErrDeadlockRETRY_Boolean = True
                PENErr = ErrDeadlockRETRY_Boolean
                Exit Function
            End If
            FMain.cmdStop_Click
            NoofPermissionDenied = 0
            'by Abhi on 25-Mar-2013 for caseid 2786 PenGDS Sabre PNR missing - SabreM1 row not found. we need to skip this file
            If Val(ErrNumber) = -2147220991 Then
                'SendERROR "Warning in PenGDS[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & ") Moved to folder 'Error M1 Missing'. PenGDS is automatically Resumed."
                'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & ") Moved to folder 'Error M1 Missing'. " & App.ProductName & " is automatically Resumed."
                'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & ") Moved to folder 'Error M1 Missing'. " & App.ProductName & " is automatically Resumed.", ErrNumber
                SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ErrDetails_String & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & ") Moved to folder 'Error M1 Missing'. " & App.ProductName & " is automatically Resumed.", ErrNumber
                ErrDetails_String = ""
                'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                'by Abhi on 07-Jul-2014 for caseid 4262 Sabre File not found unknown error
                'FMain.stbUpload.Panels(2).Text = ""
                'FMain.stbUpload.Panels(3).Text = ""
                'by Abhi on 07-Jul-2014 for caseid 4262 Sabre File not found unknown error
            Else
                'by Abhi on 15-May-2013 for caseid 3121 Pengds should not stuck on error it should move files to error folder for later checking
                If Trim(FMain.stbUpload.Panels(2).Text) = "" And Trim(FMain.stbUpload.Panels(3).Text) = "" Or Val(ErrNumber) = -2147217871 Then
                    'SendERROR "Warning in PenGDS[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & "). PenGDS is automatically Resumed."
                    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                    'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed."
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                    'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed.", ErrNumber
                    SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ErrDetails_String & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed.", ErrNumber
                    ErrDetails_String = ""
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                    'by Abhi on 07-Jul-2014 for caseid 4262 Sabre File not found unknown error
                    'FMain.stbUpload.Panels(2).Text = ""
                    'FMain.stbUpload.Panels(3).Text = ""
                    'by Abhi on 07-Jul-2014 for caseid 4262 Sabre File not found unknown error
                Else
                    'by Abhi on 15-May-2013 for caseid 3121 Pengds should not stuck on error it should move files to error folder for later checking
                    'Me.WindowState = vbNormal
                    'Me.Show
                    'SendStatus "Error"
                    'SendERROR "ERROR in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & ")"
                    If Dir(FMain.stbUpload.Panels("SOURCEPATH").Text & "\" & FMain.stbUpload.Panels("FILENAME").Text) <> "" Then
                        If Dir(FMain.stbUpload.Panels("SOURCEPATH").Text & "\Error Files\", vbDirectory) = "" Then
                            MkDir FMain.stbUpload.Panels("SOURCEPATH").Text & "\Error Files\"
                        End If
                        'by Abhi on 07-Jul-2014 for caseid 4258 Change in logic for file name checking when moving to target or error files
                        FMain.stbUpload.Panels("FILENAME").Text = IfExistsinTargetRename(FMain.stbUpload.Panels("SOURCEPATH").Text, FMain.stbUpload.Panels("FILENAME").Text, FMain.stbUpload.Panels("SOURCEPATH").Text & "\Error Files\")
                        'by Abhi on 07-Jul-2014 for caseid 4258 Change in logic for file name checking when moving to target or error files
                        fsObj.CopyFile FMain.stbUpload.Panels("SOURCEPATH").Text & "\" & FMain.stbUpload.Panels("FILENAME").Text, FMain.stbUpload.Panels("SOURCEPATH").Text & "\Error Files\", True
                        'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
                        'On Error Resume Next
                        'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
                        fsObj.DeleteFile FMain.stbUpload.Panels("SOURCEPATH").Text & "\" & FMain.stbUpload.Panels("FILENAME").Text, True
                    End If
                    'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
                    'On Error GoTo 0
                    'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
                    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                    'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & ") Moved to folder 'Error Files'. " & App.ProductName & " is automatically Resumed."
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                    'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & ") Moved to folder 'Error Files'. " & App.ProductName & " is automatically Resumed.", ErrNumber
                    SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ErrDetails_String & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & ") Moved to folder 'Error Files'. " & App.ProductName & " is automatically Resumed.", ErrNumber
                    ErrDetails_String = ""
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                    'by Abhi on 07-Jul-2014 for caseid 4262 Sabre File not found unknown error
                    'FMain.stbUpload.Panels(2).Text = ""
                    'FMain.stbUpload.Panels(3).Text = ""
                    'by Abhi on 07-Jul-2014 for caseid 4262 Sabre File not found unknown error
                End If
            End If
            'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
            'ret = MsgBox("Error: " & ErrNumber & vbCrLf & ErrDescription, vbAbortRetryIgnore, App.Title)
            'If ret = vbAbort Then
            '    ShowStopAccess = False
            '    cmdStop_Click
            'ElseIf ret = vbRetry Then
            '    Resume
            'ElseIf ret = vbIgnore Then
            '    Resume Next
            'End If
            'by Abhi on 25-Mar-2013 for caseid 2786 PenGDS Sabre PNR missing - SabreM1 row not found. we need to skip this file
            'by Abhi on 15-May-2013 for caseid 3121 Pengds should not stuck on error it should move files to error folder for later checking
            'If Val(ErrNumber) = -2147220991 Then
                ret = vbOK
            'Else
            '    ret = MsgBox("Error: " & ErrNumber & vbCrLf & ErrDescription, vbOKCancel + vbCritical, App.Title)
            'End If
            If ret = vbOK Then
                FMain.cmdStart_Click
            End If
        End If
    End If
End Function
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side

'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
Function GetEMAILBodyFileName(ByVal BODY_String As String) As String
Dim vEMAILBody_String As String
Dim vEMAILBodyFileName_String As String
Dim tsObj As TextStream
    
    vEMAILBody_String = TextToHTML(BODY_String)
    
    'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
    'vEMAILBodyFileName_String = App.Path & "\" & Replace(fsObj.GetTempName, ".tmp", ".htm")
    vEMAILBodyFileName_String = PENTempName(TempNameTypeFile, ".htm")
    'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
    Set tsObj = fsObj.CreateTextFile(vEMAILBodyFileName_String, True)
    tsObj.Write vEMAILBody_String
    tsObj.Close

GetEMAILBodyFileName = vEMAILBodyFileName_String
End Function

Public Function TextToHTML(ByVal pText_String As String) As String
Dim vHTML_String As String
Dim vFontName_String As String
Dim vFontSize_Integer As Integer
Dim vCurPos_Long As Long
Dim vLength_Long As Long
Dim vChar_String As String
Dim vDOUBLEChar_String As String
Dim vhrefTagFound_Integer As Integer
Dim vhrefText_String As String
Dim vhrefURL_String As String


    ' If nothing in the box, exit
    vLength_Long = Len(pText_String)
    If vLength_Long = 0 Then Exit Function
    
    vFontName_String = "Arial"
    vFontSize_Integer = 2
    
    vHTML_String = "<html>" & vbCrLf
    
    vHTML_String = vHTML_String & "<head>" & vbCrLf
    vHTML_String = vHTML_String & "<meta http-equiv=""Content-Language"" content=""en-us"">" & vbCrLf
    vHTML_String = vHTML_String & "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">" & vbCrLf
    vHTML_String = vHTML_String & "</head>" & vbCrLf
    
    vHTML_String = vHTML_String & "<body>" & vbCrLf
    
    vHTML_String = vHTML_String & "<font face=""" & vFontName_String & """ size=""" & vFontSize_Integer & """>" & vbCrLf
    For vCurPos_Long = 1 To vLength_Long
        vChar_String = Mid(pText_String, vCurPos_Long, 1)
        vDOUBLEChar_String = Mid(pText_String, vCurPos_Long, 2)
        
        If Asc(vChar_String) = 10 Then
        'Nothing
        ElseIf Asc(vChar_String) <> vbKeyReturn Then
            'by Abhi on 30-Nov-2010 for caseid 1561 PenEMAIL for Crystal Report
            If Asc(vChar_String) = vbKeyTab Then
                vHTML_String = vHTML_String & "&nbsp;&nbsp;&nbsp; "
            Else
                If vDOUBLEChar_String = "@@" Then
                    vCurPos_Long = vCurPos_Long + 1
                    vChar_String = Mid(pText_String, vCurPos_Long, 1)
                    vhrefTagFound_Integer = vhrefTagFound_Integer + 1
                    If vhrefTagFound_Integer = 3 Then
                        vHTML_String = vHTML_String & "<a href=""" & vhrefURL_String & """>" & vhrefText_String & "</a>"
                        vhrefTagFound_Integer = 0
                        vhrefText_String = ""
                        vhrefURL_String = ""
                    End If
                Else
                    If vhrefTagFound_Integer = 1 Then
                        vhrefText_String = vhrefText_String & vChar_String
                    ElseIf vhrefTagFound_Integer = 2 Then
                        vhrefURL_String = vhrefURL_String & vChar_String
                    Else
                        vHTML_String = vHTML_String & vChar_String
                    End If
                End If
            End If
        Else
            vHTML_String = vHTML_String & "<br>" & vbCrLf
        End If
        DoEvents
    Next
    vHTML_String = vHTML_String & "<br>" & vbCrLf
    vHTML_String = vHTML_String & "</body>" & vbCrLf
    
    vHTML_String = vHTML_String & "</font>" & vbCrLf
    vHTML_String = vHTML_String & "</html>" & vbCrLf

TextToHTML = vHTML_String
End Function
'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email

'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
Public Function PadC(ByVal sSrc As String, ByVal iLen As Integer, Optional ByVal sChr As String = " ") As String
Dim iLeft As Integer
Dim iRight As Integer

    If iLen > Len(sSrc) Then
        If iLen Mod 2 = 0 Then
            'This is an even length output string...
            iLeft = (iLen - Len(sSrc)) \ 2
        Else
            'This is an odd length output string...
            iLeft = ((iLen + 1) - Len(sSrc)) \ 2
        End If
        iRight = iLen - (iLeft + Len(sSrc))
        PadC = String(iLeft, sChr) & sSrc & String(iRight, sChr)
    Else
        PadC = sSrc
    End If

End Function

Public Function PadR(ByVal sSrc As String, ByVal iLen As Integer, Optional ByVal sChr As String = " ") As String
Dim x As Integer

    x = iLen - Len(sSrc)
    If x > 0 Then
        PadR = sSrc & String(x, sChr)
    Else
        PadR = sSrc
    End If

End Function

Public Function PadL(ByVal sSrc As String, ByVal iLen As Integer, Optional ByVal sChr As String = " ") As String
Dim x As Integer

    x = iLen - Len(sSrc)
    If x > 0 Then
        PadL = String(x, sChr) & sSrc
    Else
        PadL = sSrc
    End If

End Function
'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value

'by Abhi on 16-May-2016 for caseid 6375 PenGDS User’s standard Temp folder  should create before user’s session’s temporary folder
''by Abhi on 11-Jul-2015 for caseid 5393 PenAIR Run time error 76 path not found
'Public Function CreateUsersTemporaryFolder(Optional ByVal vShowMsgBox_Boolean As Boolean = False)
'Dim vMsgBoxPrompt_String As String
'
'    UsersTemporaryFolder_String = Environ$("temp")
'
'    PenEMAILTemporaryFolder_String = UsersTemporaryFolder_String
'    'by Abhi on 18-Nov-2015 for caseid 5798 PenGDS Email Temporary Folder should use user’s session’s temporary folders instead of user’s standard Temp folder
'    'PenEMAILTemporaryFolder_String = Left(PenEMAILTemporaryFolder_String, InStr(1, PenEMAILTemporaryFolder_String, "\Temp", vbTextCompare)) & "Temp"
'    'by Abhi on 18-Nov-2015 for caseid 5798 PenGDS Email Temporary Folder should use user’s session’s temporary folders instead of user’s standard Temp folder
'
'    'UsersTemporaryFolder_String = UsersTemporaryFolder_String & "\abc\78"
'    'PenEMAILTemporaryFolder_String = UsersTemporaryFolder_String
'    'PenEMAILTemporaryFolder_String = Left(PenEMAILTemporaryFolder_String, InStr(1, PenEMAILTemporaryFolder_String, "\abc", vbTextCompare)) & "abc"
'
'    If Dir(PenEMAILTemporaryFolder_String, vbDirectory) = "" Then
'        If vShowMsgBox_Boolean = True Then
'            MsgBox "Path not found '" & PenEMAILTemporaryFolder_String & "'" & vbCrLf & vbCrLf & "Press OK to create.", vbInformation, App.Title
'        End If
'        PenEMAILFileSystemObject.CreateFolder PenEMAILTemporaryFolder_String
'        vMsgBoxPrompt_String = "Path created     '" & PenEMAILTemporaryFolder_String & "'" & vbCrLf & vbCrLf
'    End If
'    If PenEMAILTemporaryFolder_String <> UsersTemporaryFolder_String Then
'        If Dir(UsersTemporaryFolder_String, vbDirectory) = "" Then
'            If vShowMsgBox_Boolean = True Then
'                MsgBox vMsgBoxPrompt_String & "Path not found '" & UsersTemporaryFolder_String & "'" & vbCrLf & vbCrLf & "Press OK to create.", vbInformation, App.Title
'            End If
'            PenEMAILFileSystemObject.CreateFolder UsersTemporaryFolder_String
'        End If
'    End If
'    'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
'    PenEMAILTemporaryFolder_String = PenEMAILTemporaryFolder_String & "\" & App.ProductName
'    'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
'    'by Abhi on 18-Nov-2015 for caseid 5798 PenGDS Email Temporary Folder should use user’s session’s temporary folders instead of user’s standard Temp folder
'    If Dir(PenEMAILTemporaryFolder_String, vbDirectory) = "" Then
'        PenEMAILFileSystemObject.CreateFolder PenEMAILTemporaryFolder_String
'    End If
'    'by Abhi on 18-Nov-2015 for caseid 5798 PenGDS Email Temporary Folder should use user’s session’s temporary folders instead of user’s standard Temp folder
'End Function
''by Abhi on 11-Jul-2015 for caseid 5393 PenAIR Run time error 76 path not found

Public Function CreateUsersTemporaryFolder(Optional ByVal vShowMsgBox_Boolean As Boolean = False)
Dim vMsgBoxPrompt_String As String

    'Desktop - User’s standard Temp folder
    'C:\Users\Abhi\AppData\Local\Temp
    'Server  - User’s session’s temporary folder
    'C:\Users\Abhi\AppData\Local\Temp\14
    
    UsersTemporaryFolder_String = Environ$("temp")
    
    'for testing
    'UsersTemporaryFolder_String = UsersTemporaryFolder_String & "\Temp"
    'UsersTemporaryFolder_String = UsersTemporaryFolder_String & "\Temp\14"
    'for testing
    
    PenEMAILTemporaryFolder_String = UsersTemporaryFolder_String
    
    If Dir(PenEMAILTemporaryFolder_String, vbDirectory + vbHidden) = "" Then
        vMsgBoxPrompt_String = DirsMsgBoxPrompt(PenEMAILTemporaryFolder_String)
        If vMsgBoxPrompt_String <> "" Then
            vMsgBoxPrompt_String = vMsgBoxPrompt_String & vbCrLf & "Press OK to create."
            MsgBox vMsgBoxPrompt_String, vbExclamation, App.Title
            MkDirs PenEMAILTemporaryFolder_String
        End If
    End If
    
    PenEMAILTemporaryFolder_String = PenEMAILTemporaryFolder_String & "\" & App.ProductName
    If Dir(PenEMAILTemporaryFolder_String, vbDirectory) = "" Then
        PenEMAILFileSystemObject.CreateFolder PenEMAILTemporaryFolder_String
    End If
End Function

Private Function DirsMsgBoxPrompt(ByVal pPath_String As String) As String
Dim vBackslashPosition_Long As Long
Dim vDirsMsgBoxPrompt_String As String
    
    If Right$(pPath_String, 1) <> "\" Then
        pPath_String = pPath_String & "\"
        vBackslashPosition_Long = InStr(1, pPath_String, "\")
    End If
    Do While vBackslashPosition_Long > 0
        If Dir$(Left$(pPath_String, vBackslashPosition_Long), vbDirectory + vbHidden) = "" Then
            vDirsMsgBoxPrompt_String = vDirsMsgBoxPrompt_String & Left$(pPath_String, vBackslashPosition_Long) & vbCrLf
        End If
        vBackslashPosition_Long = InStr(vBackslashPosition_Long + 1, pPath_String, "\")
    Loop
    If vDirsMsgBoxPrompt_String <> "" Then
        vDirsMsgBoxPrompt_String = "Windows Temporary folder path not found" & vbCrLf & vbCrLf & vDirsMsgBoxPrompt_String
    End If
DirsMsgBoxPrompt = vDirsMsgBoxPrompt_String
End Function

Public Function MkDirs(ByVal pPath_String As String) As Boolean
Dim vBackslashPosition_Long As Long
Dim vMkDirs_Boolean As Boolean
'by Abhi on 13-Apr-2018 for caseid 8504 Unicode Chinese font support in folder form
Dim vBackslashPosition4InStr_Long
'by Abhi on 13-Apr-2018 for caseid 8504 Unicode Chinese font support in folder form
   
    vMkDirs_Boolean = True  'assume success
    If Right$(pPath_String, 1) <> "\" Then
        pPath_String = pPath_String + "\"
        'by Abhi on 13-Apr-2018 for caseid 8504 Unicode Chinese font support in folder form
        'vBackslashPosition_Long = InStr(1, pPath_String, "\")
        vBackslashPosition4InStr_Long = 1
        If Left$(pPath_String, 2) = "\\" Then
            vBackslashPosition4InStr_Long = InStr(3, pPath_String, "\") + 1
        End If
        vBackslashPosition_Long = InStr(vBackslashPosition4InStr_Long, pPath_String, "\")
        'by Abhi on 13-Apr-2018 for caseid 8504 Unicode Chinese font support in folder form
    End If
    Do While vBackslashPosition_Long > 0
        If Dir$(Left$(pPath_String, vBackslashPosition_Long), vbDirectory + vbHidden) = "" Then
            MkDir Left$(pPath_String, vBackslashPosition_Long)
        End If
        vBackslashPosition_Long = InStr(vBackslashPosition_Long + 1, pPath_String, "\")
    Loop
MkDirs = vMkDirs_Boolean
End Function

'by Abhi on 16-May-2016 for caseid 6375 PenGDS User’s standard Temp folder  should create before user’s session’s temporary folder

'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer
Public Function PENTempName(ByVal pTempNameTypeEnum As TempNameTypeEnum, Optional ByVal pExtension_String As String = ".tmp") As String
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
Dim vPENTempName_String As String
Dim vVbFileAttribute As VbFileAttribute

    vVbFileAttribute = pTempNameTypeEnum
ErrorPathnotfoundRETRY:
    Do
        If Dir(PenEMAILTemporaryFolder_String, vbDirectory) = "" Then
            PenEMAILFileSystemObject.CreateFolder PenEMAILTemporaryFolder_String
        End If
        
        With CreateObject("scriptlet.typelib")
            vPENTempName_String = LCase(Replace(Mid(.Guid, 2, 36), "-", ""))
        End With
        If vVbFileAttribute <> VbFileAttribute.vbDirectory Then
            vPENTempName_String = vPENTempName_String & pExtension_String
            vVbFileAttribute = vbNormal
        End If
        vPENTempName_String = PenEMAILTemporaryFolder_String & "\" & vPENTempName_String
    Loop Until Dir(vPENTempName_String, vVbFileAttribute) = ""

PENTempName = vPENTempName_String
Exit Function
PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    If PENErr_Number = 76 Then
        Call CreateUsersTemporaryFolder
        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
        'GoTo ErrorPathnotfoundRETRY
        Resume ErrorPathnotfoundRETRY
        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    End If
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " (" & PENTempName & ")"
    PENTempName = ""
    Exit Function
End Function
'by Abhi on 05-Aug-2015 for caseid 5400 sending two etickets to one customer

'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
Function FieldExists(ByVal rs, ByVal fieldName) As Boolean
On Error Resume Next
    FieldExists = rs.Fields(fieldName).Name <> ""
    If Err <> 0 Then FieldExists = False
    Err.Clear
End Function

'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
'Function KeyExists(ByVal pCollection As Collection, ByVal fieldName) As Boolean
'On Error Resume Next
'    KeyExists = pCollection(fieldName) <> ""
'    If Err <> 0 Then KeyExists = False
'    Err.Clear
'End Function
Public Function KeyExists(Col, Index) As Boolean
On Error GoTo ExistsTryNonObject
    Dim O As Object

    Set O = Col(Index)
    KeyExists = True
    Exit Function

ExistsTryNonObject:
    KeyExists = KeyExistsNonObject(Col, Index)
End Function

Private Function KeyExistsNonObject(Col, Index) As Boolean
On Error GoTo ExistsNonObjectErrorHandler
    Dim v As Variant

    v = Col(Index)
    KeyExistsNonObject = True
    Exit Function

ExistsNonObjectErrorHandler:
    KeyExistsNonObject = False
End Function
'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO


'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
Public Function FindInitialAndName(mData)
On Error GoTo errPara
Dim aa, Data
Data = mData
Dim retn(1) As String
'by Abhi on 20-Apr-2015 for caseid 5145 Worldspan -Flat file loading issue
Dim rsSelect As New ADODB.Recordset
Dim vRecordCount_Long As Long
Dim vi_Long As Long
Dim vTITLE_String As String
Dim vTITLELen_Integer As Integer
'by Abhi on 20-Apr-2015 for caseid 5145 Worldspan -Flat file loading issue

Data = Replace(Data, "*CHD", "")
'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
Data = Replace(Data, "*YTH", "")
'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
'by Abhi on 20-Apr-2015 for caseid 5145 Worldspan -Flat file loading issue
'aa = UCase(Right(Data, 2))
'If (aa = "MR") Then
'    retn(0) = Mid(Data, 1, Len(Data) - 2)
'    retn(1) = "MR"
'End If
'
'aa = UCase(Right(Data, 3))
'If (aa = "MRS") Then
'    retn(0) = Mid(Data, 1, Len(Data) - 3)
'    retn(1) = "MRS"
'End If
'
'aa = UCase(Right(Data, 4))
'If (aa = "MISS") Then
'    retn(0) = Mid(Data, 1, Len(Data) - 4)
'    retn(1) = "MISS"
'End If
'
'aa = UCase(Right(Data, 4))
'If (aa = "MSTR") Then
'    retn(0) = Mid(Data, 1, Len(Data) - 4)
'    retn(1) = "MSTR"
'End If
'
'aa = UCase(Right(Data, 4))
'If (aa = "PROF") Then
'    retn(0) = Mid(Data, 1, Len(Data) - 4)
'    retn(1) = "PROF"
'End If
rsSelect.Open "SELECT TITLE FROM TitleMaster WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
vRecordCount_Long = rsSelect.RecordCount
For vi_Long = 1 To vRecordCount_Long
    vTITLE_String = Trim(SkipNull(rsSelect.Fields("TITLE")))
    vTITLELen_Integer = Len(vTITLE_String)
    If UCase(vTITLE_String) = UCase(Right(Data, vTITLELen_Integer)) Then
        'by Abhi on 27-Jul-2015 for caseid 5437 World span title separation issue
        'retn(0) = Replace(Data, vTITLE_String, "")
        retn(0) = Left(Data, Len(Data) - vTITLELen_Integer)
        'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
        retn(0) = Trim(retn(0))
        'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
        'by Abhi on 27-Jul-2015 for caseid 5437 World span title separation issue
        retn(1) = vTITLE_String
        'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
        retn(1) = Trim(retn(1))
        'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
        Exit For
    End If
    DoEvents
    rsSelect.MoveNext
Next
rsSelect.Close
Set rsSelect = Nothing
'by Abhi on 20-Apr-2015 for caseid 5145 Worldspan -Flat file loading issue

If retn(0) = "" Then
    retn(0) = Data
    'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
    retn(0) = Trim(retn(0))
    'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
    retn(1) = ""
End If

FindInitialAndName = retn
Exit Function
errPara:

If retn(0) = "" Then
    retn(0) = Data
    'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
    retn(0) = Trim(retn(0))
    'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
    retn(1) = ""
End If
End Function
'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER

'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
Public Function DateTime12hrsFormat(ByVal pDATEandTime_String As String) As String
    pDATEandTime_String = Trim(pDATEandTime_String)
    pDATEandTime_String = Replace(pDATEandTime_String, ".", "/", , , vbTextCompare)
    DateTime12hrsFormat = Format(pDATEandTime_String, "DD/MMM/YYYY hh:nn:ss AM/PM") '"DD/MMM/YYYY" Penguin Date Format
End Function

Public Function NowFormat(ByVal pNow_String As String) As String
    NowFormat = Format(pNow_String, "dd/mmm/yyyy h:mm:ss AMPM")
End Function
'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS

'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
Public Function HelpServiceGetServiceURL(ByVal pService_String As String, Optional ByVal pShowProgress_Boolean As Boolean = True) As String
Dim vURL As String
Dim vURLTYPE_String As String
Dim vHelpId_String As String
Dim vPage_String As String
Dim vURLQueryString_String As String
Dim vHelpServiceURL_String As String

    If Trim(pService_String) = "" Then Exit Function
    
    If pShowProgress_Boolean = True Then
        ShowProgress "Opening ... Please wait"
    End If
    DoEvents
    vHelpServiceURL_String = INIRead(INIPenAIR_String, "General", "HelpServiceURL", "")
    If vHelpServiceURL_String <> "" Then
        vURLTYPE_String = "F"
        vURLQueryString_String = "?HelpId=" & pService_String & "&LID=EN&URLTYPE=" & vURLTYPE_String & ""
        vURL = vHelpServiceURL_String & vURLQueryString_String
        DoEvents
        vPage_String = InetOpenURL(vURL)
        DoEvents
    End If
    If pShowProgress_Boolean = True Then
        HideProgress
    End If
    

HelpServiceGetServiceURL = vPage_String
End Function
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS

'by Abhi on 08-Mar-2018 for caseid 8331 Amadeus,Worldspan and Sabre - Pengds penline email replacement
Public Function PenlineEmailReplacement(ByVal PEMAIL_String As String) As String
    PEMAIL_String = Trim(PEMAIL_String)
    'Galileo, Amadeus
    PEMAIL_String = Replace(PEMAIL_String, "//", "@", , , vbTextCompare)
    PEMAIL_String = Replace(PEMAIL_String, "?", "_", , , vbTextCompare)
    'Sabre
    PEMAIL_String = Replace(PEMAIL_String, "*", "@", , , vbTextCompare)
PenlineEmailReplacement = PEMAIL_String
End Function
'by Abhi on 08-Mar-2018 for caseid 8331 Amadeus,Worldspan and Sabre - Pengds penline email replacement

'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
Public Function mFnCheckValidDate(SDate As String, Optional ByVal pMsgBoxPromptOptional_String As String, Optional ByVal pShowMsgBoxOptional_Boolean As Boolean = True) As Boolean
Dim dd As Date
mFnCheckValidDate = False
    If IsDate(SDate) Then
        If CDate(SDate) < CDate("01/01/1900") Or CDate(SDate) > CDate("31/12/9999") Then
            If pShowMsgBoxOptional_Boolean = True Then
                MsgBox "Invalid Date.           " & pMsgBoxPromptOptional_String, vbCritical, App.Title
            End If
            Exit Function
        End If
    Else
        If pShowMsgBoxOptional_Boolean = True Then
            MsgBox "Invalid Date.           " & pMsgBoxPromptOptional_String, vbCritical, App.Title
        End If
        Exit Function
    End If
mFnCheckValidDate = True
End Function

Public Function mSetFocus(ByVal TheObj As Object, Optional ElseObj As Object)
On Error Resume Next
If TheObj.Visible = True And TheObj.Enabled = True Then
    TheObj.SetFocus
ElseIf Not ElseObj Is Nothing Then
    If ElseObj.Visible = True And ElseObj.Enabled = True Then
        ElseObj.SetFocus
    End If
End If
End Function
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
