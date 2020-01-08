Attribute VB_Name = "MainModules"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public connstr As String
Public mPathComp As String

Public dbCompany As ADODB.Connection
Public rsCompany As New ADODB.Recordset

Public fsObj As New FileSystemObject
Public vgsBackend As String
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


Public Sub Main()
    If App.PrevInstance = True Then
        End
    End If
    Set dbCompany = New ADODB.Connection
    vgsBackend = INIRead(App.Path & "\PenSoft.ini", "General", "Backend", "SQL Server")
    vgsServer = INIRead(App.Path & "\PenSoft.ini", "General", "Server", "(Local)")
    vgsCompany = INIRead(App.Path & "\PenSoft.ini", "General", "Company", "PENCOMPANY")
    vgsUID = INIRead(App.Path & "\PenSoft.ini", "General", "UID", "PenSoft")
    vgsPWD = INIRead(App.Path & "\PenSoft.ini", "General", "PWD", "penac")
    vglServers = Val(INIRead(App.Path & "\PenSoft.ini", "General", "Servers", "0"))
    vglDatabases = Val(INIRead(App.Path & "\PenSoft.ini", "General", "Databases", "0"))
    vgsDatabase = INIRead(App.Path & "\PenSoft.ini", "General", "Database", "PENDEMO")
    
    If vglServers > 0 Then
        FServers.Caption = "PenAIR Servers"
        FServers.lstServers.Clear
        FServers.lstServers.AddItem vgsServer
        For vgliSer = 1 To vglServers
             FServers.lstServers.AddItem INIRead(App.Path & "\PenSoft.ini", "General", "Server" & vgliSer, "")
        Next
        FServers.lstServers.ListIndex = 0
        FServers.Show 1
    End If
    
    If vglDatabases > 0 Then
        FServers.Caption = "PenAIR Databases"
        FServers.lstServers.Clear
        FServers.lstServers.AddItem vgsDatabase
        For vgliSer = 1 To vglDatabases
             FServers.lstServers.AddItem INIRead(App.Path & "\PenSoft.ini", "General", "Database" & vgliSer, "")
        Next
        FServers.lstServers.ListIndex = 0
        FServers.Show 1
    End If
    
    'mPathComp = vgsCompany 'App.Path & "\Company.dat"
    mPathComp = vgsDatabase '"PENCOMPANY"
    
    With dbCompany
        .CursorLocation = adUseClient
        .CommandTimeout = 4
    End With
    
    'DBEngine.RegisterDatabase "PenSoftCompany", "Microsoft Access Driver (*.mdb)", True, "DBQ=" & mPathComp & ";PWD=penac"
    DBEngine.RegisterDatabase "PenSoftCompany", vgsBackend, True, "Description=PenSoftCompany" & vbCr & "SERVER=" & vgsServer & vbCr & "DATABASE=" & mPathComp & vbCr & "Network=DBMSSOCN" '& vbCr '& "Trusted_Connection=Yes" & vbCr & "UID=" & vgsUID & vbCr & "PWD=" & vgsPWD & vbCr
    'DBEngine.RegisterDatabase "PenSoftCompany", "SQL Server", True, "Description=PenSoftCompany" & vbCr & "SERVER=" & vgsServer & vbCr & "DATABASE=" & mPathComp & vbCr & "Network=DBMSSOCN" '& vbCr '& "Trusted_Connection=Yes" & vbCr & "UID=" & vgsUID & vbCr & "PWD=" & vgsPWD & vbCr
    dbCompany.Open "PenSoftCompany", vgsUID, vgsPWD
    If rsCompany.State = 1 Then rsCompany.Close
    rsCompany.Open "SELECT * FROM PEN0000 WHERE CDEFAULT = 'T' ", dbCompany, adOpenDynamic, adLockBatchOptimistic, 1
    If Not rsCompany.EOF Then
        'mPathComp = rsCompany!CPATH
        mPathComp = vgsDatabase
    End If
    ConnectDataBaseAccounts (mPathComp)
    If rsSelect.State = 1 Then rsSelect.Close
    rsSelect.Open "Select SABRESTATUS,STATUS,AUTOSTART,UPLOADDIRNAME,DESTDIRNAME,SABREUPLOADDIRNAME,SABREDESTDIRNAME,AMDUPLOADDIRNAME,AMDDESTDIRNAME,AMDSTATUS,WSPUPLOADDIRNAME,WSPDESTDIRNAME,WSPSTATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    If rsSelect.EOF = False Then
        FMain.txtGalilieoSource = SkipNull(rsSelect!UPLOADDIRNAME)
        FMain.txtGalilieoDest = SkipNull(rsSelect!DESTDIRNAME)
        FMain.chkSabre = Val(SkipNull(rsSelect!SABRESTATUS))
        
        FMain.txtSource = SkipNull(rsSelect!SABREUPLOADDIRNAME)
        FMain.txtDest = SkipNull(rsSelect!SABREDESTDIRNAME)
        FMain.chkGalilieo = Val(SkipNull(rsSelect!Status))
        
        FMain.txtWorldspanSource = SkipNull(rsSelect!WSPUPLOADDIRNAME)
        FMain.txtWorldspanDest = SkipNull(rsSelect!WSPDESTDIRNAME)
        FMain.chkWorldspan = Val(SkipNull(rsSelect!WSPSTATUS))
        
        FMain.txtAmadeusSource = SkipNull(rsSelect!AMDUPLOADDIRNAME)
        FMain.txtAmadeusDest = SkipNull(rsSelect!AMDDESTDIRNAME)
        FMain.chkAmadeus = Val(SkipNull(rsSelect!AMDSTATUS))
        
        FMain.chkAutoStart = Val(SkipNull(rsSelect!AutoStart))
    End If
    FMain.Show
    DoEvents
    If FMain.chkAutoStart = vbChecked Then
        FMain.cmdStart_Click
    End If
End Sub
Public Function mSeqNumberGen(Field As String)
On Error GoTo myerror
    Dim vlno As Long
    Dim rsUpdateCode As New ADODB.Recordset
abc:
        Query = "Select " & Field & " as NextNumber from PEN0001 where yearid =(select Max(cast(yearid as bigint)) from PEN0001) "
        If rsUpdateCode.State = 1 Then rsUpdateCode.Close
        rsUpdateCode.Open Query, dbCompany, adOpenDynamic, adLockOptimistic
        If rsUpdateCode.EOF = False Then
            vlno = Val(IIf(IsNull(rsUpdateCode!NextNumber), 1, rsUpdateCode!NextNumber))
            rsUpdateCode.Fields("NextNumber") = vlno + 1
            rsUpdateCode.Update
            mSeqNumberGen = vlno
        End If
Exit Function
myerror:
    If Err <> 0 Then
        MsgBox "The microsoft database engine stopped the process because you and another user are attempting to change the same data at the same time" & vbCrLf & "Press OK to Continue"
        rsUpdateCode.CancelUpdate
        GoTo abc
    End If
End Function
Public Sub ConnectDataBaseAccounts(mPath)
Set dbCompany = New ADODB.Connection
mPath = mPath
DBEngine.RegisterDatabase "PenSoft", "SQL Server", True, "Description=PenSoft" & vbCr & "SERVER=" & vgsServer & vbCr & "DATABASE=" & mPath & vbCr
dbCompany.Open "PenSoft", vgsUID, vgsPWD
End Sub

Public Function INIRead(sINIFile As String, sSection As String, sKey As String, sDefault As String) As String
    Dim sTemp As String * 256
    Dim nLength As Integer
    sTemp = Space$(256)
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)
    INIRead = Left$(sTemp, nLength)
End Function

Public Function SkipNull(value, Optional default = Empty)
    Dim A
    A = IIf(IsNull(value) = True, default, value)
    SkipNull = A
End Function

Public Sub DeleteFileToRecycleBin(fileName As String)

'Kill fileName
  Dim fop As SHFILEOPTSTRUCT

  With fop
    .wFunc = FO_DELETE
    .pFrom = fileName
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
