VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FWorldSpan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "World Span"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1965
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   120
         TabIndex        =   14
         Top             =   1980
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Upload"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   1980
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.FileListBox File1 
         Height          =   1065
         Left            =   5160
         TabIndex        =   11
         Top             =   300
         Width           =   1695
      End
      Begin VB.CommandButton cmdDirUpload 
         Caption         =   "DirUpload"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   1140
         Width           =   1095
      End
      Begin VB.TextBox txtDirName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1680
         TabIndex        =   9
         Top             =   300
         Width           =   3135
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Files Delete"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   1140
         Width           =   1095
      End
      Begin VB.TextBox txtRelLocation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1680
         TabIndex        =   7
         Top             =   660
         Width           =   3135
      End
      Begin VB.CheckBox chkGal 
         Caption         =   "Include"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   390
         TabIndex        =   6
         Top             =   1140
         Width           =   1365
      End
      Begin MSComctlLib.ProgressBar Pgr 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1620
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         Caption         =   " Dir. Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   " Rel. Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   675
         Width           =   1440
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UpLoad"
      Height          =   345
      Left            =   3900
      TabIndex        =   3
      Top             =   2160
      Width           =   1275
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "::"
      Height          =   285
      Left            =   6120
      TabIndex        =   1
      Top             =   5280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1710
      TabIndex        =   0
      Top             =   5280
      Width           =   4395
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   750
      Top             =   5610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "FileName"
      Height          =   255
      Left            =   780
      TabIndex        =   2
      Top             =   5310
      Width           =   855
   End
End
Attribute VB_Name = "FWorldSpan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim dbCompany As New ADODB.Connection
Dim FNAME As String
Public RECIDQHeader As Boolean

Private Sub CmdClear_Click()
ClearAll
End Sub

Private Sub CmdDelete_Click()
'On Error GoTo Note
'    Dim Count As Integer
'    File1.Path = txtDirName.Text
'    File1.Pattern = "*.mir"
'
'    If txtDirName.Text = "" Or File1.ListCount = 0 Then MsgBox "AIR Files Not Found", vbInformation: Exit Sub
'    If MsgBox("Do you want to Delete the " & File1.ListCount & " MIR Files ?", vbYesNo) = vbYes Then
'        For Count = 0 To File1.ListCount - 1
'            txtFileName.Text = File1.Path & "\" & File1.List(Count)
'            DeleteFileToRecycleBin (txtFileName)
'        Next
'    End If
'Exit Sub
'Note:
'    MsgBox Err.Description & ". Invalid Drive Location", vbCritical: Exit Sub
'
End Sub

Public Sub cmdDirUpload_Click()
'On Error GoTo Note
    Dim Count As Integer
    
    FMain.SendStatus FMain.SSTab1.TabCaption(3)
    File1.Path = txtDirName.Text
    File1.Pattern = FMain.txtWorldspanExt
    If txtDirName.Text = "" Or File1.ListCount = 0 Then FMain.stbUpload.Panels(1).Text = "MIR Files Not Found...": Exit Sub
    'Uploading Each File
    Open (App.Path & "\_UploadingSQL_") For Random As #1
    Close #1
    FMain.cmdStop.Enabled = False
    DoEvents
    FMain.lblFName.Caption = "Uploading Worldspan..."
    FMain.stbUpload.Panels(1).Text = "Reading..."
    FMain.stbUpload.Panels(2).Text = "Worldspan"
    FMain.prgUpload.Max = IIf((File1.ListCount = 1), File1.ListCount, File1.ListCount - 1)
    Pgr.value = 5
    
    'by Abhi on 18-May-2011 for caseid 1757 Added Events for GDSAuto
    Call EventLogDelete
    'by Abhi on 06-Sep-2017 for caseid 7766 Penline and Company currency change in penair is not reflecting in pengds
    Call getPENLINEID
    COMCID_String = getFromFileTable("COMCID")
    'by Abhi on 06-Sep-2017 for caseid 7766 Penline and Company currency change in penair is not reflecting in pengds
    For Count = 0 To File1.ListCount - 1
        'by Abhi on 17-Nov-2011 for caseid 1551 PenGDS and GDSAuto Monitoring in PenAIR
        INIWrite INIPenAIR_String, "PenGDS", "PenGDSACTIVE", Now
        DoEvents
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        ErrDetails_String = ""
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
        LUFPNR_String = ""
        'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
        FMain.stbUpload.Panels(3).Text = File1.List(Count)
        txtFileName.Text = File1.Path & "\" & File1.List(Count)
        txtFileName.Text = IfExistsinTargetRename(File1.Path, File1.List(Count), txtRelLocation.Text)
        'by Abhi on 07-Oct-2010 for caseid 1516 PenGDS Amadeus slow reading
        'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
        'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
        'If (txtFileName.Text) = "" Then Exit Sub
        If Trim(txtFileName.Text) <> "" Then
        'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
            'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
            FMain.stbUpload.Panels(1).Text = "Reading..."
            FNAME = txtFileName.Text
            'FMain.SendStatus FNAME
            txtFileName.Text = File1.Path & "\" & txtFileName.Text
            AddFile txtFileName.Text, FNAME
        'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
        End If
        'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
        If Pgr.value < 90 Then
            Pgr.value = Pgr.value + 10
        End If
        FMain.prgUpload.value = Count
        DoEvents
    Next
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    If FMain.fForcedStop_Boolean = True Then
        Exit Sub
    End If
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    Pgr.value = 100
    File1.Refresh
    FMain.prgUpload.value = 0
    FMain.stbUpload.Panels(1).Text = "Process Completed."
    FMain.cmdStop.Enabled = True
    FMain.stbUpload.Panels(2).Text = ""
    FMain.stbUpload.Panels(3).Text = ""
    'by Abhi on 16-Mar-2012 for caseid 2151 change waiting to Active - Waiting for GDS files... for pengds
    FMain.stbUpload.Panels(1).Text = "Active - Waiting for GDS files..."
    'FMain.txtGalilieoSource = ""
    'FMain.txtGalilieoDest = ""
    FMain.lblFName.Caption = ""
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    'fsObj.DeleteFile (App.Path & "\_UploadingSQL_"), True
    If Dir(App.Path & "\_UploadingSQL_") <> "" Then
        fsObj.DeleteFile (App.Path & "\_UploadingSQL_"), True
    End If
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    FMain.SendStatus "Started"
    Me.Hide
Exit Sub
Note:
    Me.Hide
    Exit Sub
    Resume
End Sub

Private Sub Command1_Click()
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then
Exit Sub
End If
Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
AddFile Text1.Text
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
'dbCompany.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=sa;Initial Catalog=PENDATA001;Data Source=PEN1\PENSOFT"
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'On Error GoTo Note
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'Dim rsSelect As New ADODB.Recordset
    Me.Icon = FMain.Icon
    'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsSelect.Open "Select WSPUPLOADDIRNAME,WSPDESTDIRNAME From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsSelect.Open "Select WSPUPLOADDIRNAME,WSPDESTDIRNAME From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
    'If rsSelect.EOF = False Then
    '    txtDirName = IIf(IsNull(rsSelect!WSPUPLOADDIRNAME), "", rsSelect!WSPUPLOADDIRNAME)
    '    txtRelLocation = IIf(IsNull(rsSelect!WSPDESTDIRNAME), "", rsSelect!WSPDESTDIRNAME)
    'End If
    'rsSelect.Close
    'If rsSelect.State = 1 Then rsSelect.Close
    ''by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    ''rsSelect.Open "Select WSPSTATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    'rsSelect.Open "Select WSPSTATUS From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
    'chkGal.value = IIf((IsNull(rsSelect!WSPSTATUS) = True), 0, rsSelect!WSPSTATUS)
    txtDirName = FMain.txtWorldspanSource
    txtRelLocation = FMain.txtWorldspanDest
    chkGal.value = FMain.chkWorldspan
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    
    File1.Path = txtDirName.Text
    File1.Refresh
    File1.Pattern = FMain.txtWorldspanExt
    File1.Refresh
    Exit Sub
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'Note:
'    If Me.Visible = True Then Unload Me
End Sub

Private Sub Text1_Change()
Caption = Len(Text1)
End Sub



Public Function AddFile(FileName As String, Optional fTitle = "")
'On Error GoTo Note
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
On Error GoTo PENErr
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
Dim id As Long
Dim head As String
Dim temp
Dim UploadNo As Long
Dim LineStr  As String
Dim tsObj As TextStream
Dim Lines, crrItem As String
Dim FileObj As New FileSystemObject
Dim HeaderChecked As Boolean
'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
Dim vSQL_String As String
'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
Dim vDateLastModified_String As String
'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
Dim vFileDateCreated_String As String
'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS

    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    GDSDeadlockRETRY_Integer = 0
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS

'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
DeadlockRETRY:
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    WorldSpan_AirFareSLNO = 0
    'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
    LUFPNR_String = ""
    'by Abhi on 25-Mar-2011 for caseid 1652 PenGDS Permission denied
    'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    'Sleep 2000
    'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    NoofPermissionDenied = 0
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    GIT_CUSTOMERUSERCODE_String = ""
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
    GIT_PENWAIT_String = "N"
    'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
    
    'by Abhi on 21-Apr-2010 for caseid 1320 NOFOLDER for in penline
    If isNOFOLDERExists(FileName) = False Then
        Set tsObj = fsObj.OpenTextFile(FileName, ForReading)
        'by Abhi on 04-Jan-2010 for caseid 1582 first free numbers sequence number genaration
        'UploadNo = mSeqNumberGen("UPLOADNO")
        'by Abhi on 13-Feb-2012 for caseid 1582 first free numbers as stored procedure in pengds
        'UploadNo = PENFirstFreeNumber("UPLOADNO")
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        dbCompany.BeginTrans
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        PENErr_BeginTrans = True
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        UploadNo = FirstFreeNumber("UPLOADNO")
        'by Abhi on 28-Apr-2010 for caseid 1205 PENFARE old format error in Worldspan in PenGDS will not stuck
        PENFAREPNO_Long = 1
        'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
        PENAIRTKTPNO_Long = 1
        'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
        RECIDQHeader = False
        Do While tsObj.AtEndOfStream <> True
            DoEvents
            LineStr = tsObj.ReadLine
            'by Abhi on 09-Oct-2010 for caseid 1519 PenGDS special character is ASCII code 0 in Worldspan ticketno
            LineStr = Replace(LineStr, Chr(0), "")
            Lines = Split(LineStr, Chr(13))
            If HeaderChecked = False Then
                If InStr(1, Left(LineStr, 20), "QU ", vbTextCompare) = 0 Then Exit Do
                HeaderChecked = True
            End If
            For j = 0 To UBound(Lines)
                crrItem = Lines(j)
                WorldSpan.PostLine crrItem, UploadNo, fTitle
                DoEvents
            Next
            DoEvents
        Loop
        tsObj.Close
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        FMain.stbUpload.Panels(1).Text = "Writing GDS In Tray..."
        DoEvents
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'vSQL_String = "" _
        '     & "INSERT INTO dbo.GDSINTRAYTABLE " _
        '     & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC]) " _
        '     & "SELECT     'PRT' AS GIT_ID, 'WORLDSPAN' AS GIT_GDS, 'WORLDSPAN' AS GIT_INTERFACE, dbo.WSPPNRADD.UpLoadNo AS GIT_UPLOADNO, " _
        '     & "                      dbo.WSPPNRADD.DLCI AS GIT_PNRDATE, dbo.WSPPNRADD.PNRADD AS GIT_PNR, dbo.WSPPNAME.SURNAME AS GIT_LASTNAME, " _
        '     & "                      dbo.WSPPNAME.FRSTNAME AS GIT_FIRSTNAME, dbo.WSPAGNTSINE.BAGENT AS GIT_BOOKINGAGENT, dbo.WSPPNAME.DOCNO AS GIT_TICKETNUMBER, " _
        '     & "                      dbo.WSPPNRADD.FNAME AS GIT_FILENAME, dbo.WSPPNRADD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.WSPAGNTSINE.BSID AS GIT_PCC " _
        '     & "FROM         dbo.WSPPNAME INNER JOIN " _
        '     & "                      dbo.WSPPNRADD ON dbo.WSPPNAME.UpLoadNo = dbo.WSPPNRADD.UpLoadNo INNER JOIN " _
        '     & "                      dbo.WSPAGNTSINE ON dbo.WSPPNRADD.UpLoadNo = dbo.WSPAGNTSINE.UpLoadNo " _
        '     & "WHERE     (dbo.WSPPNRADD.UpLoadNo = " & UploadNo & ") AND (dbo.WSPPNRADD.RecID <> '')"
        vDateLastModified_String = fsObj.GetFile(FileName).DateLastModified
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
        vDateLastModified_String = DateTime12hrsFormat(vDateLastModified_String)
        vFileDateCreated_String = fsObj.GetFile(FileName).DateCreated
        vFileDateCreated_String = DateTime12hrsFormat(vFileDateCreated_String)
        'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
        'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
        vSQL_String = "" _
            & "SELECT     TOP (1) BRANCH " _
            & "FROM         dbo.WSPPENLINE " _
            & "WHERE     (UpLoadNo = " & UploadNo & ") AND (RecID = 'PEN') AND (BRANCH <> '')"
        GIT_PENLINEBRID_String = getFromExecuted(vSQL_String, "BRANCH")
        'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
        
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        'by Abhi on 31-Aug-2017 for caseid 7742 Worldspan EMD Ticktets from Record -E – For an Issued EMD
        'ErrDetails_String = " GDSINTRAYTABLE."
        ErrDetails_String = " GDSINTRAYTABLE 1."
        'by Abhi on 31-Aug-2017 for caseid 7742 Worldspan EMD Ticktets from Record -E – For an Issued EMD
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        'by Abhi on 31-Aug-2017 for caseid 7742 Worldspan EMD Ticktets from Record -E – For an Issued EMD
        'vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC],[GIT_FILEISSUEDATE],[GIT_CUSTOMERUSERCODE],[GIT_PENWAIT],[GIT_FILECREATEDDATE]) " _
            & "SELECT     'PRT' AS GIT_ID, 'WORLDSPAN' AS GIT_GDS, 'WORLDSPAN' AS GIT_INTERFACE, dbo.WSPPNRADD.UpLoadNo AS GIT_UPLOADNO, " _
            & "                      dbo.WSPPNRADD.DLCI AS GIT_PNRDATE, dbo.WSPPNRADD.PNRADD AS GIT_PNR, dbo.WSPPNAME.SURNAME AS GIT_LASTNAME, " _
            & "                      dbo.WSPPNAME.FRSTNAME AS GIT_FIRSTNAME, dbo.WSPAGNTSINE.BAGENT AS GIT_BOOKINGAGENT, dbo.WSPPNAME.DOCNO AS GIT_TICKETNUMBER, " _
            & "                      dbo.WSPPNRADD.FNAME AS GIT_FILENAME, dbo.WSPPNRADD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.WSPAGNTSINE.BSID AS GIT_PCC, '" & vDateLastModified_String & "' AS GIT_FILEISSUEDATE, " _
            & "                      '" & SkipChars(GIT_CUSTOMERUSERCODE_String) & "' AS GIT_CUSTOMERUSERCODE, '" & SkipChars(GIT_PENWAIT_String) & "' AS GIT_PENWAIT, " _
            & "                      '" & vFileDateCreated_String & "' AS GIT_FILECREATEDDATE " _
            & "FROM         dbo.WSPPNAME INNER JOIN " _
            & "                      dbo.WSPPNRADD ON dbo.WSPPNAME.UpLoadNo = dbo.WSPPNRADD.UpLoadNo INNER JOIN " _
            & "                      dbo.WSPAGNTSINE ON dbo.WSPPNRADD.UpLoadNo = dbo.WSPAGNTSINE.UpLoadNo " _
            & "WHERE     (dbo.WSPPNRADD.UpLoadNo = " & UploadNo & ") AND (dbo.WSPPNRADD.RecID <> '')"
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC],[GIT_FILEISSUEDATE],[GIT_CUSTOMERUSERCODE],[GIT_PENWAIT],[GIT_FILECREATEDDATE],[GIT_ISSUEDBY], [GIT_PENLINEBRID]) " _
            & "SELECT     'PRT' AS GIT_ID, 'WORLDSPAN' AS GIT_GDS, 'WORLDSPAN' AS GIT_INTERFACE, dbo.WSPPNRADD.UpLoadNo AS GIT_UPLOADNO, " _
            & "                      dbo.WSPPNRADD.DLCI AS GIT_PNRDATE, dbo.WSPPNRADD.PNRADD AS GIT_PNR, dbo.WSPPNAME.SURNAME AS GIT_LASTNAME, " _
            & "                      dbo.WSPPNAME.FRSTNAME AS GIT_FIRSTNAME, dbo.WSPAGNTSINE.BAGENT AS GIT_BOOKINGAGENT, dbo.WSPPNAME.DOCNO AS GIT_TICKETNUMBER, " _
            & "                      dbo.WSPPNRADD.FNAME AS GIT_FILENAME, dbo.WSPPNRADD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.WSPAGNTSINE.BSID AS GIT_PCC, '" & vDateLastModified_String & "' AS GIT_FILEISSUEDATE, " _
            & "                      '" & SkipChars(GIT_CUSTOMERUSERCODE_String) & "' AS GIT_CUSTOMERUSERCODE, '" & SkipChars(GIT_PENWAIT_String) & "' AS GIT_PENWAIT, " _
            & "                      '" & vFileDateCreated_String & "' AS GIT_FILECREATEDDATE, dbo.WSPAGNTSINE.TAGENT AS GIT_ISSUEDBY, '" & GIT_PENLINEBRID_String & "' AS GIT_PENLINEBRID " _
            & "FROM         dbo.WSPPNAME INNER JOIN " _
            & "                      dbo.WSPPNRADD ON dbo.WSPPNAME.UpLoadNo = dbo.WSPPNRADD.UpLoadNo INNER JOIN " _
            & "                      dbo.WSPAGNTSINE ON dbo.WSPPNRADD.UpLoadNo = dbo.WSPAGNTSINE.UpLoadNo " _
            & "WHERE     (dbo.WSPPNRADD.UpLoadNo = " & UploadNo & ") AND (dbo.WSPPNRADD.RecID <> '') AND (dbo.WSPPNAME.DOCNO <> 'J')"
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        'by Abhi on 31-Aug-2017 for caseid 7742 Worldspan EMD Ticktets from Record -E – For an Issued EMD
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        dbCompany.Execute vSQL_String
        
        'by Abhi on 31-Aug-2017 for caseid 7742 Worldspan EMD Ticktets from Record -E – For an Issued EMD
        ErrDetails_String = " GDSINTRAYTABLE 2."
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC],[GIT_FILEISSUEDATE],[GIT_CUSTOMERUSERCODE],[GIT_PENWAIT],[GIT_FILECREATEDDATE],[GIT_ISSUEDBY], [GIT_PENLINEBRID]) " _
            & "SELECT     'PRT' AS GIT_ID, 'WORLDSPAN' AS GIT_GDS, 'WORLDSPAN' AS GIT_INTERFACE, dbo.WSPPNRADD.UpLoadNo AS GIT_UPLOADNO, " _
            & "                      dbo.WSPPNRADD.DLCI AS GIT_PNRDATE, dbo.WSPPNRADD.PNRADD AS GIT_PNR, dbo.WSPPNAME.SURNAME AS GIT_LASTNAME, " _
            & "                      dbo.WSPPNAME.FRSTNAME AS GIT_FIRSTNAME, dbo.WSPAGNTSINE.BAGENT AS GIT_BOOKINGAGENT, dbo.WSPIssuedEMD.EMDNumber AS GIT_TICKETNUMBER, " _
            & "                      dbo.WSPPNRADD.FNAME AS GIT_FILENAME, dbo.WSPPNRADD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.WSPAGNTSINE.BSID AS GIT_PCC, '" & vDateLastModified_String & "' AS GIT_FILEISSUEDATE, " _
            & "                      '" & SkipChars(GIT_CUSTOMERUSERCODE_String) & "' AS GIT_CUSTOMERUSERCODE, '" & SkipChars(GIT_PENWAIT_String) & "' AS GIT_PENWAIT, " _
            & "                      '" & vFileDateCreated_String & "' AS GIT_FILECREATEDDATE, dbo.WSPAGNTSINE.TAGENT AS GIT_ISSUEDBY, '" & GIT_PENLINEBRID_String & "' AS GIT_PENLINEBRID " _
            & "FROM         dbo.WSPPNAME INNER JOIN " _
            & "                      dbo.WSPPNRADD ON dbo.WSPPNAME.UpLoadNo = dbo.WSPPNRADD.UpLoadNo INNER JOIN " _
            & "                      dbo.WSPAGNTSINE ON dbo.WSPPNRADD.UpLoadNo = dbo.WSPAGNTSINE.UpLoadNo INNER JOIN " _
            & "                      dbo.WSPIssuedEMD ON dbo.WSPPNAME.UpLoadNo = dbo.WSPIssuedEMD.UpLoadNo AND " _
            & "                      dbo.WSPPNAME.SURNAMEFRSTNAMENO = dbo.WSPIssuedEMD.SurnameFirstnameNumber " _
            & "WHERE     (dbo.WSPPNRADD.UpLoadNo = " & UploadNo & ") AND (dbo.WSPPNRADD.RecID <> '')"
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        dbCompany.Execute vSQL_String
        'by Abhi on 31-Aug-2017 for caseid 7742 Worldspan EMD Ticktets from Record -E – For an Issued EMD
        
        FMain.stbUpload.Panels(1).Text = "File moving to Destination..."
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    End If
    
    FileObj.CopyFile txtFileName.Text, txtRelLocation.Text & "\", True
'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
'On Error Resume Next
'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
    FileObj.DeleteFile txtFileName.Text, True
    'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
    INIWrite App.Path & "\PenSEARCH.ini", "Worldspan", "LUFPNR", LUFPNR_String
    INIWrite App.Path & "\PenSEARCH.ini", "Worldspan", "LUFDate", DateFormat(Date)
    INIWrite App.Path & "\PenSEARCH.ini", "Worldspan", "LUFTime", TimeFormat12HRS(time) & "(" & TimeFormat(time) & ")"
    'by Abhi on 18-May-2011 for caseid 1757 Added Events for GDSAuto
    'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
    'Call EventLog("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & LUFPNR_String & " - " & FMain.stbUpload.Panels(3).Text)
    Call EventLog("#PenGDS# " & PadR(FMain.stbUpload.Panels(2).Text, 9) & " - " & PadR(LUFPNR_String, 6) & " - " & PadR(FMain.stbUpload.Panels(3).Text, 27))
    'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
    'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    'by Abhi on 06-Dec-2018 for caseid 9584 Sabre Error: -2147168242 - No transaction is active
    'dbCompany.CommitTrans
    If PENErr_BeginTrans = True Then
        dbCompany.CommitTrans
    End If
    'by Abhi on 06-Dec-2018 for caseid 9584 Sabre Error: -2147168242 - No transaction is active
    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
    PENErr_BeginTrans = False
    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
    'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
Note:
    Me.Hide
Exit Function
PENErr:
    'by Abhi on 14-Feb-2014 for caseid 3741 Pengds -Galileo file stuck Incorrect syntax near ARCY
    If (tsObj Is Nothing) = False Then
        tsObj.Close
    End If
    'by Abhi on 14-Feb-2014 for caseid 3741 Pengds -Galileo file stuck Incorrect syntax near ARCY
    'If PENErr(Err) = True Then GoTo DeadlockRETRY
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'If PENErr() = True Then GoTo DeadlockRETRY
    If PENErr() = True Then Resume DeadlockRETRY
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
End Function




