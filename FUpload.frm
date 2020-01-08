VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FUpload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Galileo Uploading..."
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   10
      Top             =   1200
      Width           =   1365
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
      TabIndex        =   8
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Files Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
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
      TabIndex        =   5
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdDirUpload 
      Caption         =   "DirUpload"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   5160
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar Pgr 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Upload"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label1 
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
      TabIndex        =   9
      Top             =   735
      Width           =   1440
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
      TabIndex        =   6
      Top             =   360
      Width           =   1440
   End
End
Attribute VB_Name = "FUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'               Copyright (C) 2005 Penguin Software Exports Pvt Ltd
'     Description :     Module For Uploading the data from Galileo Systems
'     Created By  :     Jino
'     Date        :     30-May-2005
'     Modified By :
'
'
'-------------------------------------------------------------------------------

Dim i As Integer
Dim UploadNo As Long
Dim LastSec As String
Dim NoteError As Boolean
Dim FNAME As String
Dim ITNO As Long
Dim LastSecNo_String As String

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
'by Abhi on 08-Apr-2009 for slno
Public fGalPassengerSLNO As Long
'by Abhi on 12-Feb-2014 for caseid 3726 A29 EMDVAR (EMD VARIABLE DATA)
Dim fA29PAX_String As String
'by Abhi on 12-Feb-2014 for caseid 3726 A29 EMDVAR (EMD VARIABLE DATA)
'by Abhi on 21-Aug-2018 for caseid 8911 Galileo-GDS Upload GalA10APO
Dim fMIRT50TRC_String As String
Dim fMIRType_String As String
Dim fMIRTypeID_String As String
'by Abhi on 21-Aug-2018 for caseid 8911 Galileo-GDS Upload GalA10APO

Private Sub chkGal_Click()
'If chkGal.Value = 0 Then
'    dbCompany.Execute "Update File set STATUS=0"
'Else
'    dbCompany.Execute "Update File set STATUS=1"
'End If
End Sub

Private Sub CmdDelete_Click()
'On Error GoTo Note
'    Dim Count As Integer
'    File1.Path = txtDirName.Text
'    File1.Pattern = "*.mir"
'
'    If txtDirName.Text = "" Or File1.ListCount = 0 Then MsgBox "MIR Files Not Found", vbInformation: Exit Sub
'    If MsgBox("Do you want to Delete the " & File1.ListCount & " MIR Files ?", vbYesNo) = vbYes Then
'        For Count = 0 To File1.ListCount - 1
'            txtFileName.Text = File1.Path & "\" & File1.List(Count)
'            DeleteFileToRecycleBin (txtFileName)
'        Next
'    End If
'Exit Sub
'Note:
'    MsgBox Err.Description & ". Invalid Drive Location", vbCritical: Exit Sub
End Sub
Public Sub cmdDirUpload_Click()
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'On Error GoTo Note
    Dim Count As Integer
    
    FMain.SendStatus FMain.SSTab1.TabCaption(1)
    File1.Path = txtDirName.Text
    File1.Pattern = FMain.txtGalileoExt
    If txtDirName.Text = "" Or File1.ListCount = 0 Then FMain.stbUpload.Panels(1).Text = "MIR Files Not Found...": Exit Sub
    'Uploading Each File
    Open (App.Path & "\_UploadingSQL_") For Random As #1
    Close #1
    FMain.cmdStop.Enabled = False
    DoEvents
    FMain.lblFName.Caption = "Uploading Galilieo..."
    FMain.stbUpload.Panels(1).Text = "Reading..."
    FMain.stbUpload.Panels(2).Text = "Galilieo"
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
        'If Trim(txtFileName.Text) = "" Then Exit Sub
        If Trim(txtFileName.Text) <> "" Then
        'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
            'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
            FMain.stbUpload.Panels(1).Text = "Reading..."
            FNAME = txtFileName.Text
            'FMain.SendStatus FNAME
            txtFileName.Text = File1.Path & "\" & txtFileName.Text
            cmdLoad_Click
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
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'Exit Sub
'Note:
'    Me.Hide
End Sub
Private Sub cmdLoad_Click()
'On Error GoTo Note
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
On Error GoTo PENErr
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    Dim FileSystem As FileSystemObject
    Dim CopyStatus As Boolean
    Dim LineItem As String
    Dim FileObj, LineDet
    'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    Dim vSQL_String As String
    'by Abhi on 28-Jan-2014 for caseid 3686 Galileo-PNR File with line feed is not showing intray
    Dim tsObj As TextStream
    'by Abhi on 28-Jan-2014 for caseid 3686 Galileo-PNR File with line feed is not showing intray
    'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    Dim vDateLastModified_String As String
    'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
    Dim vFileDateCreated_String As String
    'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
    
    If txtFileName = "" Then Exit Sub
    
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    GDSDeadlockRETRY_Integer = 0
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
DeadlockRETRY:
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    LUFPNR_String = ""
    NoteError = False
    CopyStatus = False
    
    Set FileObj = CreateObject("Scripting.FileSystemObject")
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    NoofPermissionDenied = 0
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    'by Abhi on 21-Apr-2010 for caseid 1320 NOFOLDER for in penline
    If isNOFOLDERExists(txtFileName.Text) = False Then
'by Abhi on 28-Jan-2014 for caseid 3686 Galileo-PNR File with line feed is not showing intray
'        Set LineDet = FileObj.OpenTextFile(txtFileName.Text, ForReading, TristateUseDefault)
'        'by Abhi on 08-Apr-2009 for slno
'        fGalPassengerSLNO = 1
'        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
'        dbCompany.BeginTrans
'        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
'        Do While LineDet.AtEndOfLine <> True
'            DoEvents
'            LineItem = LineDet.ReadLine
'            'If fnChecktheDatabase(LineItem) = True Then CopyStatus = True: Exit Do
'            SelectDetails (LineItem)
'            DoEvents
'        Loop
'        LineDet.Close
        Set tsObj = fsObj.OpenTextFile(txtFileName.Text, ForReading)
        'by Abhi on 08-Apr-2009 for slno
        'by Abhi on 28-Oct-2014 for caseid 4665 Additional checking in Passenger type picking in Gal File
        'fGalPassengerSLNO = 1
        fGalPassengerSLNO = 0
        'by Abhi on 28-Oct-2014 for caseid 4665 Additional checking in Passenger type picking in Gal File
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        dbCompany.BeginTrans
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        PENErr_BeginTrans = True
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        GIT_CUSTOMERUSERCODE_String = ""
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        GIT_PENWAIT_String = "N"
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        
        UploadNo = FirstFreeNumber("UPLOADNO")
        i = 0
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        Do While tsObj.AtEndOfStream <> True
            DoEvents
            LineItem = tsObj.ReadLine
            'If fnChecktheDatabase(LineItem) = True Then CopyStatus = True: Exit Do
            If Trim(LineItem) <> "" Then
                SelectDetails (LineItem)
            End If
            DoEvents
        Loop
        tsObj.Close
        'by Abhi on 08-Feb-2014 for caseid 3710 Galileo PNR Upload eror-Type mismatch due to ARNK segment
        vSQL_String = "" _
            & "UPDATE    GalA04 " _
            & "SET              A04CDE = NULL " _
            & "WHERE     (A04NME = N'ARNK') AND (UPLOADNO = " & UploadNo & ")"
        dbCompany.Execute vSQL_String
        'by Abhi on 08-Feb-2014 for caseid 3710 Galileo PNR Upload eror-Type mismatch due to ARNK segment
'by Abhi on 28-Jan-2014 for caseid 3686 Galileo-PNR File with line feed is not showing intray
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        FMain.stbUpload.Panels(1).Text = "Writing GDS In Tray..."
        DoEvents
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'vSQL_String = "" _
        '    & "INSERT INTO dbo.GDSINTRAYTABLE " _
        '    & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC]) " _
        '    & "SELECT     'MIR' AS GIT_ID, 'GALILEO' AS GIT_GDS, 'GALILEO' AS GIT_INTERFACE, dbo.GalHeader2.UPLOADNO AS GIT_UPLOADNO, dbo.GalHeader2.T50DTL AS GIT_PNRDATE, " _
        '    & "                      dbo.GalHeader2.T50RCL AS GIT_PNR, dbo.GalPassenger.A02NMC AS GIT_LASTNAME, '' AS GIT_FIRSTNAME, RIGHT(dbo.GalHeader2.T50AGS, 2) " _
        '    & "                      AS GIT_BOOKINGAGENT, CAST(dbo.GalPassenger.A02TKT AS nvarchar) AS GIT_TICKETNUMBER, dbo.GalHeader1.FNAME AS GIT_FILENAME, " _
        '    & "                      dbo.GalHeader1.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.GalHeader2.T50BPC AS GIT_PCC " _
        '    & "FROM         dbo.GalHeader1 RIGHT OUTER JOIN " _
        '    & "                      dbo.GalPassenger ON dbo.GalHeader1.UPLOADNO = dbo.GalPassenger.UPLOADNO LEFT OUTER JOIN " _
        '    & "                      dbo.GalHeader2 ON dbo.GalPassenger.UPLOADNO = dbo.GalHeader2.UPLOADNO " _
        '    & "Where (dbo.GalHeader2.UploadNo = " & UploadNo & ")"
        'dbCompany.Execute vSQL_String
        '
        'vSQL_String = "" _
        '    & "INSERT INTO dbo.GDSINTRAYTABLE " _
        '    & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC]) " _
        '    & "SELECT     'MIR' AS GIT_ID, 'GALILEO' AS GIT_GDS, 'GALILEO' AS GIT_INTERFACE, dbo.GalHeader2.UPLOADNO AS GIT_UPLOADNO, dbo.GalHeader2.T50DTL AS GIT_PNRDATE, " _
        '    & "                      dbo.GalHeader2.T50RCL AS GIT_PNR, dbo.GalA29.A29NAM AS GIT_LASTNAME, '' AS GIT_FIRSTNAME, RIGHT(dbo.GalHeader2.T50AGS, 2) " _
        '    & "                      AS GIT_BOOKINGAGENT, CAST(RIGHT(dbo.GalA29.A29EMD, 10) AS nvarchar) AS GIT_TICKETNUMBER, dbo.GalHeader1.FNAME AS GIT_FILENAME, " _
        '    & "                      dbo.GalHeader1.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.GalHeader2.T50BPC AS GIT_PCC " _
        '    & "FROM         dbo.GalHeader1 RIGHT OUTER JOIN " _
        '    & "                      dbo.GalA29 ON dbo.GalHeader1.UPLOADNO = dbo.GalA29.UPLOADNO LEFT OUTER JOIN " _
        '    & "                      dbo.GalHeader2 ON dbo.GalA29.UPLOADNO = dbo.GalHeader2.UPLOADNO " _
        '    & "Where (dbo.GalHeader2.UploadNo = " & UploadNo & ")"
        vDateLastModified_String = fsObj.GetFile(txtFileName.Text).DateLastModified
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
        vDateLastModified_String = DateTime12hrsFormat(vDateLastModified_String)
        vFileDateCreated_String = fsObj.GetFile(txtFileName.Text).DateCreated
        vFileDateCreated_String = DateTime12hrsFormat(vFileDateCreated_String)
        'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
        'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
        vSQL_String = "" _
            & "SELECT     TOP (1) A14BRANCH " _
            & "FROM         GalA14 " _
            & "WHERE     (UPLOADNO = " & UploadNo & ") AND (A14SEC = N'A14') AND (A14BRANCH <> '')"
        GIT_PENLINEBRID_String = getFromExecuted(vSQL_String, "A14BRANCH")
        'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        ErrDetails_String = " GDSINTRAYTABLE 1."
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC],[GIT_FILEISSUEDATE],[GIT_CUSTOMERUSERCODE],[GIT_PENWAIT],[GIT_FILECREATEDDATE],[GIT_ISSUEDBY], [GIT_PENLINEBRID]) " _
            & "SELECT     'MIR' AS GIT_ID, 'GALILEO' AS GIT_GDS, 'GALILEO' AS GIT_INTERFACE, dbo.GalHeader2.UPLOADNO AS GIT_UPLOADNO, dbo.GalHeader2.T50DTL AS GIT_PNRDATE, " _
            & "                      dbo.GalHeader2.T50RCL AS GIT_PNR, dbo.GalPassenger.A02NMC AS GIT_LASTNAME, '' AS GIT_FIRSTNAME, RIGHT(dbo.GalHeader2.T50AGS, 2) " _
            & "                      AS GIT_BOOKINGAGENT, CAST(dbo.GalPassenger.A02TKT AS nvarchar) AS GIT_TICKETNUMBER, dbo.GalHeader1.FNAME AS GIT_FILENAME, " _
            & "                      dbo.GalHeader1.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.GalHeader2.T50BPC AS GIT_PCC, '" & vDateLastModified_String & "' AS GIT_FILEISSUEDATE, " _
            & "                      '" & SkipChars(GIT_CUSTOMERUSERCODE_String) & "' AS GIT_CUSTOMERUSERCODE, '" & SkipChars(GIT_PENWAIT_String) & "' AS GIT_PENWAIT, " _
            & "                      '" & vFileDateCreated_String & "' AS GIT_FILECREATEDDATE, dbo.GalHeader2.T50SIN AS GIT_ISSUEDBY, '" & GIT_PENLINEBRID_String & "' AS GIT_PENLINEBRID " _
            & "FROM         dbo.GalHeader1 RIGHT OUTER JOIN " _
            & "                      dbo.GalPassenger ON dbo.GalHeader1.UPLOADNO = dbo.GalPassenger.UPLOADNO LEFT OUTER JOIN " _
            & "                      dbo.GalHeader2 ON dbo.GalPassenger.UPLOADNO = dbo.GalHeader2.UPLOADNO " _
            & "Where (dbo.GalHeader2.UploadNo = " & UploadNo & ")"
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        dbCompany.Execute vSQL_String
        
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        ErrDetails_String = " GDSINTRAYTABLE 2."
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC],[GIT_FILEISSUEDATE],[GIT_CUSTOMERUSERCODE],[GIT_PENWAIT],[GIT_FILECREATEDDATE],[GIT_ISSUEDBY], [GIT_PENLINEBRID]) " _
            & "SELECT     'MIR' AS GIT_ID, 'GALILEO' AS GIT_GDS, 'GALILEO' AS GIT_INTERFACE, dbo.GalHeader2.UPLOADNO AS GIT_UPLOADNO, dbo.GalHeader2.T50DTL AS GIT_PNRDATE, " _
            & "                      dbo.GalHeader2.T50RCL AS GIT_PNR, dbo.GalA29.A29NAM AS GIT_LASTNAME, '' AS GIT_FIRSTNAME, RIGHT(dbo.GalHeader2.T50AGS, 2) " _
            & "                      AS GIT_BOOKINGAGENT, CAST(RIGHT(dbo.GalA29.A29EMD, 10) AS nvarchar) AS GIT_TICKETNUMBER, dbo.GalHeader1.FNAME AS GIT_FILENAME, " _
            & "                      dbo.GalHeader1.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.GalHeader2.T50BPC AS GIT_PCC, '" & vDateLastModified_String & "' AS GIT_FILEISSUEDATE, " _
            & "                      '" & SkipChars(GIT_CUSTOMERUSERCODE_String) & "' AS GIT_CUSTOMERUSERCODE, '" & SkipChars(GIT_PENWAIT_String) & "' AS GIT_PENWAIT, " _
            & "                      '" & vFileDateCreated_String & "' AS GIT_FILECREATEDDATE, dbo.GalHeader2.T50SIN AS GIT_ISSUEDBY, '" & GIT_PENLINEBRID_String & "' AS GIT_PENLINEBRID " _
            & "FROM         dbo.GalHeader1 RIGHT OUTER JOIN " _
            & "                      dbo.GalA29 ON dbo.GalHeader1.UPLOADNO = dbo.GalA29.UPLOADNO LEFT OUTER JOIN " _
            & "                      dbo.GalHeader2 ON dbo.GalA29.UPLOADNO = dbo.GalHeader2.UPLOADNO " _
            & "Where (dbo.GalHeader2.UploadNo = " & UploadNo & ")"
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        dbCompany.Execute vSQL_String
        'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
        dbCompany.CommitTrans
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        PENErr_BeginTrans = False
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
        FMain.stbUpload.Panels(1).Text = "File moving to Destination..."
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    End If
    
    If CopyStatus = False And NoteError = False Then
        'FileObj.MoveFile txtFileName.Text, txtRelLocation.Text & "\"
        FileObj.CopyFile txtFileName.Text, txtRelLocation.Text & "\", True
'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
'On Error Resume Next
'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
        FileObj.DeleteFile txtFileName.Text, True
        'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
        INIWrite App.Path & "\PenSEARCH.ini", "Galileo", "LUFPNR", LUFPNR_String
        INIWrite App.Path & "\PenSEARCH.ini", "Galileo", "LUFDate", DateFormat(Date)
        INIWrite App.Path & "\PenSEARCH.ini", "Galileo", "LUFTime", TimeFormat12HRS(time) & "(" & TimeFormat(time) & ")"
        'by Abhi on 18-May-2011 for caseid 1757 Added Events for GDSAuto
        'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
        'Call EventLog("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & LUFPNR_String & " - " & FMain.stbUpload.Panels(3).Text)
        Call EventLog("#PenGDS# " & PadR(FMain.stbUpload.Panels(2).Text, 9) & " - " & PadR(LUFPNR_String, 6) & " - " & PadR(FMain.stbUpload.Panels(3).Text, 27))
        'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
        'dbCompany.CommitTrans
        'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    End If
    Exit Sub
Note:
    Me.Hide
Exit Sub
PENErr:
    'by Abhi on 14-Feb-2014 for caseid 3741 Pengds -Galileo file stuck Incorrect syntax near ARCY
    If (tsObj Is Nothing) = False Then
        tsObj.Close
    End If
    'by Abhi on 14-Feb-2014 for caseid 3741 Pengds -Galileo file stuck Incorrect syntax near ARCY
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'If PENErr() = True Then GoTo DeadlockRETRY
    If PENErr() = True Then Resume DeadlockRETRY
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
End Sub

'by Abhi on 28-Jan-2014 for caseid 3686 Galileo-PNR File with line feed is not showing intray
'Public Sub SelectDetails(LineStr As String)
'Dim Item As String
'i = 1
'    'by Abhi on 04-Jan-2010 for caseid 1582 first free numbers sequence number genaration
'    'UploadNo = mSeqNumberGen("UPLOADNO")
'    'by Abhi on 13-Feb-2012 for caseid 1582 first free numbers as stored procedure in pengds
'    'UploadNo = PENFirstFreeNumber("UPLOADNO")
'    UploadNo = FirstFreeNumber("UPLOADNO")
'    '''CnCompany.BeginTrans
'Step1:
'    If InStr(1, LineStr, Chr(13)) > 0 Then
'        Item = Mid(LineStr, 1, InStr(1, LineStr, Chr(13)))
'        LineStr = Mid(LineStr, InStr(1, LineStr, Chr(13)) + 1, Len(LineStr))
'        fnLineDivision (Item)
'        i = i + 1
'        If NoteError = True Then Exit Sub '''CnCompany.RollbackTrans:
'        GoTo Step1
'    End If
'   ''''''CnCompany.CommitTrans
'    Pgr.value = 100
'End Sub

Public Sub SelectDetails(LineStr As String)
Dim Item As String
If InStr(1, LineStr, Chr(13)) > 0 Then
    i = 1
End If
    If InStr(1, LineStr, Chr(13)) > 0 Then
Step1:
        If InStr(1, LineStr, Chr(13)) > 0 Then
            Item = Mid(LineStr, 1, InStr(1, LineStr, Chr(13)))
            LineStr = Mid(LineStr, InStr(1, LineStr, Chr(13)) + 1, Len(LineStr))
            fnLineDivision (Item)
            i = i + 1
            If NoteError = True Then Exit Sub '''CnCompany.RollbackTrans:
            GoTo Step1
        End If
    Else
        i = i + 1
        fnLineDivision (LineStr)
        Exit Sub
    End If
   ''''''CnCompany.CommitTrans
    Pgr.value = 100
End Sub
'by Abhi on 28-Jan-2014 for caseid 3686 Galileo-PNR File with line feed is not showing intray


Public Sub fnLineDivision(LStr As String)
    Dim Coll As New Collection
    'by Abhi on 03-Dec-2009 for Galileo Hotel HHL PNR uploading
    Dim vOPTDATAID_String As String
    Dim vSQL_String As String
    'by Abhi on 25-Feb-2010 for caseid 1227 Booking no or staff id is showing blank gds intray
    Dim vT50DFT_String As String
    Dim vT50DTL_String As String
    Dim vT50PNR_String As String
    'by Abhi on 11-Jan-2010 for caseid 1588 Galileo Penline AC BB
    Dim CollPenSplited As New Collection
    'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
    Dim TempPENPNO As Long
    Dim aPEN
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Dim aPENSplited
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
    'by Abhi on 11-Feb-2014 for caseid 3721 Galileo EMD documents A29 and A30 section
    Dim Splited
    'by Abhi on 11-Feb-2014 for caseid 3721 Galileo EMD documents A29 and A30 section
    'by Abhi on 26-May-2014 for caseid 4088 Galileo Reissue Ticket
    Dim vA11Last_String As String
    Dim vA11LastStart_Long As Long
    Dim vA11Lastspecific_String As String
    Dim vA11PGRI_String As String
    Dim vA11PGR_String As String
    'by Abhi on 26-May-2014 for caseid 4088 Galileo Reissue Ticket
    'by Abhi on 28-Oct-2014 for caseid 4665 Additional checking in Passenger type picking in Gal File
    Dim vPos_Long As Long
    Dim vA02PNRPTC_String As String
    'by Abhi on 28-Oct-2014 for caseid 4665 Additional checking in Passenger type picking in Gal File
    'by Abhi on 10-Dec-2015 for caseid 5829 Operating Airline for Galileo
    Dim vANL_String As String
    Dim vACL_String As String
    'by Abhi on 10-Dec-2015 for caseid 5829 Operating Airline for Galileo
    'by Abhi on 10-Dec-2016 for caseid 6106 Galileo-we need to add one more field in A04Gal table for DDL and pick the departure date from this new field
    Dim vDDL_String As String
    'by Abhi on 10-Dec-2016 for caseid 6106 Galileo-we need to add one more field in A04Gal table for DDL and pick the departure date from this new field
    'by Abhi on 17-Jan-2018 for caseid 8156 Galileo Ticket date picking
    Dim vA02NTD_String As String
    'by Abhi on 17-Jan-2018 for caseid 8156 Galileo Ticket date picking
    Dim vLStrPEN_String As String
    
    If Asc(LStr) = 13 Then Exit Sub
    
    
    'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
    'On Error GoTo myErr
    
            'T5 Table
            'by Abhi on 07-Oct-2015 for caseid 5630 Galileo File Error- The conversion of char data type to smalldatetime
            'If fnMidValue(LStr, 1, 2) = "T5" Then
            If fnMidValue(LStr, 1, 2) = "T5" And (Trim(fnMidValue(LStr, 3, 2)) = "1V" Or Trim(fnMidValue(LStr, 3, 2)) = "1G") Then
            'by Abhi on 07-Oct-2015 for caseid 5630 Galileo File Error- The conversion of char data type to smalldatetime
                'Header1
                LastSec = "T5"
                'by Abhi on 25-Feb-2010 for caseid 1227 Booking no or staff id is showing blank gds intray
                vT50DFT_String = Format(fnMidDate(LStr, 62, 7), "dd/mmm/yyyy")
                If IsDate(vT50DFT_String) = False Then
                    vT50DFT_String = Format(Date, "dd/mmm/yyyy")
                End If
                
                'by Abhi on 21-Aug-2018 for caseid 8911 Galileo-GDS Upload GalA10APO
                fMIRT50TRC_String = fnMidValue(LStr, 5, 4)
                Select Case fMIRT50TRC_String
                    Case "7733"
                        fMIRType_String = "GALILEO"
                        fMIRTypeID_String = "GCS"
                    Case "5880"
                        fMIRType_String = "APOLLO"
                        fMIRTypeID_String = "APO"
                End Select
                'by Abhi on 21-Aug-2018 for caseid 8911 Galileo-GDS Upload GalA10APO
                
                'by Abhi on 25-Feb-2010 for caseid 1227 Booking no or staff id is showing blank gds intray
                'by Abhi on 17-Sep-2019 for caseid 10708 Galileo -Airline numeric code storing issue
                'SQL = "INSERT INTO GalHeader1 (UPLOADNO,T50BID,T50TRC,T50TRC1,T50MIR,T50SIZE,T50SEQ,T50DTE,T50TME," & _
                    "T50ISC,T50ISA,T50ISN,T50DFT,T50INP,T50OUT,FNAME,GDSAutoFailed) values (" & Val(UploadNo) & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 2)) & "','" & Trim(fnMidValue(LStr, 3, 2)) & "'," & _
                    "" & Val(fnMidValue(LStr, 5, 4)) & "," & Val(fnMidValue(LStr, 9, 2)) & "," & _
                    "" & Val(fnMidValue(LStr, 11, 5)) & "," & Val(fnMidValue(LStr, 16, 5)) & "," & _
                    "'" & Format(fnMidDate(LStr, 21, 7), "dd/mmm/yyyy") & "','" & Trim(fnMidValue(LStr, 28, 5)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 33, 2)) & "'," & Val(fnMidValue(LStr, 35, 3)) & "," & _
                    "'" & Trim(fnMidValue(LStr, 38, 24)) & "','" & vT50DFT_String & "'," & _
                    "'" & Trim(fnMidValue(LStr, 69, 6)) & "','" & Trim(fnMidValue(LStr, 75, 6)) & "','" & FNAME & "',0)"
                SQL = "INSERT INTO GalHeader1 (UPLOADNO,T50BID,T50TRC,T50TRC1,T50MIR,T50SIZE,T50SEQ,T50DTE,T50TME," & _
                    "T50ISC,T50ISA,T50ISN,T50DFT,T50INP,T50OUT,FNAME,GDSAutoFailed) values (" & Val(UploadNo) & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 2)) & "','" & Trim(fnMidValue(LStr, 3, 2)) & "'," & _
                    "" & Val(fnMidValue(LStr, 5, 4)) & "," & Val(fnMidValue(LStr, 9, 2)) & "," & _
                    "" & Val(fnMidValue(LStr, 11, 5)) & "," & Val(fnMidValue(LStr, 16, 5)) & "," & _
                    "'" & Format(fnMidDate(LStr, 21, 7), "dd/mmm/yyyy") & "','" & Trim(fnMidValue(LStr, 28, 5)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 33, 2)) & "','" & fnMidValue(LStr, 35, 3) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 38, 24)) & "','" & vT50DFT_String & "'," & _
                    "'" & Trim(fnMidValue(LStr, 69, 6)) & "','" & Trim(fnMidValue(LStr, 75, 6)) & "','" & FNAME & "',0)"
                'by Abhi on 17-Sep-2019 for caseid 10708 Galileo -Airline numeric code storing issue
                dbCompany.Execute SQL
            
            'A02 Passenger Table
            ElseIf fnMidValue(LStr, 1, 3) = "A02" Then
                LastSec = "A02"
                'by Abhi on 28-Oct-2014 for caseid 4665 Additional checking in Passenger type picking in Gal File
                fGalPassengerSLNO = fGalPassengerSLNO + 1
                'by Abhi on 28-Oct-2014 for caseid 4665 Additional checking in Passenger type picking in Gal File
                'by Abhi on 12-Feb-2014 for caseid 3727 Galileo multiple filed fares indicator-A02FFN
                'SQL = "INSERT INTO GalPassenger (UPLOADNO,A02SEC,A02NMC,A02TRN,A02YIN,A02TKT,A02NBK," & _
                    "A02INV,A02PIC,A02FIN,A02EIN,A02SLNO) values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 4, 33)) & "'," & Val(fnMidValue(LStr, 37, 11)) & "," & _
                    "" & Val(fnMidValue(LStr, 48, 1)) & "," & Val(fnMidValue(LStr, 49, 10)) & "," & _
                    "" & Val(fnMidValue(LStr, 59, 2)) & ",'" & Trim(fnMidValue(LStr, 61, 9)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 70, 6)) & "'," & Val(fnMidValue(LStr, 76, 2)) & "," & _
                    "" & Val(fnMidValue(LStr, 78, 2)) & "," & fGalPassengerSLNO & ")"
                SQL = "INSERT INTO GalPassenger (UPLOADNO,A02SEC,A02NMC,A02TRN,A02YIN,A02TKT,A02NBK," & _
                    "A02INV,A02PIC,A02FIN,A02EIN,A02SLNO,A02FFN) values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 4, 33)) & "'," & Val(fnMidValue(LStr, 37, 11)) & "," & _
                    "" & Val(fnMidValue(LStr, 48, 1)) & "," & Val(fnMidValue(LStr, 49, 10)) & "," & _
                    "" & Val(fnMidValue(LStr, 59, 2)) & ",'" & Trim(fnMidValue(LStr, 61, 9)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 70, 6)) & "'," & Val(fnMidValue(LStr, 76, 2)) & "," & _
                    "" & Val(fnMidValue(LStr, 78, 2)) & "," & fGalPassengerSLNO & ",'" & fnMidValue(LStr, 80, 1) & "')"
                'by Abhi on 12-Feb-2014 for caseid 3727 Galileo multiple filed fares indicator-A02FFN
                dbCompany.Execute SQL
                'by Abhi on 08-Apr-2009 for slno
                'by Abhi on 28-Oct-2014 for caseid 4665 Additional checking in Passenger type picking in Gal File
                'fGalPassengerSLNO = fGalPassengerSLNO + 1
                'by Abhi on 28-Oct-2014 for caseid 4665 Additional checking in Passenger type picking in Gal File
            'A04 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A04" Then
                LastSec = "A04"
'                sql = "INSERT INTO GalA04 (UPLOADNO,A04SEC,A04ITN,A04CDE,A04NUM,A04NME,A04FLT,A04CLS," & _
'                    "A04STS,A04DTE,A04TME,A04ARV,A04IND,A04OCC,A04OCN,A04DCC,A04DCN,A04DOM,A04SET," & _
'                    "A04SVC,A04STP,A04STO,A04BAG,A04AIR,A04DTR,A04MIL,A04FCI,A04SIC,A04COG,A04CGC," & _
'                    "A04CGN,A04CGD,A04CGT,A04GCR,A04GRR,A04AFC,A04ACC,A04FFI,A04FFM,A04TK1,A04TKT," & _
'                    "A04JTI,A04JTM) values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
'                    "" & Val(fnMidValue(LStr, 4, 2)) & ",'" & Trim(fnMidValue(LStr, 6, 2)) & "'," & _
'                    "" & Val(fnMidValue(LStr, 8, 3)) & ",'" & Trim(fnMidValue(LStr, 11, 12)) & "'," & _
'                    "'" & Trim(fnMidValue(LStr, 23, 4)) & "','" & Trim(fnMidValue(LStr, 27, 2)) & "'," & _
'                    "'" & Trim(fnMidValue(LStr, 29, 2)) & "','" & Trim(fnMidValue(LStr, 31, 5)) & "'," & _
'                    "'" & Trim(fnMidValue(LStr, 36, 5)) & "','" & Trim(fnMidValue(LStr, 41, 5)) & "'," & _
'                    "" & Val(fnMidValue(LStr, 46, 1)) & ",'" & Trim(fnMidValue(LStr, 47, 3)) & "'," & _
'                    "'" & Trim(fnMidValue(LStr, 50, 13)) & "','" & Trim(fnMidValue(LStr, 63, 3)) & "'," & _
'                    "'" & Trim(fnMidValue(LStr, 66, 13)) & "','" & Trim(fnMidValue(LStr, 79, 1)) & "'," & _
'                    "'" & Trim(fnMidValue(LStr, 80, 1)) & "','" & Trim(fnMidValue(LStr, 81, 4)) & "'," & _
'                    "'" & Trim(fnMidValue(LStr, 85, 1)) & "'," & Val(fnMidValue(LStr, 86, 1)) & "," & _
'                    "'" & Trim(fnMidValue(LStr, 87, 3)) & "','" & Trim(fnMidValue(LStr, 90, 4)) & "'," & _
'                    "'" & Trim(fnMidValue(LStr, 94, 3)) & "','" & Trim(fnMidValue(LStr, 97, 5)) & "'," & _
'                    "'" & Trim(fnMidValue(LStr, 102, 1)) & "'," & Val(fnMidValue(LStr, 103, 1)) & "," & _
'                    "'" & Trim(fnMidValue(LStr, 104, 4)) & "','" & Trim(fnMidValue(LStr, 108, 3)) & "'," & _
'                    "" & Val(fnMidValue(LStr, 111, 0)) & "," & Val(fnMidValue(LStr, 111, 1)) & "," & _
'                    "'" & Trim(fnMidValue(LStr, 112, 5)) & "','" & Trim(fnMidValue(LStr, 117, 4)) & "'," & _
'                    "'" & Trim(fnMidValue(LStr, 121, 6)) & "','" & Trim(fnMidValue(LStr, 127, 3)) & "'," & _
'                    "'" & Trim(fnMidValue(LStr, 129, 12)) & "','" & Trim(fnMidValue(LStr, 141, 3)) & "'," & Val(fnMidValue(LStr, 144, 5)) & "," & _
'                    "'" & Trim(fnMidValue(LStr, 149, 3)) & "','" & Trim(fnMidValue(LStr, 152, 1)) & "','" & Trim(fnMidValue(LStr, 153, 3)) & "'," & Val(fnMidValue(LStr, 156, 5)) & ")"
                
                'If InStr(1, LStr, "GCR:") > 0 And InStr(1, LStr, "AC:") > 0 And InStr(1, LStr, "FF:") > 0 And InStr(1, LStr, "TK:") > 0 Then
                
                'by Abhi on 10-Dec-2015 for caseid 5829 Operating Airline for Galileo
                'SQL = "INSERT INTO GalA04 (UPLOADNO,A04SEC,A04ITN,A04CDE,A04NUM,A04NME,A04FLT,A04CLS," & _
                    "A04STS,A04DTE,A04TME,A04ARV,A04IND,A04OCC,A04OCN,A04DCC,A04DCN,A04DOM,A04SET," & _
                    "A04SVC,A04STP,A04STO,A04BAG,A04AIR,A04DTR,A04MIL,A04FCI,A04SIC," & _
                    "A04TK1,A04TKT,A04JTI,A04JTM) " & _
                    "values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                    "" & Val(fnMidValue(LStr, 4, 2)) & ",'" & Trim(fnMidValue(LStr, 6, 2)) & "'," & _
                    "" & Val(fnMidValue(LStr, 8, 3)) & ",'" & Trim(fnMidValue(LStr, 11, 12)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 23, 4)) & "','" & Trim(fnMidValue(LStr, 27, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 29, 2)) & "','" & Trim(fnMidValue(LStr, 31, 5)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 36, 5)) & "','" & Trim(fnMidValue(LStr, 41, 5)) & "'," & _
                    "" & Val(fnMidValue(LStr, 46, 1)) & ",'" & Trim(fnMidValue(LStr, 47, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 50, 13)) & "','" & Trim(fnMidValue(LStr, 63, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 66, 13)) & "','" & Trim(fnMidValue(LStr, 79, 1)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 80, 1)) & "','" & Trim(fnMidValue(LStr, 81, 4)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 85, 1)) & "'," & Val(fnMidValue(LStr, 86, 1)) & "," & _
                    "'" & Trim(fnMidValue(LStr, 87, 3)) & "','" & Trim(fnMidValue(LStr, 90, 4)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 94, 3)) & "','" & Trim(fnMidValue(LStr, 97, 5)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 102, 1)) & "'," & Val(fnMidValue(LStr, 103, 1)) & "," & _
                    "'" & Trim(fnMidValue(LStr, 104, 3)) & "','" & Trim(fnMidValue(LStr, 107, 1)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 108, 3)) & "'," & Val(fnMidValue(LStr, 111, 5)) & ")"
                vPos_Long = InStr(1, LStr, "ANL:")
                If vPos_Long = 0 Then
                    vANL_String = ""
                Else
                    vANL_String = Trim(Mid(LStr, vPos_Long + 4, 24))
                End If
                
                vPos_Long = InStr(1, LStr, "ACL:")
                If vPos_Long = 0 Then
                    vACL_String = ""
                Else
                    vACL_String = Trim(Mid(LStr, vPos_Long + 4, 40))
                End If
                'by Abhi on 10-Dec-2016 for caseid 6106 Galileo-we need to add one more field in A04Gal table for DDL and pick the departure date from this new field
                vPos_Long = InStr(1, LStr, "DDL:")
                If vPos_Long = 0 Then
                    vDDL_String = ""
                Else
                    vDDL_String = Trim(Mid(LStr, vPos_Long + 4, 7))
                End If
                'by Abhi on 10-Dec-2016 for caseid 6106 Galileo-we need to add one more field in A04Gal table for DDL and pick the departure date from this new field
                'by Abhi on 10-Dec-2016 for caseid 6106 Galileo-we need to add one more field in A04Gal table for DDL and pick the departure date from this new field
                SQL = "INSERT INTO GalA04 (UPLOADNO,A04SEC,A04ITN,A04CDE,A04NUM,A04NME,A04FLT,A04CLS," & _
                    "A04STS,A04DTE,A04TME,A04ARV,A04IND,A04OCC,A04OCN,A04DCC,A04DCN,A04DOM,A04SET," & _
                    "A04SVC,A04STP,A04STO,A04BAG,A04AIR,A04DTR,A04MIL,A04FCI,A04SIC," & _
                    "A04TK1,A04TKT,A04JTI,A04JTM,A04ANL,A04NML,A04ACL,A04ACN,A04DDL,A04DTL) " & _
                    "values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                    "" & Val(fnMidValue(LStr, 4, 2)) & ",'" & Trim(fnMidValue(LStr, 6, 2)) & "'," & _
                    "" & Val(fnMidValue(LStr, 8, 3)) & ",'" & Trim(fnMidValue(LStr, 11, 12)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 23, 4)) & "','" & Trim(fnMidValue(LStr, 27, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 29, 2)) & "','" & Trim(fnMidValue(LStr, 31, 5)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 36, 5)) & "','" & Trim(fnMidValue(LStr, 41, 5)) & "'," & _
                    "" & Val(fnMidValue(LStr, 46, 1)) & ",'" & Trim(fnMidValue(LStr, 47, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 50, 13)) & "','" & Trim(fnMidValue(LStr, 63, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 66, 13)) & "','" & Trim(fnMidValue(LStr, 79, 1)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 80, 1)) & "','" & Trim(fnMidValue(LStr, 81, 4)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 85, 1)) & "'," & Val(fnMidValue(LStr, 86, 1)) & "," & _
                    "'" & Trim(fnMidValue(LStr, 87, 3)) & "','" & Trim(fnMidValue(LStr, 90, 4)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 94, 3)) & "','" & Trim(fnMidValue(LStr, 97, 5)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 102, 1)) & "'," & Val(fnMidValue(LStr, 103, 1)) & "," & _
                    "'" & Trim(fnMidValue(LStr, 104, 3)) & "','" & Trim(fnMidValue(LStr, 107, 1)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 108, 3)) & "'," & Val(fnMidValue(LStr, 111, 5)) & "," & _
                    "'" & "ANL:" & "','" & SkipChars(vANL_String) & "'," & _
                    "'" & "ACL:" & "','" & SkipChars(vACL_String) & "'," & _
                    "'" & "DDL:" & "','" & SkipChars(vDDL_String) & "'" & _
                    ")"
                'by Abhi on 10-Dec-2016 for caseid 6106 Galileo-we need to add one more field in A04Gal table for DDL and pick the departure date from this new field
                'by Abhi on 10-Dec-2015 for caseid 5829 Operating Airline for Galileo
                dbCompany.Execute SQL
            'caseid UNKNOWN by Abhi on 23-Jun-2008
            'A05 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A05" Then
                LastSec = "A05"
                SQL = "INSERT INTO GalA04 (UPLOADNO,A04SEC,A04ITN,A04CDE,A04NUM,A04NME,A04FLT,A04CLS," & _
                    "A04STS,A04DTE,A04TME,A04ARV,A04IND,A04OCC,A04OCN,A04DCC,A04DCN,A04SVC,A04STP) " & _
                    "values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                    "" & Val(fnMidValue(LStr, 4, 2)) & ",'" & Trim(fnMidValue(LStr, 6, 2)) & "'," & _
                    "" & Val(fnMidValue(LStr, 8, 3)) & ",'" & Trim(fnMidValue(LStr, 11, 12)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 23, 4)) & "','" & Trim(fnMidValue(LStr, 27, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 29, 2)) & "','" & Trim(fnMidValue(LStr, 31, 5)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 36, 5)) & "','" & Trim(fnMidValue(LStr, 41, 5)) & "'," & _
                    "" & Val(fnMidValue(LStr, 46, 1)) & ",'" & Trim(fnMidValue(LStr, 47, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 50, 13)) & "','" & Trim(fnMidValue(LStr, 63, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 66, 13)) & "','" & Trim(fnMidValue(LStr, 79, 4)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 83, 1)) & "')"
                dbCompany.Execute SQL
'                ITNO = Val(fnMidValue(LStr, 4, 2))
'                If InStr(1, LStr, "GCR:") > 0 Then
'                    sql = "UPDATE GalA07 Set A07ITT='" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
'                        "A07IT1='" & Trim(fnMidValue(LStr, 4, 8)) & "',A07IT1C='" & Trim(fnMidValue(LStr, 12, 2)) & "'," & _
'                        "A07IT2='" & Trim(fnMidValue(LStr, 14, 8)) & "',A07IT2C='" & Trim(fnMidValue(LStr, 22, 2)) & "'," & _
'                        "A07IT3='" & Trim(fnMidValue(LStr, 24, 8)) & "',A07IT3C='" & Trim(fnMidValue(LStr, 32, 2)) & "'," & _
'                        "A07IT4='" & Trim(fnMidValue(LStr, 34, 8)) & "',A07IT4C='" & Trim(fnMidValue(LStr, 42, 2)) & "'," & _
'                        "A07IT5='" & Trim(fnMidValue(LStr, 44, 8)) & "',A07IT5C='" & Trim(fnMidValue(LStr, 52, 2)) & "'," & _
'                        "A07IT6='" & Trim(fnMidValue(LStr, 54, 8)) & "',A07IT6C='" & Trim(fnMidValue(LStr, 62, 2)) & "'," & _
'                        "A07IT7='" & Trim(fnMidValue(LStr, 64, 8)) & "',A07IT7C='" & Trim(fnMidValue(LStr, 72, 2)) & "'," & _
'                        "A07IT8='" & Trim(fnMidValue(LStr, 74, 8)) & "',A07IT8C='" & Trim(fnMidValue(LStr, 82, 2)) & "'," & _
'                        "A07IT9='" & Trim(fnMidValue(LStr, 84, 8)) & "',A07IT9C='" & Trim(fnMidValue(LStr, 92, 2)) & "'," & _
'                        "A07IT10='" & Trim(fnMidValue(LStr, 94, 8)) & "',A07IT10C='" & Trim(fnMidValue(LStr, 102, 2)) & "'," & _
'                        "A07IT11='" & Trim(fnMidValue(LStr, 104, 8)) & "',A07IT11C='" & Trim(fnMidValue(LStr, 112, 2)) & "'," & _
'                        "A07IT12='" & Trim(fnMidValue(LStr, 114, 8)) & "',A07IT12C='" & Trim(fnMidValue(LStr, 122, 2)) & "' " & _
'                        "Where UPLOADNO=" & UploadNo & " and A04ITN=" & ITNO
'                    dbCompany.Execute sql
'                End If
            'A07 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A07" Then
                LastSec = "A07"
                'by Abhi on 06-Mar-2017 for caseid 7248 Special charcater in Currency code while saving PNR File
                SQL = "INSERT INTO GalA07 (UPLOADNO,A07SEC,A07FSI,A07CRB,A07TBF,A07CRT,A07TTA,A07CRE," & _
                    "A07EQV,A07CUR,A07TI1,A07TT1,A07TC1,A07TI2,A07TT2,A07TC2,A07TI3,A07TT3,A07TC3," & _
                    "A07TI4,A07TT4,A07TC4,A07TI5,A07TT5,A07TC5,A07XFI,A07XFT) values (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & Val(fnMidValue(LStr, 4, 2)) & "," & _
                    "'" & Trim(fnMidValue(LStr, 6, 3)) & "','" & Trim(fnMidValue(LStr, 9, 12)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 21, 3)) & "','" & Trim(fnMidValue(LStr, 24, 12)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 36, 3)) & "','" & Trim(fnMidValue(LStr, 39, 12)) & "'," & _
                    "'" & Trim(SkipCharsNonPrintable(fnMidValue(LStr, 51, 3))) & "','" & Trim(fnMidValue(LStr, 54, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 57, 8)) & "','" & Trim(fnMidValue(LStr, 65, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 67, 3)) & "','" & Trim(fnMidValue(LStr, 70, 8)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 78, 2)) & "','" & Trim(fnMidValue(LStr, 80, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 83, 8)) & "','" & Trim(fnMidValue(LStr, 91, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 93, 3)) & "','" & Trim(fnMidValue(LStr, 96, 8)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 104, 2)) & "','" & Trim(fnMidValue(LStr, 106, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 109, 8)) & "','" & Trim(fnMidValue(LStr, 117, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 119, 3)) & "','" & Trim(fnMidValue(LStr, 122, 16)) & "')"
                'by Abhi on 06-Mar-2017 for caseid 7248 Special charcater in Currency code while saving PNR File
                dbCompany.Execute SQL
            
            'A08 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A08" Then
                LastSec = "A08"
                'by Abhi on 10-May-2014 for caseid 3840 Baggage allowance for galileo from A08 line
                'SQL = "INSERT INTO GalA08 (UPLOADNO,A08SEC,A08FSI,A08ITN,A08FBC,A08VAL,A08NVBC," & _
                    "A08NVAC,A08TDGC,A08ENDI,A08END) values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                    "" & Val(fnMidValue(LStr, 4, 2)) & "," & Val(fnMidValue(LStr, 6, 2)) & "," & _
                    "'" & Trim(fnMidValue(LStr, 8, 8)) & "','" & Trim(fnMidValue(LStr, 16, 8)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 24, 7)) & "','" & Trim(fnMidValue(LStr, 31, 7)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 38, 6)) & "','" & Trim(fnMidValue(LStr, 44, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 46, 60)) & "')"
                SQL = "INSERT INTO GalA08 (UPLOADNO,A08SEC,A08FSI,A08ITN,A08FBC,A08VAL,A08NVBC," & _
                    "A08NVAC,A08TDGC,A08ENDI,A08END,A08BAGI,A08BAG) values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                    "" & Val(fnMidValue(LStr, 4, 2)) & "," & Val(fnMidValue(LStr, 6, 2)) & "," & _
                    "'" & Trim(fnMidValue(LStr, 8, 8)) & "','" & Trim(fnMidValue(LStr, 16, 8)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 24, 7)) & "','" & Trim(fnMidValue(LStr, 31, 7)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 38, 6)) & "','" & Trim(fnMidValue(LStr, 44, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 46, 60)) & "','" & SkipChars(fnMidValue(LStr, 123, 2)) & "'," & _
                    "'" & SkipChars(fnMidValue(LStr, 125, 3)) & "')"
                'by Abhi on 10-May-2014 for caseid 3840 Baggage allowance for galileo from A08 line
                    
                dbCompany.Execute SQL
                
            'A09 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A09" Then
                LastSec = "A09"
                SQL = "INSERT INTO GalA09 (UPLOADNO,A09SEC,A09FSI,A09TY5,A09L51) Values (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & Val(fnMidValue(LStr, 4, 2)) & "," & _
                    "" & Val(fnMidValue(LStr, 6, 1)) & ",'" & Trim(fnMidValue(LStr, 7, 61)) & "')"
                
                dbCompany.Execute SQL
            'A10 Table
            'by Abhi on 26-May-2014 for caseid 4088 Galileo Reissue Ticket
            ElseIf fnMidValue(LStr, 1, 3) = "A10" Then
                LastSec = fnMidValue(LStr, 1, 3)
                If fMIRTypeID_String = "GCS" Then 'GALILEO
                    Splited = SplitWithLengths(LStr, 3, 2, 7, 8, 4, 9, 9, 19, 9, 9, 1)
                    SQL = "INSERT INTO GalA10 (UPLOADNO,A10SEC,A10EXI,A10DOI,A10POI,A10CDEP,A10OCM,A10OIN,A10FOP,A10PEN,A10SCC,A10TYP) VALUES (" & UploadNo & "," & _
                        "'" & SkipChars(Splited(0)) & "','" & Val(Splited(1)) & "','" & SkipChars(Splited(2)) & "','" & SkipChars(Splited(3)) & "'," & _
                        "'" & SkipChars(Splited(4)) & "','" & SkipChars(Splited(5)) & "','" & SkipChars(Splited(6)) & "','" & SkipChars(Splited(7)) & "'," & _
                        "'" & SkipChars(Splited(8)) & "','" & SkipChars(Splited(9)) & "','" & SkipChars(Splited(10)) & "'" & _
                        ")"
                    dbCompany.Execute SQL
                ElseIf fMIRTypeID_String = "APO" Then 'APOLLO
                    Splited = SplitWithLengths(LStr, 3, 2, 13, 1, 2, 4, 9, 1, 3, 12, 9, 9)
                    SQL = "" _
                        & "INSERT INTO GalA10APO " _
                        & "                      (UPLOADNO, A10SEC, A10EXI, A10OTN, A10BLK, A10ONB, A10OCI, A10OIN, A10TYP, A10CUR, A10OTF, A10PEN, A10SCC) " _
                        & "VALUES     (" _
                            & UploadNo & ", N'" & SkipChars(Splited(0)) & "', " & Val(Splited(1)) & ", N'" & SkipChars(Splited(2)) & "'" _
                            & ", N'" & SkipChars(Splited(3)) & "', N'" & SkipChars(Splited(4)) & "', N'" & SkipChars(Splited(5)) & "', N'" & SkipChars(Splited(6)) & "'" _
                            & ", N'" & SkipChars(Splited(7)) & "', N'" & SkipChars(Splited(8)) & "', N'" & SkipChars(Splited(9)) & "', N'" & SkipChars(Splited(10)) & "'" _
                            & ", N'" & SkipChars(Splited(11)) & "'" _
                        & ")"
                    dbCompany.Execute SQL
                End If
            'by Abhi on 26-May-2014 for caseid 4088 Galileo Reissue Ticket
            'A11 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A11" Then
                LastSec = "A11"
                'by Abhi on 26-May-2014 for caseid 4088 Galileo Reissue Ticket
                'SQL = "INSERT INTO GalA11 (UPLOADNO,A11SEC,A11TYP,A11AMT,A11REF,A11CCC,A11CCN,A11EXP,A11APP,A11BLK,A11MAN,A11APC) VALUES (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 6, 12)) & "','" & Trim(fnMidValue(LStr, 18, 1)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 19, 2)) & "','" & Trim(fnMidValue(LStr, 21, 20)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 41, 4)) & "','" & Trim(fnMidValue(LStr, 45, 8)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 53, 1)) & "','" & Trim(fnMidValue(LStr, 54, 1)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 55, 8)) & "')"
                'by Abhi on 10-Jul-2014 for caseid 4278 Galileo error invalid procedure call or argument at sec A11
                'vA11Last_String = fnMidValue(LStr, 56, Len(LStr) - 55)
                If Len(LStr) >= 59 Then
                    vA11Last_String = fnMidValue(LStr, 56, Len(LStr) - 55)
                End If
                'by Abhi on 10-Jul-2014 for caseid 4278 Galileo error invalid procedure call or argument at sec A11
                vA11LastStart_Long = InStr(1, vA11Last_String, "P:", vbTextCompare)
                If vA11LastStart_Long > 0 Then
                    vA11Lastspecific_String = Mid(vA11Last_String, vA11LastStart_Long, 2 + 2)
                End If
                Splited = SplitWithLengths(vA11Lastspecific_String, 2, 2)
                vA11PGRI_String = Splited(0)
                vA11PGR_String = Splited(1)
                SQL = "INSERT INTO GalA11 (UPLOADNO,A11SEC,A11TYP,A11AMT,A11REF,A11CCC,A11CCN,A11EXP,A11APP,A11BLK,A11MAN,A11APC,A11PPO,A11PGRI,A11PGR) VALUES (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 6, 12)) & "','" & Trim(fnMidValue(LStr, 18, 1)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 19, 2)) & "','" & Trim(fnMidValue(LStr, 21, 20)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 41, 4)) & "','" & Trim(fnMidValue(LStr, 45, 8)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 45, 1)) & "','" & Trim(fnMidValue(LStr, 46, 1)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 47, 6)) & "','" & Trim(fnMidValue(LStr, 53, 3)) & "'," & _
                    "'" & vA11PGRI_String & "'," & Val(vA11PGR_String) & "" & _
                    ")"
                'by Abhi on 26-May-2014 for caseid 4088 Galileo Reissue Ticket
                dbCompany.Execute SQL
                
            'A12 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A12" Then
                LastSec = "A12"
                SQL = "INSERT INTO GalA12 (UPLOADNO,A12SEC,A12CTY,A12LOC,A12PHA) VALUES (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 7, 2)) & "','" & Trim(fnMidValue(LStr, 9, 61)) & "')"
                dbCompany.Execute SQL
            'by Abhi on 08-Feb-2014 for caseid 3719 Delivery Address from Gal file
            'A13 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A13" Then
                LastSec = "A13"
                'by Abhi on 14-Feb-2014 for caseid 3741 Pengds -Galileo file stuck Incorrect syntax near ARCY
                'SQL = "INSERT INTO GalA13 (UPLOADNO,A13SEC,A13ADT,A13DTA) VALUES (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 6, 223)) & "')"
                SQL = "INSERT INTO GalA13 (UPLOADNO,A13SEC,A13ADT,A13DTA) VALUES (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 2)) & "'," & _
                    "'" & SkipChars(fnMidValue(LStr, 6, 223)) & "')"
                'by Abhi on 14-Feb-2014 for caseid 3741 Pengds -Galileo file stuck Incorrect syntax near ARCY
                dbCompany.Execute SQL
            'by Abhi on 08-Feb-2014 for caseid 3719 Delivery Address from Gal file
            'A14 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A14" Then
                LastSec = "A14"
                'by Abhi on 11-Jan-2010 for caseid 1588 Galileo Penline PENAUTOOFF
                LStr = Replace(LStr, Chr(13), "")
                If InStr(UCase(LStr), "PENAUTOOFF") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENAUTOOFF") > 0 Then
                        LStr = Left(LStr, 3) & "PENAUTOOFF" & Mid(LStr, InStr(1, LStr, "PENAUTOOFF", vbTextCompare) + Len("PENAUTOOFF"))
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & Trim(fnMidValue(LStr, 4, 10)) & "')"
                        dbCompany.Execute SQL
                    End If
                'by Abhi on 11-Jan-2010 for caseid 1588 Galileo Penline AC BB
                'by Abhi on 25-Feb-2011 for caseid 1618 PenGDS GDS intray type mismatch with ARNK and penline
                'ElseIf InStr(UCase(LStr), "PEN/") > 0 Then
                'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
                'ElseIf InStr(UCase(LStr), "PEN/") > 0 And (InStr(UCase(LStr), "AC-") > 0 Or InStr(UCase(LStr), "BB-") > 0) Then
                'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                'ElseIf InStr(UCase(LStr), "PEN/") > 0 And (InStr(UCase(LStr), "AC-") > 0 Or InStr(UCase(LStr), "BB-") > 0 Or InStr(UCase(LStr), "DES-") > 0 Or InStr(UCase(LStr), "REFE-") > 0 Or InStr(UCase(LStr), "CC-") > 0 Or InStr(UCase(LStr), "BRANCH-") > 0 Or InStr(UCase(LStr), "MC-") > 0 Or InStr(UCase(LStr), "INETREF-") > 0) Then
                'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                'ElseIf InStr(UCase(LStr), "PEN/") > 0 And (InStr(UCase(LStr), "AC-") > 0 Or InStr(UCase(LStr), "BB-") > 0 Or InStr(UCase(LStr), "DES-") > 0 Or InStr(UCase(LStr), "REFE-") > 0 Or InStr(UCase(LStr), "CC-") > 0 Or InStr(UCase(LStr), "BRANCH-") > 0 Or InStr(UCase(LStr), "MC-") > 0 Or InStr(UCase(LStr), "INETREF-") > 0 Or InStr(UCase(LStr), "TKTDEADLINE-") > 0) Then
                'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                'ElseIf InStr(UCase(LStr), "PEN/") > 0 And (InStr(UCase(LStr), "AC-") > 0 Or InStr(UCase(LStr), "BB-") > 0 Or InStr(UCase(LStr), "DES-") > 0 Or InStr(UCase(LStr), "REFE-") > 0 Or InStr(UCase(LStr), "CC-") > 0 Or InStr(UCase(LStr), "BRANCH-") > 0 Or InStr(UCase(LStr), "MC-") > 0 Or InStr(UCase(LStr), "INETREF-") > 0 Or InStr(UCase(LStr), "TKTDEADLINE-") > 0 Or InStr(UCase(LStr), "DEPT-") > 0) Then
                ElseIf InStr(UCase(LStr), "PEN/") > 0 And (InStr(UCase(LStr), "AC-") > 0 Or InStr(UCase(LStr), "BB-") > 0 _
                    Or InStr(UCase(LStr), "DES-") > 0 Or InStr(UCase(LStr), "REFE-") > 0 Or InStr(UCase(LStr), "CC-") > 0 _
                    Or InStr(UCase(LStr), "BRANCH-") > 0 Or InStr(UCase(LStr), "MC-") > 0 Or InStr(UCase(LStr), "INETREF-") > 0 _
                    Or InStr(UCase(LStr), "TKTDEADLINE-") > 0 Or InStr(UCase(LStr), "DEPT-") > 0 Or InStr(UCase(LStr), "DEPOSITAMT-") > 0 _
                    Or InStr(UCase(LStr), "DEPOSITDUEDATE-") > 0) Then
                'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
                    If InStr(UCase(LStr), PENLINEID_String & "PEN/") > 0 Then
                        Set CollPenSplited = A14PEN_PENLINE(LStr)
                        'LStr = Left(LStr, 3) & "PENAUTOOFF" & Mid(LStr, InStr(1, LStr, "PENAUTOOFF", vbTextCompare) + Len("PENAUTOOFF"))
                        'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
                        'SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14AC,A14BB) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','PEN','" & SkipChars(Left(CollPenSplited("AC"), 50)) & "','" & SkipChars(Left(CollPenSplited("BB"), 50)) & "')"
                        'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                        'SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14AC,A14BB,A14DES,A14REFE,A14CUSTCC,A14BRANCH,A14MC,A14INETREF) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','PEN','" & SkipChars(Left(CollPenSplited("AC"), 50)) & "','" & SkipChars(Left(CollPenSplited("BB"), 50)) & "','" & SkipChars(Left(CollPenSplited("DES"), 50)) & "','" & SkipChars(Left(CollPenSplited("REFE"), 50)) & "','" & SkipChars(Left(CollPenSplited("CUSTCC"), 200)) & "','" & SkipChars(Left(CollPenSplited("BRANCH"), 50)) & "','" & SkipChars(Left(CollPenSplited("MC"), 50)) & "','" & SkipChars(Left(CollPenSplited("INETREF"), 100)) & "')"
                        'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                        'SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14AC,A14BB,A14DES,A14REFE,A14CUSTCC,A14BRANCH,A14MC,A14INETREF,A14TKTDEADLINE) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','PEN','" & SkipChars(Left(CollPenSplited("AC"), 50)) & "','" & SkipChars(Left(CollPenSplited("BB"), 50)) & "','" & SkipChars(Left(CollPenSplited("DES"), 50)) & "','" & SkipChars(Left(CollPenSplited("REFE"), 50)) & "','" & SkipChars(Left(CollPenSplited("CUSTCC"), 200)) & "','" & SkipChars(Left(CollPenSplited("BRANCH"), 50)) & "','" & SkipChars(Left(CollPenSplited("MC"), 50)) & "','" & SkipChars(Left(CollPenSplited("INETREF"), 100)) & "'," & _
                            "'" & SkipChars(Left(CollPenSplited("TKTDEADLINE"), 9)) & "')"
                        'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14AC,A14BB,A14DES,A14REFE,A14CUSTCC,A14BRANCH,A14MC,A14INETREF,A14TKTDEADLINE,A14DEPT,A14DEPOSITAMT,A14DEPOSITDUEDATE) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','PEN','" & SkipChars(Left(CollPenSplited("AC"), 50)) & "','" & SkipChars(Left(CollPenSplited("BB"), 50)) & "','" & SkipChars(Left(CollPenSplited("DES"), 50)) & "','" & SkipChars(Left(CollPenSplited("REFE"), 50)) & "','" & SkipChars(Left(CollPenSplited("CUSTCC"), 200)) & "','" & SkipChars(Left(CollPenSplited("BRANCH"), 50)) & "','" & SkipChars(Left(CollPenSplited("MC"), 50)) & "','" & SkipChars(Left(CollPenSplited("INETREF"), 100)) & "'," & _
                            "'" & SkipChars(Left(CollPenSplited("TKTDEADLINE"), 9)) & "', '" & SkipChars(Left(CollPenSplited("A14DEPT"), 50)) & "', " & Val(CollPenSplited("A14DEPOSITAMT")) & ", '" & SkipChars(CollPenSplited("A14DEPOSITDUEDATE")) & "')"
                        'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                        'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                        'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                        'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
                        dbCompany.Execute SQL
                        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
                        If Trim(Left(CollPenSplited("AC"), 50)) <> "" Then
                            GIT_CUSTOMERUSERCODE_String = Left(CollPenSplited("AC"), 50)
                        End If
                        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
                    End If
                'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
                ElseIf InStr(UCase(LStr), "PEN/") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PEN/") > 0 Then
                        For TempPENPNO = 1 To 99
                            'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
                            'If InStr(1, LStr, "E" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, LStr, "T" & TempPENPNO & "-", vbTextCompare) > 0 Then
                            If InStr(1, LStr, "E" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, LStr, "T" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, LStr, "DOB" & TempPENPNO & "-", vbTextCompare) > 0 Then
                            'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
                                Set CollPenSplited = A14PEN_PENLINEPassenger(LStr, TempPENPNO)
                                'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
                                'SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PEMAIL,A14PTELE,A14PNO) values (" & UploadNo & "," & _
                                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','PEN','" & SkipChars(Left(CollPenSplited("PEMAIL"), 200)) & "','" & SkipChars(Left(CollPenSplited("PTELE"), 50)) & "'," & Val(CollPenSplited("PNO")) & ")"
                                'by Abhi on 11-Oct-2017 for caseid 7862 Passenger Email id validation in folder should also save multiple email id with coma separator
                                'SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PEMAIL,A14PTELE,A14PNO,A14PDOB) values (" & UploadNo & "," & _
                                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','PEN','" & SkipChars(Left(CollPenSplited("PEMAIL"), 200)) & "','" & SkipChars(Left(CollPenSplited("PTELE"), 50)) & "'," & Val(CollPenSplited("PNO")) & ",'" & SkipChars(Left(CollPenSplited("PDOB"), 9)) & "')"
                                SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PEMAIL,A14PTELE,A14PNO,A14PDOB) values (" & UploadNo & "," & _
                                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','PEN','" & SkipChars(CollPenSplited("PEMAIL")) & "','" & SkipChars(Left(CollPenSplited("PTELE"), 50)) & "'," & Val(CollPenSplited("PNO")) & ",'" & SkipChars(Left(CollPenSplited("PDOB"), 9)) & "')"
                                'by Abhi on 11-Oct-2017 for caseid 7862 Passenger Email id validation in folder should also save multiple email id with coma separator
                                'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
                                dbCompany.Execute SQL
                            End If
                            DoEvents
                        Next
                    End If
                ElseIf InStr(UCase(LStr), "PENRT") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENRT") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENRT" & Mid(LStr, InStr(1, LStr, "PENRT", vbTextCompare) + Len("PENRT"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENRT) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENPOL") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENPOL") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENPOL" & Mid(LStr, InStr(1, LStr, "PENPOL", vbTextCompare) + Len("PENPOL"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENPOL) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENPROJ") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENPROJ") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENPROJ" & Mid(LStr, InStr(1, LStr, "PENPROJ", vbTextCompare) + Len("PENPROJ"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENPROJ) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENCC") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENCC") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENCC" & Mid(LStr, InStr(1, LStr, "PENCC", vbTextCompare) + Len("PENCC"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14CUSTCC) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENEID") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENEID") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENEID" & Mid(LStr, InStr(1, LStr, "PENEID", vbTextCompare) + Len("PENEID"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENEID) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENPO") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENPO") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENPO" & Mid(LStr, InStr(1, LStr, "PENPO", vbTextCompare) + Len("PENPO"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENPO) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 255)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENHFRC") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENHFRC") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENHFRC" & Mid(LStr, InStr(1, LStr, "PENHFRC", vbTextCompare) + Len("PENHFRC"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENHFRC) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENLFRC") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENLFRC") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENLFRC" & Mid(LStr, InStr(1, LStr, "PENLFRC", vbTextCompare) + Len("PENLFRC"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENLFRC) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENHIGHF") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENHIGHF") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENHIGHF" & Mid(LStr, InStr(1, LStr, "PENHIGHF", vbTextCompare) + Len("PENHIGHF"))
                        'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                        'aPEN = Split(LStr, "-")
                        vLStrPEN_String = "PENHIGHF" & Mid(LStr, InStr(1, LStr, "PENHIGHF", vbTextCompare) + Len("PENHIGHF"))
                        vLStrPEN_String = Replace(vLStrPEN_String, "/", "-", , , vbTextCompare)
                        aPEN = Split(vLStrPEN_String, "-")
                        'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                        ReDim Preserve aPEN(3)
                        'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                        'SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENHIGHF) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "'," & Val(aPEN(2)) & ")"
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENHIGHF,A14PENHIGHFTKTNO) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(0) & "'," & Val(aPEN(1)) & ",'" & Left(SkipChars(aPEN(3)), 50) & "')"
                        'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENLOWF") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENLOWF") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENLOWF" & Mid(LStr, InStr(1, LStr, "PENLOWF", vbTextCompare) + Len("PENLOWF"))
                        'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                        'aPEN = Split(LStr, "-")
                        vLStrPEN_String = "PENLOWF" & Mid(LStr, InStr(1, LStr, "PENLOWF", vbTextCompare) + Len("PENLOWF"))
                        vLStrPEN_String = Replace(vLStrPEN_String, "/", "-", , , vbTextCompare)
                        aPEN = Split(vLStrPEN_String, "-")
                        'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                        ReDim Preserve aPEN(3)
                        'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                        'SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENLOWF) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "'," & Val(aPEN(2)) & ")"
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENLOWF,A14PENLOWFTKTNO) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(0) & "'," & Val(aPEN(1)) & ",'" & Left(SkipChars(aPEN(3)), 50) & "')"
                        'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENUC1") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENUC1") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENUC1" & Mid(LStr, InStr(1, LStr, "PENUC1", vbTextCompare) + Len("PENUC1"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENUC1) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENUC2") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENUC2") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENUC2" & Mid(LStr, InStr(1, LStr, "PENUC2", vbTextCompare) + Len("PENUC2"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENUC2) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENUC3") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENUC3") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENUC3" & Mid(LStr, InStr(1, LStr, "PENUC3", vbTextCompare) + Len("PENUC3"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENUC3) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENBB") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENBB") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENBB" & Mid(LStr, InStr(1, LStr, "PENBB", vbTextCompare) + Len("PENBB"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14BB) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENAGROSS/") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENAGROSS/") > 0 Then
                        LStr = Replace(LStr, "-", "/")
                        LStr = Left(LStr, 6) & "PENAGROSS" & Mid(LStr, InStr(1, LStr, "PENAGROSS", vbTextCompare) + Len("PENAGROSS"))
                        'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
                        'aPEN = SplitForce(LStr, "/", 6)
                        'SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENAGROSSADULT,A14PENAGROSSCHILD,A14PENAGROSSINFANT,A14PENAGROSSPACKAGE) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "'," & Val(aPEN(2)) & "," & Val(aPEN(3)) & "," & Val(aPEN(4)) & "," & Val(aPEN(5)) & ")"
                        aPEN = SplitForce(LStr, "/", 7)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENAGROSSADULT,A14PENAGROSSCHILD,A14PENAGROSSINFANT,A14PENAGROSSPACKAGE,A14PENAGROSSYOUTH) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "'," & Val(aPEN(2)) & "," & Val(aPEN(3)) & "," & Val(aPEN(4)) & "," & Val(aPEN(5)) & "," & Val(aPEN(6)) & ")"
                        'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENFARE/") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENFARE/") > 0 Then
                        LStr = Replace(LStr, "-", "/")
                        LStr = Left(LStr, 6) & "PENFARE" & Mid(LStr, InStr(1, LStr, "PENFARE", vbTextCompare) + Len("PENFARE"))
                        'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
                        'aPEN = SplitForce(LStr, "/", 7)
                        'SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENFAREFARE,A14PENFARETAXES,A14PENFARESELL,A14PENFAREPASSTYPE,A14PENFARESUPPID) values (" & UploadNo & "," & _
                        '    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "'," & Val(aPEN(2)) & "," & Val(aPEN(3)) & "," & Val(aPEN(4)) & ",'" & SkipChars(Left(aPEN(5), 50)) & "','" & SkipChars(Left(aPEN(6), 50)) & "')"
                        aPEN = SplitForce(LStr, "/", 8)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENFAREFARE,A14PENFARETAXES,A14PENFARESELL,A14PENFAREPASSTYPE,A14PENFARESUPPID,A14PENFARETICKETTYPE) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "'," & Val(aPEN(2)) & "," & Val(aPEN(3)) & "," & Val(aPEN(4)) & ",'" & SkipChars(Left(aPEN(5), 50)) & "','" & SkipChars(Left(aPEN(6), 50)) & "','" & SkipChars(Left(aPEN(7), 50)) & "')"
                        'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
                        dbCompany.Execute SQL
                    End If
                ElseIf InStr(UCase(LStr), "PENO/") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENO/") > 0 Then
                        LStr = Replace(LStr, "-", "/")
                        LStr = Left(LStr, 6) & "PENO" & Mid(LStr, InStr(1, LStr, "PENO", vbTextCompare) + Len("PENO"))
                        aPEN = SplitForce(LStr, "/", 8)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENOPRDID,A14PENOQTY,A14PENORATE,A14PENOSELL,A14PENOSUPPID,A14PENOPAYMETHODID) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "','" & SkipChars(Left(aPEN(2), 50)) & "'," & Val(aPEN(3)) & "," & Val(aPEN(4)) & "," & Val(aPEN(5)) & ",'" & SkipChars(Left(aPEN(6), 50)) & "','" & SkipChars(Left(aPEN(7), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
                ElseIf InStr(UCase(LStr), "PENAIRTKT/") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENAIRTKT/") > 0 Then
                        Set CollPenSplited = A14PEN_PENAIRTKT(LStr)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14AIRTKTPAX,A14AIRTKTTKT,A14AIRTKTDATE) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','PENAIRTKT','" & SkipChars(CollPenSplited("PAX")) & "','" & SkipChars(CollPenSplited("TKT")) & "','" & SkipChars(CollPenSplited("DATE")) & "')"
                        dbCompany.Execute SQL
                    End If
                'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
                'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
                ElseIf InStr(UCase(LStr), "PENBILLCUR") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENBILLCUR") > 0 Then
                        LStr = Left(LStr, 6) & "PENBILLCUR" & Mid(LStr, InStr(1, LStr, "PENBILLCUR", vbTextCompare) + Len("PENBILLCUR"))
                        aPEN = Split(LStr, "-")
                        ReDim Preserve aPEN(2)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENBILLCUR) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & SkipChars(aPEN(1)) & "','" & SkipChars(Left(aPEN(2), 4)) & "')"
                        dbCompany.Execute SQL
                    End If
                'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
                'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
                ElseIf InStr(UCase(LStr), "PENWAIT") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENWAIT") > 0 Then
                        LStr = Left(LStr, 3) & "PENWAIT" & Mid(LStr, InStr(1, LStr, "PENWAIT", vbTextCompare) + Len("PENWAIT"))
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & Trim(fnMidValue(LStr, 4, 7)) & "')"
                        dbCompany.Execute SQL
                        GIT_PENWAIT_String = "Y"
                    End If
                'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
                'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
                ElseIf InStr(UCase(LStr), "PENVC") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENVC") > 0 Then
                        LStr = Left(LStr, 3) & "PENVC" & Mid(LStr, InStr(1, LStr, "PENVC", vbTextCompare) + Len("PENVC"))
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & Trim(fnMidValue(LStr, 4, 5)) & "')"
                        dbCompany.Execute SQL
                    End If
                'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
                'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
                'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
                ElseIf InStr(UCase(LStr), "PENCS/") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENCS/") > 0 Then
                        LStr = Replace(LStr, "-", "/")
                        LStr = Left(LStr, 6) & "PENCS" & Mid(LStr, InStr(1, LStr, "PENCS", vbTextCompare) + Len("PENCS"))
                        aPEN = SplitForce(LStr, "/", 4)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENCSLABELID,A14PENCSLISTID) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(1) & "',N'" & SkipChars(Left(Trim(aPEN(2)), 20)) & "',N'" & SkipChars(Left(Trim(aPEN(3)), 20)) & "')"
                        dbCompany.Execute SQL
                    End If
                'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                ElseIf InStr(UCase(LStr), "PENRC") > 0 Then
                    If InStr(UCase(LStr), PENLINEID_String & "PENRC") > 0 Then
                        'LStr=replace(LStr,""
                        LStr = Left(LStr, 6) & "PENRC" & Mid(LStr, InStr(1, LStr, "PENRC", vbTextCompare) + Len("PENRC"))
                        vLStrPEN_String = "PENRC" & Mid(LStr, InStr(1, LStr, "PENRC", vbTextCompare) + Len("PENRC"))
                        vLStrPEN_String = Replace(vLStrPEN_String, "/", "-", , , vbTextCompare)
                        aPEN = Split(vLStrPEN_String, "-")
                        ReDim Preserve aPEN(3)
                        SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID,A14PEN,A14PENRC,A14PENRCTKTNO) values (" & UploadNo & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 3)) & "','','','" & aPEN(0) & "','" & SkipChars(Left(aPEN(1), 50)) & "','" & SkipChars(Left(aPEN(3), 50)) & "')"
                        dbCompany.Execute SQL
                    End If
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                Else
                    'by Abhi on 09-Nov-2009 VLOCATOR in AirSegDetails
                    'SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK) values (" & UploadNo & "," & _
                        "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 64)) & "')"
                    SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK,A14VLVendorID) values (" & UploadNo & "," & _
                        "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 64)) & "','" & Trim(fnMidValue(LStr, 21, 2)) & "')"
                    dbCompany.Execute SQL
                End If
            
            'A15 Table by Abhi
            ElseIf fnMidValue(LStr, 1, 3) = "A15" Then
                LastSec = "A15"
                SQL = "INSERT INTO GalA15 (UPLOADNO,A15SEC,A15SEG,A15RMK) values (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 2)) & "','" & Trim(fnMidValue(LStr, 6, 87)) & "')"
                dbCompany.Execute SQL
            
            'by Abhi on 03-Dec-2009 for Galileo Hotel HHL PNR uploading
            'A16 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A16" Then
                ''By jiji 31/Dec/2012 for case id 2760
                'by Abhi on 29-Sep-2015 for caseid 5532 Galileo HHL SEGMENTS
                'If fnMidValue(LStr, 4, 1) = "1" Or fnMidValue(LStr, 4, 1) = "7" Then
                If fnMidValue(LStr, 4, 1) = "1" Or fnMidValue(LStr, 4, 1) = "7" Or fnMidValue(LStr, 4, 1) = "A" Then
                'by Abhi on 29-Sep-2015 for caseid 5532 Galileo HHL SEGMENTS
                
                    LastSec = "A16HHL"
                    LastSecNo_String = Trim(fnMidValue(LStr, 5, 2))
                    'by Abhi on 23-Feb-2011 for caseid 1617 PenGDS Galileo Incorrect syntax near AEROPORTO
                    SQL = "INSERT INTO GalA16HHL (UPLOADNO,A16SEC,A16TYP,A16NUM,A16DTE,A16PRP,A16HCC,A16CTY,A16MUL,A16STT,A16OUT,A16DAY," & _
                          "A16NME,A16FON,A16FAX,A16RTT,A16RMS,A16LOC) values (" & UploadNo & "," & _
                          "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 1)) & "','" & Trim(fnMidValue(LStr, 5, 2)) & "'," & _
                          "'" & Trim(fnMidValue(LStr, 7, 7)) & "','" & Trim(fnMidValue(LStr, 14, 6)) & "','" & Trim(fnMidValue(LStr, 20, 2)) & "'," & _
                          "'" & Trim(fnMidValue(LStr, 22, 4)) & "','" & Trim(fnMidValue(LStr, 26, 1)) & "','" & Trim(fnMidValue(LStr, 27, 2)) & "'," & _
                          "'" & Trim(fnMidValue(LStr, 29, 5)) & "','" & Trim(fnMidValue(LStr, 34, 3)) & "','" & Trim(fnMidValue(LStr, 37, 20)) & "'," & _
                          "'" & Trim(fnMidValue(LStr, 57, 17)) & "','" & Trim(fnMidValue(LStr, 74, 17)) & "','" & Trim(fnMidValue(LStr, 91, 1)) & "'," & _
                          "'" & Trim(fnMidValue(LStr, 92, 8)) & "','" & SkipChars(Trim(fnMidValue(LStr, 100, 20))) & "')"
                    dbCompany.Execute SQL
                'by Abhi on 13-May-2019 for caseid 7191 Galileo Car hire mapping
                ElseIf fnMidValue(LStr, 4, 1) = "B" Or fnMidValue(LStr, 4, 1) = "8" Then
                    LastSec = "A16CCRCAR"
                    LastSecNo_String = Trim(fnMidValue(LStr, 5, 2))
                    Splited = SplitWithLengths(LStr, 3, 1, 2, 7, 12, 2, 5, 3, 4, 2, 4, 1, 26, 26, 40)
                    SQL = "" _
                        & "INSERT INTO GalA16CCRCAR (UPLOADNO, A16SEC, A16TYP, A16NUM, A16DTE, A16CAR, A16STA, A16CDT, A16DAY, A16CCC, A16CVC, A16CCT, A16CNI, A16PUP, A16DOL, A16PHN) " _
                        & "VALUES (" & UploadNo & ",'" & SkipChars(Trim(Splited(0))) & "','" & SkipChars(Trim(Splited(1))) & "'," & Val(Trim(Splited(2))) & ",'" & SkipChars(Trim(Splited(3))) & "','" & SkipChars(Trim(Splited(4))) & "', " _
                        & "'" & SkipChars(Trim(Splited(5))) & "','" & SkipChars(Trim(Splited(6))) & "'," & Val(Trim(Splited(7))) & ",'" & SkipChars(Trim(Splited(8))) & "','" & SkipChars(Trim(Splited(9))) & "','" & SkipChars(Trim(Splited(10))) & "', " _
                        & " " & Val(Trim(Splited(11))) & ",'" & SkipChars(Trim(Splited(12))) & "','" & SkipChars(Trim(Splited(13))) & "','" & SkipChars(Trim(Splited(14))) & "'" _
                        & ")"
                    dbCompany.Execute SQL
                'by Abhi on 13-May-2019 for caseid 7191 Galileo Car hire mapping
                End If
            
'            ElseIf fnMidValue(LStr, 1, 4) = "PEN\" Then
'                LastSec = "PEN"
'                Set Coll = RecPEN_PENLINE(LStr)
'                InsertDataByFieldName Coll, "GalPen", "PEN", UploadNo
            'by Abhi on 11-May-2019 for caseid 1656 Galileo-A22 Seat Number
            'A22 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A22" Then
                LastSec = "A22"
                LastSec = fnMidValue(LStr, 1, 3)
                Splited = SplitWithLengths(LStr, 3, 2, 2, 1, 6, 20)
                SQL = "INSERT INTO GalA22 (UPLOADNO,A22SEC,A22SEG,A22PAS,A22STT,A22SEN,A22SCHX) VALUES (" & UploadNo & "," & _
                    "'" & SkipChars(Splited(0)) & "'," & Val(Splited(1)) & "," & Val(Splited(2)) & ",'" & SkipChars(Splited(3)) & "','" & SkipChars(Splited(4)) & "','" & SkipChars(Splited(5)) & "'" & _
                    ")"
                dbCompany.Execute SQL
            'by Abhi on 11-May-2019 for caseid 1656 Galileo-A22 Seat Number
            'A27 Table
            'by Abhi on 06-Apr-2015 for caseid 5078 Galileo -Carrier Fee should be added with Airticketdetails ->Sell1 field
            ElseIf fnMidValue(LStr, 1, 3) = "A27" Then
                LastSec = fnMidValue(LStr, 1, 3)
                Splited = SplitWithLengths(LStr, 3, 1, 1, 2, 3, 12, 3, 12)
                SQL = "INSERT INTO GalA27 (UPLOADNO,A27SEC,A27CFC,A27MOD,A27FSI,A27CRT,A27TTA,A27CRG,A27GTA) VALUES (" & UploadNo & "," & _
                    "'" & SkipChars(Splited(0)) & "','" & SkipChars(Splited(1)) & "','" & SkipChars(Splited(2)) & "','" & SkipChars(Splited(3)) & "'," & _
                    "" & Val(Splited(4)) & ",'" & SkipChars(Splited(5)) & "','" & SkipChars(Splited(6)) & "','" & SkipChars(Splited(7)) & "'" & _
                    ")"
                dbCompany.Execute SQL
            'by Abhi on 06-Apr-2015 for caseid 5078 Galileo -Carrier Fee should be added with Airticketdetails ->Sell1 field
            'by Abhi on 11-Feb-2014 for caseid 3721 Galileo EMD documents A29 and A30 section
            'A29 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A29" Then
                LastSec = fnMidValue(LStr, 1, 3)
                Splited = SplitWithLengths(LStr, 3, 2, 2, 4, 1, 1, 1, 55, 13, 13, 2, 3, 12, 3, 12, 3, 12, 3, 12, 13, 2, 12, 1, 8)
                'by Abhi on 12-Feb-2014 for caseid 3726 A29 EMDVAR (EMD VARIABLE DATA)
                fA29PAX_String = Splited(1)
                'by Abhi on 12-Feb-2014 for caseid 3726 A29 EMDVAR (EMD VARIABLE DATA)
                SQL = "INSERT INTO GalA29 (UPLOADNO,A29SEC,A29PAX,A29SEG,A29SVC,A29BKI,A29EID,A29RFI,A29NAM,A29EMD,A29CNJ,A29VCR,A29CBA,A29BAM,A29CEA,A29EAM,A29CTX,A29TAM,A29CTA,A29TLA,A29BSR,A29TYP,A29AMT,A29ACI,A29APC,A29EDS,A29EDI,A29EMDVAR) VALUES (" & UploadNo & "," & _
                    "'" & SkipChars(Splited(0)) & "','" & SkipChars(Splited(1)) & "','" & SkipChars(Splited(2)) & "','" & SkipChars(Splited(3)) & "'," & _
                    "'" & SkipChars(Splited(4)) & "','" & SkipChars(Splited(5)) & "','" & SkipChars(Splited(6)) & "','" & SkipChars(Splited(7)) & "'," & _
                    "'" & SkipChars(Splited(8)) & "','" & SkipChars(Splited(9)) & "','" & SkipChars(Splited(10)) & "','" & SkipChars(Splited(11)) & "'," & _
                    "'" & SkipChars(Splited(12)) & "','" & SkipChars(Splited(13)) & "','" & SkipChars(Splited(14)) & "','" & SkipChars(Splited(15)) & "'," & _
                    "'" & SkipChars(Splited(16)) & "','" & SkipChars(Splited(17)) & "','" & SkipChars(Splited(18)) & "','" & SkipChars(Splited(19)) & "'," & _
                    "'" & SkipChars(Splited(20)) & "','" & SkipChars(Splited(21)) & "','" & SkipChars(Splited(22)) & "','" & SkipChars(Splited(23)) & "'," & _
                    "'" & "" & "','" & "" & "','" & "" & "'" & _
                    ")"
                dbCompany.Execute SQL
            'by Abhi on 11-Feb-2014 for caseid 3721 Galileo EMD documents A29 and A30 section
            'by Abhi on 12-Feb-2014 for caseid 3726 A29 EMDVAR (EMD VARIABLE DATA)
            ElseIf fnMidValue(LStr, 1, 3) = "A30" Then
                LastSec = fnMidValue(LStr, 1, 3)
            'by Abhi on 12-Feb-2014 for caseid 3726 A29 EMDVAR (EMD VARIABLE DATA)
            Else
                'by Abhi on 03-Dec-2009 for Galileo Hotel HHL PNR uploading
                vOPTDATAID_String = Left(LStr, 3)
                If LastSec = "T5" Then
                    If i = 2 Then
                        'Header2
                        'by Abhi on 25-Feb-2010 for caseid 1227 Booking no or staff id is showing blank gds intray
                        vT50PNR_String = Format(fnMidDate(LStr, 44, 7), "dd/mmm/yyyy")
                        If IsDate(vT50PNR_String) = False Then
                            vT50PNR_String = Format(Date, "dd/mmm/yyyy")
                        End If
                        vT50DTL_String = Format(fnMidDate(LStr, 54, 7), "DD/MMM/YYYY")
                        If IsDate(vT50DTL_String) = False Then
                            vT50DTL_String = Format(Date, "DD/MMM/YYYY")
                        End If
                        'by Abhi on 25-Feb-2010 for caseid 1227 Booking no or staff id is showing blank gds intray
                        SQL = "INSERT INTO GalHeader2 (UPLOADNO,T50BPC,T50TPC,T50AAN,T50RCL,T50ORL,T50OCC,T50OAM,T50AGS," & _
                            "T50SBI,T50SIN,T50DTY,T50PNR,T50EHT,T50DTL,T50NMC) Values (" & Val(UploadNo) & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 4)) & "','" & Trim(fnMidValue(LStr, 5, 4)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 9, 9)) & "','" & Trim(fnMidValue(LStr, 18, 6)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 24, 6)) & "','" & Trim(fnMidValue(LStr, 30, 2)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 32, 1)) & "','" & Trim(fnMidValue(LStr, 33, 6)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 39, 1)) & "','" & Trim(fnMidValue(LStr, 40, 2)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 42, 2)) & "','" & vT50PNR_String & "'," & _
                            "" & Val(fnMidValue(LStr, 51, 3)) & ",'" & vT50DTL_String & "'," & _
                            "" & Val(fnMidValue(LStr, 61, 3)) & ")"
                        'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
                        LUFPNR_String = Trim(fnMidValue(LStr, 18, 6))
                        dbCompany.Execute SQL
                        
                    ElseIf i = 3 Then
                        'Header3
                        SQL = "INSERT INTO GalHeader3 (UPLOADNO,T50CUR,T50FAR,T50DML,T50CUR2,T501TX,T501TC,T502TX," & _
                            "T502TC,T503TX,T503TC,T504TX,T504TC,T505TX,T505TC,T50COM,T50RTE,T50ITC) Values " & _
                            "(" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & Val(fnMidValue(LStr, 4, 12)) & "," & _
                            "" & Val(fnMidValue(LStr, 16, 1)) & ",'" & Trim(fnMidValue(LStr, 17, 3)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 20, 8)) & "','" & Trim(fnMidValue(LStr, 28, 2)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 30, 8)) & "','" & Trim(fnMidValue(LStr, 38, 2)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 40, 8)) & "','" & Trim(fnMidValue(LStr, 48, 2)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 50, 8)) & "','" & Trim(fnMidValue(LStr, 58, 2)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 60, 8)) & "','" & Trim(fnMidValue(LStr, 68, 2)) & "'," & _
                            "" & Val(fnMidValue(LStr, 70, 8)) & ",'" & fnMidValue(LStr, 78, 4) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 82, 15)) & "')"
                        dbCompany.Execute SQL
                    ElseIf i = 4 Then
                        'Header4
                        SQL = "INSERT INTO GalHeader4 (UPLOADNO,T50IN1,T50IN2,T50IN3,T50IN4,T50IN5,T50IN6,T50IN7," & _
                            "T50IN8,T50IN9,T50IN10,T50IN11,T50IN12,T50IN13,T50IN14,T50IN15,T50IN16,T50IN17," & _
                            "T50PCC,T50ISO,T50DMI,T50DST,T50DPC,T50DSQ,T50DLN,T50SMI,T50SPC,T50SHP) VALUES " & _
                            "(" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 1)) & "','" & Trim(fnMidValue(LStr, 2, 1)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 3, 1)) & "','" & Trim(fnMidValue(LStr, 4, 1)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 5, 1)) & "'," & Val(fnMidValue(LStr, 6, 1)) & "," & _
                            "'" & Trim(fnMidValue(LStr, 7, 1)) & "','" & Trim(fnMidValue(LStr, 8, 1)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 9, 1)) & "','" & Trim(fnMidValue(LStr, 10, 1)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 11, 1)) & "','" & Trim(fnMidValue(LStr, 12, 1)) & "'," & _
                            "" & Val(fnMidValue(LStr, 13, 1)) & ",'" & Trim(fnMidValue(LStr, 14, 1)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 15, 1)) & "','" & Trim(fnMidValue(LStr, 16, 1)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 17, 1)) & "','" & Trim(fnMidValue(LStr, 18, 3)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 21, 3)) & "','" & Trim(fnMidValue(LStr, 24, 2)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 26, 1)) & "','" & Trim(fnMidValue(LStr, 27, 4)) & "'," & _
                            "" & Val(fnMidValue(LStr, 31, 5)) & ",'" & Trim(fnMidValue(LStr, 36, 6)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 42, 2)) & "','" & Trim(fnMidValue(LStr, 44, 4)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 48, 4)) & "')"
                        dbCompany.Execute SQL
                    
                    ElseIf i = 5 Then
                        'Header 5
                        SQL = "INSERT INTO GalHeader5 (UPLOADNO,T50CRN,T50CPN,T50PGN,T50FFN,T50ARN,T50WLN,T50SDN," & _
                            "T50FBN,T50EXC,T50PYN,T50PHN,T50ADN,T50MSN,T50RRN,T50AXN,T50LSN) VALUES (" & UploadNo & "," & _
                            "" & Val(fnMidValue(LStr, 1, 3)) & "," & Val(fnMidValue(LStr, 4, 3)) & "," & _
                            "" & Val(fnMidValue(LStr, 7, 3)) & "," & Val(fnMidValue(LStr, 10, 3)) & "," & _
                            "" & Val(fnMidValue(LStr, 13, 3)) & "," & Val(fnMidValue(LStr, 16, 3)) & "," & _
                            "" & Val(fnMidValue(LStr, 19, 3)) & "," & Val(fnMidValue(LStr, 22, 3)) & "," & _
                            "" & Val(fnMidValue(LStr, 25, 3)) & "," & Val(fnMidValue(LStr, 28, 3)) & "," & _
                            "" & Val(fnMidValue(LStr, 31, 3)) & "," & Val(fnMidValue(LStr, 34, 3)) & "," & _
                            "" & Val(fnMidValue(LStr, 37, 3)) & "," & Val(fnMidValue(LStr, 40, 3)) & "," & _
                            "" & Val(fnMidValue(LStr, 43, 3)) & "," & Val(fnMidValue(LStr, 46, 3)) & ")"
                        dbCompany.Execute SQL
                    End If
                ElseIf LastSec = "A02" Then
                    'by Abhi on 28-Oct-2014 for caseid 4665 Additional checking in Passenger type picking in Gal File
                    'SQL = "INSERT INTO GalPassenger1 (UPLOADNO,A02PNI,A02PNR,A02PT1,A02PTL,A02PA1,A02PA2," & _
                        "A02PG1,A02PG2,A02PS1,A02PSP,A02PC1,A02PCC,A02PP1,A02PPN,A02PD1,A02PDE,A02SCI," & _
                        "A02SCD,A02SCN) Values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 4, 33)) & "','" & Trim(fnMidValue(LStr, 37, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 39, 13)) & "','" & Trim(fnMidValue(LStr, 52, 2)) & "'," & _
                        "" & Val(fnMidValue(LStr, 54, 3)) & ",'" & Trim(fnMidValue(LStr, 57, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 59, 1)) & "','" & Trim(fnMidValue(LStr, 60, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 62, 1)) & "','" & Trim(fnMidValue(LStr, 63, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 65, 2)) & "','" & Trim(fnMidValue(LStr, 67, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 69, 33)) & "','" & Trim(fnMidValue(LStr, 102, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 104, 7)) & "','" & Trim(fnMidValue(LStr, 111, 3)) & "'," & _
                        "" & Val(fnMidValue(LStr, 114, 2)) & "," & Val(fnMidValue(LStr, 116, 11)) & ")"
                    vPos_Long = InStr(1, LStr, "NR:")
                    If vPos_Long = 0 Then 'Or vPos_Long - 5 < 0 Then
                        vA02PNRPTC_String = ""
                    Else
                        'by Abhi on 15-Dec-2015 for caseid 5878 Galileo -Passenger type is wrongly picking from NR identifier in A02 line for infant passenger
                        If IsDate(Mid(LStr, vPos_Long + 3, 2) & "/" & Mid(LStr, vPos_Long + 3 + 2, 3) & "/" & Mid(LStr, vPos_Long + 3 + 2 + 3, 2)) = True Then
                            vA02PNRPTC_String = "INF"
                        Else
                        'by Abhi on 15-Dec-2015 for caseid 5878 Galileo -Passenger type is wrongly picking from NR identifier in A02 line for infant passenger
                            vA02PNRPTC_String = Mid(LStr, vPos_Long + 3 + 2, 3)
                        'by Abhi on 15-Dec-2015 for caseid 5878 Galileo -Passenger type is wrongly picking from NR identifier in A02 line for infant passenger
                        End If
                        'by Abhi on 15-Dec-2015 for caseid 5878 Galileo -Passenger type is wrongly picking from NR identifier in A02 line for infant passenger
                    End If
                    
                    'by Abhi on 17-Jan-2018 for caseid 8156 Galileo Ticket date picking
                    vPos_Long = InStr(1, LStr, "NTD:")
                    If vPos_Long = 0 Then
                        vA02NTD_String = ""
                    Else
                        vA02NTD_String = Mid(LStr, vPos_Long + 4, 7)
                        If Len(Trim(vA02NTD_String)) = 7 Then
                            vA02NTD_String = fnMidDate(vA02NTD_String, 1, 7)
                        Else
                            vA02NTD_String = ""
                        End If
                    End If
                    'by Abhi on 17-Jan-2018 for caseid 8156 Galileo Ticket date picking
                    
                    'by Abhi on 17-Jan-2018 for caseid 8156 Galileo Ticket date picking
                    SQL = "INSERT INTO GalPassenger1 (UPLOADNO,A02PNI,A02PNR,A02PT1,A02PTL,A02PA1,A02PA2," & _
                        "A02PG1,A02PG2,A02PS1,A02PSP,A02PC1,A02PCC,A02PP1,A02PPN,A02PD1,A02PDE,A02SCI," & _
                        "A02SCD,A02SCN,A02SLNO,A02PNRPTC,A02NTD) Values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 4, 33)) & "','" & Trim(fnMidValue(LStr, 37, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 39, 13)) & "','" & Trim(fnMidValue(LStr, 52, 2)) & "'," & _
                        "" & Val(fnMidValue(LStr, 54, 3)) & ",'" & Trim(fnMidValue(LStr, 57, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 59, 1)) & "','" & Trim(fnMidValue(LStr, 60, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 62, 1)) & "','" & Trim(fnMidValue(LStr, 63, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 65, 2)) & "','" & Trim(fnMidValue(LStr, 67, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 69, 33)) & "','" & Trim(fnMidValue(LStr, 102, 2)) & "'," & _
                        "'" & Trim(fnMidValue(LStr, 104, 7)) & "','" & Trim(fnMidValue(LStr, 111, 3)) & "'," & _
                        "" & Val(fnMidValue(LStr, 114, 2)) & "," & Val(fnMidValue(LStr, 116, 11)) & "," & fGalPassengerSLNO & ",'" & SkipChars(vA02PNRPTC_String) & "'," & _
                        "'" & DateFormatBlankto1900(vA02NTD_String) & "')"
                    'by Abhi on 17-Jan-2018 for caseid 8156 Galileo Ticket date picking
                    'by Abhi on 28-Oct-2014 for caseid 4665 Additional checking in Passenger type picking in Gal File
                    dbCompany.Execute SQL
                ElseIf LastSec = "A07" Then
                
                    SQL = "UPDATE GalA07 Set A07ITT='" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                        "A07IT1='" & Trim(fnMidValue(LStr, 4, 8)) & "',A07IT1C='" & Trim(fnMidValue(LStr, 12, 2)) & "'," & _
                        "A07IT2='" & Trim(fnMidValue(LStr, 14, 8)) & "',A07IT2C='" & Trim(fnMidValue(LStr, 22, 2)) & "'," & _
                        "A07IT3='" & Trim(fnMidValue(LStr, 24, 8)) & "',A07IT3C='" & Trim(fnMidValue(LStr, 32, 2)) & "'," & _
                        "A07IT4='" & Trim(fnMidValue(LStr, 34, 8)) & "',A07IT4C='" & Trim(fnMidValue(LStr, 42, 2)) & "'," & _
                        "A07IT5='" & Trim(fnMidValue(LStr, 44, 8)) & "',A07IT5C='" & Trim(fnMidValue(LStr, 52, 2)) & "'," & _
                        "A07IT6='" & Trim(fnMidValue(LStr, 54, 8)) & "',A07IT6C='" & Trim(fnMidValue(LStr, 62, 2)) & "'," & _
                        "A07IT7='" & Trim(fnMidValue(LStr, 64, 8)) & "',A07IT7C='" & Trim(fnMidValue(LStr, 72, 2)) & "'," & _
                        "A07IT8='" & Trim(fnMidValue(LStr, 74, 8)) & "',A07IT8C='" & Trim(fnMidValue(LStr, 82, 2)) & "'," & _
                        "A07IT9='" & Trim(fnMidValue(LStr, 84, 8)) & "',A07IT9C='" & Trim(fnMidValue(LStr, 92, 2)) & "'," & _
                        "A07IT10='" & Trim(fnMidValue(LStr, 94, 8)) & "',A07IT10C='" & Trim(fnMidValue(LStr, 102, 2)) & "'," & _
                        "A07IT11='" & Trim(fnMidValue(LStr, 104, 8)) & "',A07IT11C='" & Trim(fnMidValue(LStr, 112, 2)) & "'," & _
                        "A07IT12='" & Trim(fnMidValue(LStr, 114, 8)) & "',A07IT12C='" & Trim(fnMidValue(LStr, 122, 2)) & "' " & _
                        "Where UPLOADNO=" & UploadNo & ""
                    
                    dbCompany.Execute SQL
                    
                ElseIf LastSec = "A09" Then
                    
                    dbCompany.Execute "UPDATE GalA09 Set A09L52='" & Trim(fnMidValue(LStr, 1, 61)) & "' Where UPLOADNO=" & Val(UploadNo) & ""
                'by Abhi on 03-Dec-2009 for Galileo Hotel HHL PNR uploading
                ElseIf LastSec = "A16HHL" Then
                    Select Case vOPTDATAID_String
                        Case "OD-"
                            vSQL_String = "" _
                                & "UPDATE GalA16HHL Set " _
                                & "A16ODI='" & Trim(fnMidValue(LStr, 1, 3)) & "', A16ODD='" & Trim(fnMidValue(LStr, 4, 220)) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM='" & LastSecNo_String & "'"
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                        Case "FF:"
                            vSQL_String = "" _
                                & "UPDATE GalA16HHL Set " _
                                & "A16FFI='" & Trim(fnMidValue(LStr, 1, 3)) & "', A16FIP='" & Trim(fnMidValue(LStr, 4, 2)) & "', " _
                                & "A16FTI='" & Trim(fnMidValue(LStr, 6, 1)) & "', A16FFD='" & Trim(fnMidValue(LStr, 7, 60)) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM='" & LastSecNo_String & "'"
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                        Case "CF:"
                            vSQL_String = "" _
                                & "UPDATE GalA16HHL Set " _
                                & "A16CFI='" & Trim(fnMidValue(LStr, 1, 3)) & "', A16CIP='" & Trim(fnMidValue(LStr, 4, 2)) & "', " _
                                & "A16CFD='" & Trim(fnMidValue(LStr, 6, 30)) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM='" & LastSecNo_String & "'"
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                        Case "DP:"
                            vSQL_String = "" _
                                & "UPDATE GalA16HHL Set " _
                                & "A16DPI='" & Trim(fnMidValue(LStr, 1, 3)) & "', A16DPP='" & Trim(fnMidValue(LStr, 4, 2)) & "', " _
                                & "A16DTI='" & Trim(fnMidValue(LStr, 6, 1)) & "', A16DAD='" & Trim(fnMidValue(LStr, 7, 12)) & "', " _
                                & "A16DCR='" & Trim(fnMidValue(LStr, 19, 3)) & "', A16DPD='" & Trim(fnMidValue(LStr, 22, 12)) & "', " _
                                & "A16DPA='" & Trim(fnMidValue(LStr, 34, 1)) & "', A16DPF='" & Trim(fnMidValue(LStr, 35, 60)) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM='" & LastSecNo_String & "'"
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                        Case "W-:"
                            'by Abhi on 23-Feb-2011 for caseid 1617 PenGDS Galileo Incorrect syntax near AEROPORTO
                            vSQL_String = "" _
                                & "UPDATE GalA16HHL Set " _
                                & "A16WAI='" & Trim(fnMidValue(LStr, 1, 3)) & "', A16WAP='" & Trim(fnMidValue(LStr, 4, 2)) & "', " _
                                & "A16WAD='" & SkipChars(Trim(fnMidValue(LStr, 6, 100))) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM='" & LastSecNo_String & "'"
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                    End Select
                'by Abhi on 13-May-2019 for caseid 7191 Galileo Car hire mapping
                ElseIf LastSec = "A16CCRCAR" Then
                    Select Case vOPTDATAID_String
                        Case "OD-" 'OPTIONAL DATA IDENTIFIER "OD-"
                            Splited = SplitWithLengths(LStr, 3, 220)
                            vSQL_String = "" _
                                & "UPDATE GalA16CCRCAR SET " _
                                & "A16ODN = '" & SkipChars(LStr) & "', A16ODI = '" & SkipChars(Trim(Splited(0))) & "', A16ODD ='" & SkipChars(Trim(Splited(1))) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM=" & Val(LastSecNo_String) & ""
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                        Case "O1-" 'OPTIONAL DATA CONTINUATION IDENTIFIER "O1-"
                            Splited = SplitWithLengths(LStr, 3, 250)
                            vSQL_String = "" _
                                & "UPDATE GalA16CCRCAR SET " _
                                & "A16O1I = '" & SkipChars(Trim(Splited(0))) & "', A16ODIO1I = '" & SkipChars(Trim(Splited(1))) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM=" & Val(LastSecNo_String) & ""
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                        Case "FF:" 'FREEFORM DATA IDENTIFIER "FF:"
                            Splited = SplitWithLengths(LStr, 3, 2, 1, 60)
                            vSQL_String = "" _
                                & "UPDATE GalA16CCRCAR SET " _
                                & "A16FFN = '" & SkipChars(LStr) & "', A16FFI = '" & SkipChars(Trim(Splited(0))) & "', A16FIP = " & Val(Trim(Splited(1))) & ", A16FTI = '" & SkipChars(Trim(Splited(2))) & "', A16FFD = '" & SkipChars(Trim(Splited(3))) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM=" & Val(LastSecNo_String) & ""
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                        Case "CF:" 'CONFIRMATION NUMBER IDENTIFIER "CF:"
                            Splited = SplitWithLengths(LStr, 3, 2, 30)
                            vSQL_String = "" _
                                & "UPDATE GalA16CCRCAR SET " _
                                & "A16CFN = '" & SkipChars(LStr) & "', A16CFI = '" & SkipChars(Trim(Splited(0))) & "', A16CIP = " & Val(Trim(Splited(1))) & ", A16CFD = '" & SkipChars(Trim(Splited(2))) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM=" & Val(LastSecNo_String) & ""
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                        Case "DP:" 'DUE/PAID IDENTIFIER "DP:"
                            Splited = SplitWithLengths(LStr, 3, 2, 1, 12, 3, 12, 1, 60)
                            vSQL_String = "" _
                                & "UPDATE GalA16CCRCAR SET " _
                                & "A16DPN = '" & SkipChars(LStr) & "', A16DPI = '" & SkipChars(Trim(Splited(0))) & "', A16DPP = " & Val(Trim(Splited(1))) & " , A16DTI ='" & SkipChars(Trim(Splited(2))) & "', A16DAD ='" & SkipChars(Trim(Splited(3))) & "', A16DCR = '" & SkipChars(Trim(Splited(4))) & "', A16DPD = '" & SkipChars(Trim(Splited(5))) & "', A16DPA = '" & SkipChars(Trim(Splited(6))) & "', A16DPF ='" & SkipChars(Trim(Splited(7))) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM=" & Val(LastSecNo_String) & ""
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                        Case "W-:" 'ADDRESS IDENTIFIER "W-:"
                            Splited = SplitWithLengths(LStr, 3, 2, 100)
                            vSQL_String = "" _
                                & "UPDATE GalA16CCRCAR SET " _
                                & "A16WAN = '" & SkipChars(LStr) & "', A16WAI = '" & SkipChars(Trim(Splited(0))) & "', A16WAP = " & Val(Trim(Splited(1))) & ", A16WAD = '" & SkipChars(Trim(Splited(2))) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM=" & Val(LastSecNo_String) & ""
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                        Case "VC-" 'CAR EVOUCHER CONFIRMATION NUMBER IDENTIFIER "VC-"
                            Splited = SplitWithLengths(LStr, 3, 2, 20, 25, 20, 30)
                            vSQL_String = "" _
                                & "UPDATE GalA16CCRCAR SET " _
                                & "A16VCN = '" & SkipChars(LStr) & "', A16VCI = '" & SkipChars(Trim(Splited(0))) & "', A16VIP = " & Val(Trim(Splited(1))) & ", A16EVV = '" & SkipChars(Trim(Splited(2))) & "', A16EVC = '" & SkipChars(Trim(Splited(3))) & "', A16EBN = '" & SkipChars(Trim(Splited(4))) & "', A16EAN = '" & SkipChars(Trim(Splited(5))) & "' " _
                                & "Where UPLOADNO=" & Val(UploadNo) & " AND A16NUM=" & Val(LastSecNo_String) & ""
                            vSQL_String = Replace(vSQL_String, Chr(13), "")
                            dbCompany.Execute vSQL_String
                    End Select
                'by Abhi on 13-May-2019 for caseid 7191 Galileo Car hire mapping
                'by Abhi on 12-Feb-2014 for caseid 3726 A29 EMDVAR (EMD VARIABLE DATA)
                ElseIf LastSec = "A29" Then
                    Splited = SplitWithLengths(LStr, 3, 240)
                    vSQL_String = "" _
                        & "UPDATE    GalA29 " _
                        & "SET              A29EDS = '" & SkipChars(LStr) & "', A29EDI = '" & SkipChars(Splited(0)) & "', A29EMDVAR = '" & SkipChars(Splited(1)) & "' " _
                        & "WHERE     (UPLOADNO = " & Val(UploadNo) & ") AND (A29PAX = '" & fA29PAX_String & "')"
                    vSQL_String = Replace(vSQL_String, Chr(13), "")
                    dbCompany.Execute vSQL_String
                'by Abhi on 12-Feb-2014 for caseid 3726 A29 EMDVAR (EMD VARIABLE DATA)
                End If
            End If
    Exit Sub
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'myErr:
'        Resume Next
End Sub
Public Function fnMidValue(str As String, Start As Integer, Ending As Integer) As String
    fnMidValue = Mid(str, Start, Ending)
End Function
Public Function fnMidDate(str As String, Start As Integer, Ending As Integer) As String
Dim DateStr As String
    DateStr = Mid(str, Start, Ending)
    fnMidDate = Mid(DateStr, 1, 2) & "/" & Mid(DateStr, 3, 3) & "/" & Mid(DateStr, 6, 2)
'    If IsDate(fnMidDate) = False Then
'        fnMidDate = mCurrentDate
'    End If
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'On Error GoTo Note
    Dim rsSelect As New ADODB.Recordset
    Me.Icon = FMain.Icon
    'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsSelect.Open "Select UPLOADDIRNAME,DESTDIRNAME From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsSelect.Open "Select UPLOADDIRNAME,DESTDIRNAME From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
    'If rsSelect.EOF = False Then
    '    txtDirName = IIf(IsNull(rsSelect!UPLOADDIRNAME), "", rsSelect!UPLOADDIRNAME)
    '    txtRelLocation = IIf(IsNull(rsSelect!DESTDIRNAME), "", rsSelect!DESTDIRNAME)
    'End If
    'rsSelect.Close
    'If rsSelect.State = 1 Then rsSelect.Close
    ''by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    ''rsSelect.Open "Select STATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    'rsSelect.Open "Select STATUS From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
    'chkGal.value = IIf((IsNull(rsSelect!Status) = True), 0, rsSelect!Status)
    txtDirName = txtGalilieoSource
    txtRelLocation = txtGalilieoDest
    chkGal.value = chkGalilieo
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    File1.Path = txtDirName.Text
    File1.Refresh
    File1.Pattern = FMain.txtGalileoExt
    File1.Refresh
    Exit Sub
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'Note:
'    If Me.Visible = True Then Unload Me
End Sub
Public Function fnChecktheDatabase(CheckStr As String) As Boolean
Dim rsSelect As New ADODB.Recordset
Dim RecLoc As String

RecLoc = fnMidValue(CheckStr, 99, 6)
'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
'rsSelect.Open "Select * From GalHeader2 Where T50RCL='" & RecLoc & "'", dbCompany, adOpenDynamic, adLockBatchOptimistic
rsSelect.Open "Select T50RCL From GalHeader2 Where T50RCL='" & RecLoc & "'", dbCompany, adOpenForwardOnly, adLockReadOnly
If rsSelect.EOF = False Then
    fnChecktheDatabase = True
Else
    fnChecktheDatabase = False
End If
rsSelect.Close
End Function

Private Function RecPEN_PENLINE(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Splited = SplitForce(Data, "\", 9)
    
    Temp1 = ExtractBetween(Data, "AC-", "\")
    'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
    'Coll.Add Temp1, "AC"
    Coll.Add Left(Temp1, 50), "AC"
    'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
    
    Temp1 = ExtractBetween(Data, "DES-", "\")
    Coll.Add Temp1, "DES"
    
    Temp1 = ExtractBetween(Data, "REFE-", "\")
    Coll.Add Temp1, "REFE"
    
    Temp1 = ExtractBetween(Data, "SELL-ADT-", "\")
    Coll.Add Temp1, "SELLADT"
    
    Temp1 = ExtractBetween(Data, "SELL-CHD-", "\")
    Coll.Add Temp1, "SELLCHD"
    
    Temp1 = ExtractBetween(Data, "SCHARGE-", "\")
    Coll.Add Temp1, "SCHARGE"
    
    Temp1 = ExtractBetween(Data, "DEPT-", "\")
    Coll.Add Temp1, "DEPT"
    
    Temp1 = ExtractBetween(Data, "BRANCH-", "\")
    Coll.Add Temp1, "BRANCH"
    
    Temp1 = ExtractBetween(Data, "CC-", "\")
    Coll.Add Temp1, "CUSTCC"
    
    Set RecPEN_PENLINE = Coll
End Function

Private Function InsertDataByFieldName(Data As Collection, TableName As String, LineIdent As String, UploadNo As Long) As Boolean
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
On Error GoTo PENErr
Dim ErrNumber As String
Dim ErrDescription As String
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
Dim rs As New ADODB.Recordset
Dim temp
Dim j As Integer
'by Abhi on 23-Oct-2010 for caseid 1516 PenGDS Amadeus slow reading
'Rs.Open "Select * from " + TableName, dbCompany, adOpenDynamic, adLockPessimistic
'by Abhi on 12-Nov-2010 for caseid 1546 PenGDS Optimistic concurrency check failed
'Rs.Open "Select * from " + TableName & " WHERE UpLoadNo=" & UploadNo, dbCompany, adOpenForwardOnly, adLockOptimistic
rs.Open "Select * from " + TableName & " WHERE UpLoadNo=" & UploadNo, dbCompany, adOpenForwardOnly, adLockPessimistic
rs.AddNew
rs.Fields("UpLoadNo") = UploadNo
rs.Fields("PENSEC") = LineIdent
For j = 2 To rs.Fields.Count - 1
    temp = rs.Fields(j).Name
    rs.Fields(temp) = Data(temp)
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    'If Len(Data(temp)) > rs.Fields(temp).DefinedSize Then
    '    ErrDetails_String = vbCrLf & "Table=" & TableName & ", Field=" & rs.Fields(temp).Name & ", FieldSize=" & rs.Fields(temp).DefinedSize & "<>" & Len(Data(temp)) & "," & vbCrLf & Data(temp) & vbCrLf
    'End If
    'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    If Len(Data(temp)) > rs.Fields(temp).DefinedSize Then
        ErrDetails_String = vbCrLf & "SEC=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(temp).Name & ", FieldSize=" & rs.Fields(temp).DefinedSize & "<>" & Len(Data(temp)) & "," & vbCrLf & Data(temp) & vbCrLf
    End If
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    DoEvents
Next
rs.Update
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
Exit Function
PENErr:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    If KeyExists(Data, temp) = False Then
        'ErrDetails_String = vbCrLf & "Table=" & TableName & ", Field=" & temp & ", Key Does NOT Exist!," & vbCrLf & Data(temp) & vbCrLf
        ErrDetails_String = vbCrLf & "SEC=" & LineIdent & ", Table=" & TableName & ", Field=" & temp & ", Field does not exist in Collection!," & vbCrLf
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    ElseIf Trim(ErrDetails_String) <> "" Then '-2147467259 Data provider or other service returned an E_FAIL status.
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    ElseIf Len(Data(temp)) > rs.Fields(temp).DefinedSize Then '-2147217887 Multiple-step operation generated errors. Check each status value.
        ErrDetails_String = vbCrLf & "SEC=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(temp).Name & ", FieldSize=" & rs.Fields(temp).DefinedSize & "<>" & Len(Data(temp)) & "," & vbCrLf & Data(temp) & vbCrLf
    Else
        ErrDetails_String = vbCrLf & "SEC=" & LineIdent & ", Table=" & TableName & ", Field=" & temp & "," & vbCrLf & Data(temp) & vbCrLf
    End If
    Err.Raise ErrNumber, , ErrDescription
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
End Function

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

Private Function ExtractBetween(Data, startText, endText, Optional AllifNoValue As Boolean = True)
Dim aa, ab, ac
Dim result
result = ""
aa = InStr(1, Data, startText, vbTextCompare)
ac = IIf(aa > 0, aa, 1)
ab = InStr(ac, Data, endText, vbTextCompare)
If AllifNoValue = True Then
    If ab = 0 Then ab = Len(Data) + 1
End If
If aa > 0 And ab > 0 And ab > aa Then

aa = aa + Len(startText)
'ab = ab + Len(endText)
result = Mid(Data, aa, (ab - aa))
Data = Replace(Data, startText & result & endText, "")
Data = Replace(Data, startText & result, "")
End If
ExtractBetween = result
End Function

'by Abhi on 11-Jan-2010 for caseid 1588 Galileo Penline AC BB
Private Function A14PEN_PENLINE(ByVal Data1) As Collection
Dim Splited
'Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
'splited = SplitForce(Data, "/", 9)

Temp1 = ExtractBetween(Data, "AC-", "/")
'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
'Coll.Add Temp1, "AC"
Coll.Add Left(Temp1, 50), "AC"
'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters

Temp1 = ExtractBetween(Data, "BB-", "/")
Coll.Add Temp1, "BB"

'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
Temp1 = ExtractBetween(Data, "DES-", "/")
Coll.Add Temp1, "DES"
Temp1 = ExtractBetween(Data, "REFE-", "/")
Coll.Add Temp1, "REFE"
Temp1 = ExtractBetween(Data, "CC-", "/")
Coll.Add Temp1, "CUSTCC"
Temp1 = ExtractBetween(Data, "BRANCH-", "/")
Coll.Add Temp1, "BRANCH"
Temp1 = ExtractBetween(Data, "MC-", "/")
Coll.Add Temp1, "MC"
Temp1 = ExtractBetween(Data, "INETREF-", "/")
Coll.Add Temp1, "INETREF"
'by Abhi on 21-Jan-2014 for caseid 3589 Penlines for Galileo
'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
Temp1 = ExtractBetween(Data, "TKTDEADLINE-", "/")
Coll.Add Temp1, "TKTDEADLINE"
'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
Temp1 = ExtractBetween(Data, "DEPT-", "/")
Coll.Add Temp1, "A14DEPT"
'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
Temp1 = ExtractBetween(Data, "DEPOSITAMT-", "/")
Coll.Add Val(Temp1), "A14DEPOSITAMT"
Temp1 = ExtractBetween(Data, "DEPOSITDUEDATE-", "/")
Coll.Add Left(Temp1, 9), "A14DEPOSITDUEDATE"
'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount

Data1 = Data
Set A14PEN_PENLINE = Coll
End Function

Private Function A14PEN_PENLINEPassenger(ByVal Data1, ByVal PNO As Long) As Collection
Dim Splited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1

'by Abhi on 05-Aug-2015 for caseid 5471 Galileo email address penline “//” instead of @
'by Abhi on 08-Mar-2018 for caseid 8331 Amadeus,Worldspan and Sabre - Pengds penline email replacement
'Data = Replace(Data, "//", "@")
'Data = Replace(Data, "?", "_")
Data = PenlineEmailReplacement(Data)
'by Abhi on 08-Mar-2018 for caseid 8331 Amadeus,Worldspan and Sabre - Pengds penline email replacement
'by Abhi on 05-Aug-2015 for caseid 5471 Galileo email address penline “//” instead of @

Temp1 = ExtractBetween(Data, "E" & PNO & "-", "/")
    'by Abhi on 11-Oct-2017 for caseid 7862 Passenger Email id validation in folder should also save multiple email id with coma separator
    'Coll.Add Temp1, "PEMAIL"
    Coll.Add Left(Temp1, 300), "PEMAIL"
    'by Abhi on 11-Oct-2017 for caseid 7862 Passenger Email id validation in folder should also save multiple email id with coma separator
Temp1 = ExtractBetween(Data, "T" & PNO & "-", "/")
    Coll.Add Temp1, "PTELE"
Coll.Add PNO, "PNO"
'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
Temp1 = ExtractBetween(Data, "DOB" & PNO & "-", "/")
    Coll.Add Trim(Temp1), "PDOB"
'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list

Data1 = Data
Set A14PEN_PENLINEPassenger = Coll
End Function

'by Abhi on 11-Feb-2014 for caseid 3721 Galileo EMD documents A29 and A30 section
Private Function SplitWithLengths(Data, ParamArray Lengths())
Dim Nos As Integer
Dim pos As Integer, j As Integer
Nos = UBound(Lengths)
Dim temp() As String
ReDim temp(Nos) As String
pos = 1

For j = 0 To Nos
    temp(j) = Mid(Data, pos, Lengths(j))
    pos = pos + Lengths(j)
    DoEvents
Next
SplitWithLengths = temp
End Function
'by Abhi on 11-Feb-2014 for caseid 3721 Galileo EMD documents A29 and A30 section

'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
Private Function A14PEN_PENAIRTKT(ByVal Data1) As Collection
Dim Splited
'Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
'splited = SplitForce(Data, "/", 9)

Temp1 = ExtractBetween(Data, "PAX-", "/")
    Coll.Add Left(Temp1, 150), "PAX"
Temp1 = ExtractBetween(Data, "TKT-", "/")
    Coll.Add Left(Temp1, 50), "TKT"
Temp1 = ExtractBetween(Data, "DATE-", "/")
    Coll.Add Left(Temp1, 9), "DATE"

Data1 = Data
Set A14PEN_PENAIRTKT = Coll
End Function
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
