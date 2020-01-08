VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FLegacy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Legacy"
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
Attribute VB_Name = "FLegacy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FNAME As String
Public RECIDQHeader As Boolean

Public Sub cmdDirUpload_Click()
'On Error GoTo Note
    Dim Count As Integer
    
    FMain.SendStatus FMain.SSTab1.TabCaption(5)
    File1.Path = txtDirName.Text
    File1.Pattern = FMain.LegacyExtText
    If txtDirName.Text = "" Or File1.ListCount = 0 Then FMain.stbUpload.Panels(1).Text = "Legacy Files Not Found...": Exit Sub
    'Uploading Each File
    Open (App.Path & "\_UploadingSQL_") For Random As #1
    Close #1
    FMain.cmdStop.Enabled = False
    DoEvents
    FMain.lblFName.Caption = "Uploading Legacy..."
    FMain.stbUpload.Panels(1).Text = "Reading..."
    FMain.stbUpload.Panels(2).Text = "Legacy"
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
            FMain.stbUpload.Panels(1).Text = "Opening..."
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
''    Dim rsSelect As New ADODB.Recordset
''    Me.Icon = FMain.Icon
''    rsSelect.Open "Select WSPUPLOADDIRNAME,WSPDESTDIRNAME From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
''    If rsSelect.EOF = False Then
''        txtDirName = IIf(IsNull(rsSelect!WSPUPLOADDIRNAME), "", rsSelect!WSPUPLOADDIRNAME)
''        txtRelLocation = IIf(IsNull(rsSelect!WSPDESTDIRNAME), "", rsSelect!WSPDESTDIRNAME)
''    End If
''    rsSelect.Close
''    If rsSelect.State = 1 Then rsSelect.Close
''    rsSelect.Open "Select WSPSTATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
''    chkGal.value = IIf((IsNull(rsSelect!WSPSTATUS) = True), 0, rsSelect!WSPSTATUS)
''    File1.Path = txtDirName.Text
''    File1.Refresh
''    File1.Pattern = FMain.txtWorldspanExt
''    File1.Refresh
''    Exit Sub
End Sub

Private Sub Text1_Change()
Caption = Len(Text1)
End Sub


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

Private Function SplitFirstNameAndInitial(Data)
Dim aa, ub, ab, ac, AD(1)
Dim temp

'by Abhi on 28-Oct-2010 for caseid 1531 PenGDS should take *CHD in passenger name as Child
'aa = Split(Data, ".")
'ub = UBound(aa)
'If ub > 0 Then
'    ab = aa(ub)
'    temp = Len(Data) - (Len(ab) + 1)
'    If temp > 0 Then
'        ac = Left(Data, temp)
'    Else
'        ac = ""
'    End If
'Else
'temp = FindInitialAndName(aa(0))
temp = FindInitialAndName(Data)
    ac = temp(0)
    ab = temp(1)
'by Abhi on 28-Oct-2010 for caseid 1531 PenGDS should take *CHD in passenger name as Child
'End If
AD(0) = ac
AD(1) = ab
SplitFirstNameAndInitial = AD
End Function

Private Function FindInitialAndName(mData)
On Error GoTo errPara
Dim aa, Data
Data = mData
Dim retn(1) As String

Data = Replace(Data, "*CHD", "")

aa = UCase(Right(Data, 2))
If (aa = "MR") Then
    retn(0) = Mid(Data, 1, Len(Data) - 2)
    retn(1) = "MR"
End If

aa = UCase(Right(Data, 3))
If (aa = "MRS") Then
    retn(0) = Mid(Data, 1, Len(Data) - 3)
    retn(1) = "MRS"
End If

aa = UCase(Right(Data, 4))
If (aa = "MISS") Then
    retn(0) = Mid(Data, 1, Len(Data) - 4)
    retn(1) = "MISS"
End If

aa = UCase(Right(Data, 4))
If (aa = "MSTR") Then
    retn(0) = Mid(Data, 1, Len(Data) - 4)
    retn(1) = "MSTR"
End If

aa = UCase(Right(Data, 4))
If (aa = "PROF") Then
    retn(0) = Mid(Data, 1, Len(Data) - 4)
    retn(1) = "PROF"
End If



If retn(0) = "" Then
    retn(0) = Data
    retn(1) = ""
End If

FindInitialAndName = retn
Exit Function
errPara:

If retn(0) = "" Then
    retn(0) = Data
    retn(1) = ""
End If
End Function

Private Function GetAirSegDate(ByVal pDATETIME24_String As String) As String
Dim vDATE_String As String
    vDATE_String = Mid(pDATETIME24_String, 1, 10)
    If IsDate(DateFormat(vDATE_String)) = True Then
        'commented by Abhi on 31-Jul-2012 for caseid 2401 Travcom data transfer rail
        'vDATE_String = Format(CDate(vDATE_String), "DD") & Format(CDate(vDATE_String), "MMM")
        'vDATE_String = UCase(vDATE_String)
    Else
        vDATE_String = ""
    End If
GetAirSegDate = vDATE_String
End Function

Private Function GetAirSegTime(ByVal pDATETIME24_String As String) As String
Dim vTIME_String As String
    vTIME_String = Mid(pDATETIME24_String, 12)
    vTIME_String = Left(Format(vTIME_String, "HHMMAMPM"), 5)
GetAirSegTime = vTIME_String
End Function

Private Function GetBDATE(ByVal pData_String As String) As String
Dim vBDATE_String As String
    If IsDate(DateFormat(pData_String)) = True Then
        vBDATE_String = Format(CDate(pData_String), "DD") & Format(CDate(pData_String), "MMM") & Format(CDate(pData_String), "YY")
        vBDATE_String = UCase(vBDATE_String)
    End If
GetBDATE = vBDATE_String
End Function


Public Function AddFile(FileName As String, Optional fTitle = "")
'On Error GoTo Note
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
On Error GoTo PENErr
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
Dim UploadNo As Long
Dim FileObj As New FileSystemObject

Dim rsSelectExcelSheet1 As New ADODB.Recordset
Dim rsSelectExcelSheet2 As New ADODB.Recordset
Dim rsSelectExcelSheet3 As New ADODB.Recordset

Dim vInvoiceID_String As String
Dim vSheet2FormOfPayment_String As String
Dim vInvoiceDetailID_String As String
Dim vSheet2PassengerName_String As String
Dim vSheet2EticketFlag_String As String
Dim vSheet3DepartureDate_String As String
Dim vSheet3ArrivalDate_String As String
Dim vSheet1BookingDate_String As String
Dim vSheet1Department_String As String
Dim vSheet2Ref63_String As String
Dim vSheet2ProductGroupCode_String As String

Dim SubSplited, SubSplited2

Dim vCUSTCC_String As String

Dim vPassengerID_Integer As Integer

Dim vWSPPNRADD_Type As WSPPNRADD_Type
Dim vWSPAGNTSINE_Type As WSPAGNTSINE_Type
Dim vWSPCLNTACNO_Type As WSPCLNTACNO_Type
Dim vWSPFORMOFPYMNT_Type As WSPFORMOFPYMNT_Type
Dim vWSPTCARRIER_Type As WSPTCARRIER_Type
Dim vWSPTKTSEG_Type As WSPTKTSEG_Type
Dim vWSPPNAME_Type As WSPPNAME_Type
Dim vWSPAIRFARE_Type As WSPAIRFARE_Type
Dim vWSPN_REF_Type As WSPN_REF_Type
Dim vWSPPENLINE_Type As WSPPENLINE_Type
Dim vWSPHTLSEG_Type As WSPHTLSEG_Type

'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
Dim vSQL_String As String
'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
Dim vRecordNumber_Long As Long
'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
'by Abhi on 18-Aug-2014 for caseid 4439 Change the mapping  "InvoiceNo" instead of "InvoiceID" for Legacy Loading (Travcom Mapping)
Dim vInvoiceNo_String As String
'by Abhi on 18-Aug-2014 for caseid 4439 Change the mapping  "InvoiceNo" instead of "InvoiceID" for Legacy Loading (Travcom Mapping)
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
    LUFPNR_String = fTitle
    'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    'Sleep 2000
    'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    NoofPermissionDenied = 0
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
    vRecordNumber_Long = 1
    'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
    If ConnectExcelFile(FileName) = True Then
        'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
        dbCompany.BeginTrans
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        PENErr_BeginTrans = True
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the

        If rsSelectExcelSheet1.State = 1 Then rsSelectExcelSheet1.Close
        rsSelectExcelSheet1.Open "Select  [Sheet1$].* From [Sheet1$]", CnExcel, adOpenForwardOnly, adLockReadOnly
        'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
        FMain.prgUpload.Max = rsSelectExcelSheet1.RecordCount
        'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
        Do While rsSelectExcelSheet1.EOF = False
            vInvoiceID_String = RemoveQuotes(rsSelectExcelSheet1.Fields("InvoiceID"))
            vSheet1Department_String = RemoveQuotes(rsSelectExcelSheet1.Fields("Department"))
            
            'by Abhi on 18-Aug-2014 for caseid 4439 Change the mapping  "InvoiceNo" instead of "InvoiceID" for Legacy Loading (Travcom Mapping)
            vInvoiceNo_String = RemoveQuotes(rsSelectExcelSheet1.Fields("InvoiceNo"))
            'by Abhi on 18-Aug-2014 for caseid 4439 Change the mapping  "InvoiceNo" instead of "InvoiceID" for Legacy Loading (Travcom Mapping)
            
            'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
            GIT_CUSTOMERUSERCODE_String = ""
            'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
            'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
            GIT_PENWAIT_String = "N"
            'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
            
            FMain.stbUpload.Panels(1).Text = "Reading... " & vInvoiceID_String
            FMain.stbUpload.Refresh
            DoEvents
            'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
            'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
            'dbCompany.BeginTrans
            'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
            'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
            UploadNo = FirstFreeNumber("UPLOADNO")
            
            'Header WSPPNRADD
            vWSPPNRADD_Type.UpLoadNo_Long = UploadNo
            vWSPPNRADD_Type.RecID_String = ""
            vWSPPNRADD_Type.INTLVL_String = ""
            vWSPPNRADD_Type.PNRADD_String = RemoveQuotes(rsSelectExcelSheet1.Fields("Locator"))
            vWSPPNRADD_Type.FINVNO_String = ""
            vWSPPNRADD_Type.LINVNO_String = ""
            vWSPPNRADD_Type.ITNRYCHNGE_String = ""
            vWSPPNRADD_Type.DLCI_String = ""
            vWSPPNRADD_Type.TLCI_String = ""
            vWSPPNRADD_Type.FNAME_String = fTitle
            vWSPPNRADD_Type.LUPDATE_String = Now
            vWSPPNRADD_Type.GDSAutoFailed_Integer = 0
            Call WSPPNRADD_Add(vWSPPNRADD_Type)
            
            'Header WSPAGNTSINE
            'vSheet1BookingDate_String = RemoveQuotes(rsSelectExcelSheet1.Fields("BookingDate"))
            vSheet1BookingDate_String = RemoveQuotes(rsSelectExcelSheet1.Fields("InvoiceDate"))
            
            vWSPAGNTSINE_Type.UpLoadNo_Long = UploadNo
            vWSPAGNTSINE_Type.RecID_String = ""
            vWSPAGNTSINE_Type.BSID_String = ""
            vWSPAGNTSINE_Type.BIATA_String = RemoveQuotes(rsSelectExcelSheet1.Fields("IATANumber"))
            vWSPAGNTSINE_Type.BDATE_String = GetBDATE(vSheet1BookingDate_String)
            vWSPAGNTSINE_Type.BTIME_String = ""
            vWSPAGNTSINE_Type.BAGENT_String = RemoveQuotes(rsSelectExcelSheet1.Fields("BookingAgentNumber"))
            vWSPAGNTSINE_Type.TIATA_String = RemoveQuotes(rsSelectExcelSheet1.Fields("IATANumber"))
            vWSPAGNTSINE_Type.TDATE_String = DateFormat(vSheet1BookingDate_String)
            vWSPAGNTSINE_Type.TAGENT_String = RemoveQuotes(rsSelectExcelSheet1.Fields("TicketingAgentNumber"))
            Call WSPAGNTSINE_Add(vWSPAGNTSINE_Type)
            
            'Header WSPCLNTACNO
            vWSPCLNTACNO_Type.UpLoadNo_Long = UploadNo
            vWSPCLNTACNO_Type.RecID_String = ""
            vWSPCLNTACNO_Type.CLNTACNO_String = RemoveQuotes(rsSelectExcelSheet1.Fields("PENClientCode"))
            Call WSPCLNTACNO_Add(vWSPCLNTACNO_Type)
            'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
            If Trim(vWSPCLNTACNO_Type.CLNTACNO_String) <> "" Then
                GIT_CUSTOMERUSERCODE_String = vWSPCLNTACNO_Type.CLNTACNO_String
            End If
            'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
            
            'Header
            'by Abhi on 27-Nov-2012 for caseid 2479 Travcom data transfer Multple Rail tickets under single invoice id is picking only one ticket and segments
            vPassengerID_Integer = 1
            If rsSelectExcelSheet2.State = 1 Then rsSelectExcelSheet2.Close
            'by Abhi on 31-Jul-2012 for caseid 2401 Travcom data transfer rail
            'rsSelectExcelSheet2.Open "Select  [Sheet2$].* From [Sheet2$] WHERE [Sheet2$].InvoiceID='" & Val(vInvoiceID_String) & "'", CnExcel, adOpenForwardOnly, adLockReadOnly
            rsSelectExcelSheet2.Open "Select  [Sheet2$].* From [Sheet2$] WHERE [Sheet2$].InvoiceID='" & Val(vInvoiceID_String) & "' AND [Sheet2$].TransactionType='1'", CnExcel, adOpenForwardOnly, adLockReadOnly
            'by Abhi on 27-Nov-2012 for caseid 2479 Travcom data transfer Multple Rail tickets under single invoice id is picking only one ticket and segments
            'If rsSelectExcelSheet2.EOF = False Then
            Do While rsSelectExcelSheet2.EOF = False
                'Header WSPFORMOFPYMNT
                vSheet2FormOfPayment_String = RemoveQuotes(rsSelectExcelSheet2.Fields("FormOfPayment"))
                vWSPFORMOFPYMNT_Type.UpLoadNo_Long = UploadNo
                vWSPFORMOFPYMNT_Type.RecID_String = ""
                'If Val(vSheet2FormOfPayment_String) = 0 Or Val(vSheet2FormOfPayment_String) = 1 Then
                '    vWSPFORMOFPYMNT_Type.PAYCODE_String = "AR"
                'ElseIf Val(vSheet2FormOfPayment_String) = 2 Then
                '    vWSPFORMOFPYMNT_Type.PAYCODE_String = "CC"
                'Else
                '    vWSPFORMOFPYMNT_Type.PAYCODE_String = ""
                'End If
                If Val(vSheet2FormOfPayment_String) = 1 Then
                    vWSPFORMOFPYMNT_Type.PAYCODE_String = "CC"
                Else
                    vWSPFORMOFPYMNT_Type.PAYCODE_String = "AR"
                End If
                
                vWSPFORMOFPYMNT_Type.PAYNAME_String = ""
                vWSPFORMOFPYMNT_Type.FFDATA1_String = ""
                vWSPFORMOFPYMNT_Type.FFDATA2_String = ""
                vWSPFORMOFPYMNT_Type.FFDATA3_String = ""
                vWSPFORMOFPYMNT_Type.FFDATA4_String = ""
                vWSPFORMOFPYMNT_Type.STAXCODE_String = ""
                vWSPFORMOFPYMNT_Type.ITAXCODE1_String = ""
                vWSPFORMOFPYMNT_Type.ITAXCODE2_String = ""
                vWSPFORMOFPYMNT_Type.CCDETAILS_String = ""
                Call WSPFORMOFPYMNT_Add(vWSPFORMOFPYMNT_Type)
            
                'Header WSPTCARRIER
                vWSPTCARRIER_Type.UpLoadNo_Long = UploadNo
                vWSPTCARRIER_Type.RecID_String = ""
                vWSPTCARRIER_Type.VAIRCODE_String = ""
                vWSPTCARRIER_Type.VAIRNO_String = RemoveQuotes(rsSelectExcelSheet2.Fields("ValidatingCarrier"))
                vWSPTCARRIER_Type.TKTIND_String = ""
                vWSPTCARRIER_Type.INTIND_String = ""
                vWSPTCARRIER_Type.DESTCDE_String = ""
                vWSPTCARRIER_Type.PTCODE_String = ""
                Call WSPTCARRIER_Add(vWSPTCARRIER_Type)
            
                'WSPPNAME
                vSheet2PassengerName_String = RemoveQuotes(rsSelectExcelSheet2.Fields("PassengerName"))
                SubSplited = SplitForce(vSheet2PassengerName_String, "/", 2)
                SubSplited2 = SplitFirstNameAndInitial(SubSplited(1))
                
                vSheet2EticketFlag_String = RemoveQuotes(rsSelectExcelSheet2.Fields("EticketFlag"))
                
                vWSPPNAME_Type.UpLoadNo_Long = UploadNo
                vWSPPNAME_Type.RecID_String = ""
                vWSPPNAME_Type.SURNAME_String = SubSplited(0)
                vWSPPNAME_Type.FRSTNAME_String = Replace(SubSplited2(0), "*CHD", "")
                vWSPPNAME_Type.PTITLE_String = Replace(SubSplited2(1), "*CHD", "")
                vWSPPNAME_Type.pType_String = ""
                vWSPPNAME_Type.CUSTNO_String = ""
                vWSPPNAME_Type.CUSTCMNTS_String = ""
                vWSPPNAME_Type.DOCNO_String = RemoveQuotes(rsSelectExcelSheet2.Fields("TicketNumber"))
                vWSPPNAME_Type.ISSUEDATE_String = ""
                vWSPPNAME_Type.INVNO_String = ""
                vWSPPNAME_Type.SETTMNTCDENO_String = ""
                'by Abhi on 27-Nov-2012 for caseid 2479 Travcom data transfer Multple Rail tickets under single invoice id is picking only one ticket and segments
                'vWSPPNAME_Type.PassengerID_Integer = 1
                vWSPPNAME_Type.PassengerID_Integer = vPassengerID_Integer
                If Trim(vSheet2EticketFlag_String) = "True" Then
                    vWSPPNAME_Type.ETKTIND_String = "E-"
                Else
                    vWSPPNAME_Type.ETKTIND_String = "N"
                End If
                vWSPPNAME_Type.SURNAMEFRSTNAMENO_String = ""
                Call WSPPNAME_Add(vWSPPNAME_Type)
            'by Abhi on 27-Nov-2012 for caseid 2479 Travcom data transfer Multple Rail tickets under single invoice id is picking only one ticket and segments
            'End If
                vPassengerID_Integer = vPassengerID_Integer + 1
                rsSelectExcelSheet2.MoveNext
            Loop
            
            'WSPTKTSEG
            vInvoiceDetailID_String = ""
            'by Abhi on 27-Nov-2012 for caseid 2479 Travcom data transfer Multple Rail tickets under single invoice id is picking only one ticket and segments
            vPassengerID_Integer = 1
            If rsSelectExcelSheet2.State = 1 Then rsSelectExcelSheet2.Close
            'by Abhi on 31-Jul-2012 for caseid 2401 Travcom data transfer rail
            'rsSelectExcelSheet2.Open "Select  [Sheet2$].* From [Sheet2$] WHERE [Sheet2$].InvoiceID='" & Val(vInvoiceID_String) & "' AND [Sheet2$].ProductGroupCode='1'", CnExcel, adOpenForwardOnly, adLockReadOnly
            'rsSelectExcelSheet2.Open "Select  [Sheet2$].* From [Sheet2$] WHERE [Sheet2$].InvoiceID='" & Val(vInvoiceID_String) & "' AND ([Sheet2$].ProductGroupCode='1' OR [Sheet2$].ProductGroupCode='2')", CnExcel, adOpenForwardOnly, adLockReadOnly
            rsSelectExcelSheet2.Open "Select  [Sheet2$].* From [Sheet2$] WHERE [Sheet2$].InvoiceID='" & Val(vInvoiceID_String) & "' AND ([Sheet2$].ProductGroupCode='1' OR [Sheet2$].ProductGroupCode='2') AND [Sheet2$].TransactionType='1'", CnExcel, adOpenForwardOnly, adLockReadOnly
            'by Abhi on 27-Nov-2012 for caseid 2479 Travcom data transfer Multple Rail tickets under single invoice id is picking only one ticket and segments
            'If rsSelectExcelSheet2.EOF = False Then
            Do While rsSelectExcelSheet2.EOF = False
                vInvoiceDetailID_String = RemoveQuotes(rsSelectExcelSheet2.Fields("InvoiceDetailID"))
                'by Abhi on 31-Jul-2012 for caseid 2401 Travcom data transfer rail
                vSheet2ProductGroupCode_String = RemoveQuotes(rsSelectExcelSheet2.Fields("ProductGroupCode"))
                
                'by Abhi on 30-Jul-2012 for caseid 2401 Travcom data transfer Hotel rate
                'WSPAIRFARE
                vWSPAIRFARE_Type.UpLoadNo_Long = UploadNo
                vWSPAIRFARE_Type.RecID_String = ""
                vWSPAIRFARE_Type.CURCODEFARE_String = RemoveQuotes(rsSelectExcelSheet2.Fields("CurrencyCode"))
                vWSPAIRFARE_Type.FARE_String = MoneyFormat(Val(RemoveQuotes(rsSelectExcelSheet2.Fields("SellingFare"))))
                vWSPAIRFARE_Type.TAX1CODEID_String = ""
                vWSPAIRFARE_Type.TAX1AMT_String = MoneyFormat(Val(RemoveQuotes(rsSelectExcelSheet2.Fields("Tax1"))))
                vWSPAIRFARE_Type.TAX2CODEID_String = ""
                vWSPAIRFARE_Type.TAX2AMT_String = MoneyFormat(Val(RemoveQuotes(rsSelectExcelSheet2.Fields("Tax2"))))
                vWSPAIRFARE_Type.TAX3CODEID_String = ""
                vWSPAIRFARE_Type.TAX3AMT_String = MoneyFormat(Val(RemoveQuotes(rsSelectExcelSheet2.Fields("Tax4"))))
                vWSPAIRFARE_Type.CURRCDETFARE_String = ""
                vWSPAIRFARE_Type.TATFAMT_String = ""
                vWSPAIRFARE_Type.CURRCDEEFARE_String = ""
                vWSPAIRFARE_Type.EFAMT_String = ""
                vWSPAIRFARE_Type.INVAMT_String = ""
                'by Abhi on 27-Nov-2012 for caseid 2479 Travcom data transfer Multple Rail tickets under single invoice id is picking only one ticket and segments
                'vWSPAIRFARE_Type.PassengerID_Integer = 1
                vWSPAIRFARE_Type.PassengerID_Integer = vPassengerID_Integer
                vWSPAIRFARE_Type.SLNO_Integer = 1
                Call WSPAIRFARE_Add(vWSPAIRFARE_Type)
            
                If rsSelectExcelSheet3.State = 1 Then rsSelectExcelSheet3.Close
                rsSelectExcelSheet3.Open "Select  [Sheet3$].* From [Sheet3$] WHERE [Sheet3$].InvoiceID='" & Val(vInvoiceID_String) & "' AND [Sheet3$].InvoiceDetailID='" & Val(vInvoiceDetailID_String) & "'", CnExcel, adOpenForwardOnly, adLockReadOnly
                Do While rsSelectExcelSheet3.EOF = False
                    vSheet3DepartureDate_String = RemoveQuotes(rsSelectExcelSheet3.Fields("DepartureDate"))
                    vSheet3ArrivalDate_String = RemoveQuotes(rsSelectExcelSheet3.Fields("ArrivalDate"))
                    
                    vWSPTKTSEG_Type.UpLoadNo_Long = UploadNo
                    vWSPTKTSEG_Type.RecID_String = ""
                    vWSPTKTSEG_Type.AIRSEGCODE_String = ""
                    vWSPTKTSEG_Type.ITNRYSEGNO_String = ""
                    vWSPTKTSEG_Type.NOINTSTOP_String = RemoveQuotes(rsSelectExcelSheet3.Fields("StopOver"))
                    vWSPTKTSEG_Type.SHDESIGNINDI_String = ""
                    'by Abhi on 31-Jul-2012 for caseid 2401 Travcom data transfer rail
                    If vSheet2ProductGroupCode_String = "1" Then
                        vWSPTKTSEG_Type.AIRCODE_String = RemoveQuotes(rsSelectExcelSheet3.Fields("CarrierCode"))
                    Else
                        vWSPTKTSEG_Type.AIRCODE_String = RemoveQuotes(rsSelectExcelSheet2.Fields("VendorNumber"))
                    End If
                    vWSPTKTSEG_Type.FLTNO_String = RemoveQuotes(rsSelectExcelSheet3.Fields("FlightNumber"))
                    vWSPTKTSEG_Type.Class_String = RemoveQuotes(rsSelectExcelSheet3.Fields("ClassOfService"))
                    vWSPTKTSEG_Type.ORGAIRCODE_String = RemoveQuotes(rsSelectExcelSheet3.Fields("DepartureCityCode"))
                    vWSPTKTSEG_Type.DEPDATE_String = GetAirSegDate(vSheet3DepartureDate_String)
                    vWSPTKTSEG_Type.DEPTIME_String = GetAirSegTime(vSheet3DepartureDate_String)
                    vWSPTKTSEG_Type.DESTAIRCODE_String = RemoveQuotes(rsSelectExcelSheet3.Fields("ArrivalCityCode"))
                    vWSPTKTSEG_Type.ARRDATE_String = GetAirSegDate(vSheet3ArrivalDate_String)
                    vWSPTKTSEG_Type.ARRTIME_String = GetAirSegTime(vSheet3ArrivalDate_String)
                    vWSPTKTSEG_Type.BDATE_String = ""
                    vWSPTKTSEG_Type.ADATE_String = ""
                    vWSPTKTSEG_Type.BAGG_String = ""
                    vWSPTKTSEG_Type.STATUS_String = ""
                    vWSPTKTSEG_Type.MEALSERVCODE_String = ""
                    vWSPTKTSEG_Type.SEGSTOPCODE_String = ""
                    vWSPTKTSEG_Type.EQUIPTYPE_String = ""
                    If Trim(vSheet2ProductGroupCode_String) = "2" Then
                        vWSPTKTSEG_Type.EQUIPTYPE_String = "TRN"
                    End If
                    vWSPTKTSEG_Type.FBASISCODE_String = RemoveQuotes(rsSelectExcelSheet3.Fields("FareBasis"))
                    vWSPTKTSEG_Type.SEGMILEAGE_String = ""
                    vWSPTKTSEG_Type.INTSTOP_String = ""
                    vWSPTKTSEG_Type.FLTTIME_String = ""
                    vWSPTKTSEG_Type.SEGOVRRDIND_String = ""
                    vWSPTKTSEG_Type.DEPTRMNLCODE_String = ""
                    vWSPTKTSEG_Type.ARRTRMNLCODE_String = ""
                    vWSPTKTSEG_Type.BPAYAMT_String = ""
                    Call WSPTKTSEG_Add(vWSPTKTSEG_Type)
                    
                    rsSelectExcelSheet3.MoveNext
                Loop
            'by Abhi on 27-Nov-2012 for caseid 2479 Travcom data transfer Multple Rail tickets under single invoice id is picking only one ticket and segments
            'End If
                vPassengerID_Integer = vPassengerID_Integer + 1
                rsSelectExcelSheet2.MoveNext
            Loop
            
            
            'WSPN_REF
            vWSPN_REF_Type.UpLoadNo_Long = UploadNo
            vWSPN_REF_Type.RecID_String = ""
            'by Abhi on 18-Aug-2014 for caseid 4439 Change the mapping  "InvoiceNo" instead of "InvoiceID" for Legacy Loading (Travcom Mapping)
            'vWSPN_REF_Type.REF_String = vInvoiceID_String
            vWSPN_REF_Type.REF_String = vInvoiceNo_String
            'by Abhi on 18-Aug-2014 for caseid 4439 Change the mapping  "InvoiceNo" instead of "InvoiceID" for Legacy Loading (Travcom Mapping)
            Call WSPN_REF_Add(vWSPN_REF_Type)
            
            'If vInvoiceID_String = "1275033" Then
            If vInvoiceID_String = "1274376" Then
                Debug.Print vInvoiceID_String
            End If
            
            'WSPPENLINE
            If rsSelectExcelSheet2.State = 1 Then rsSelectExcelSheet2.Close
            'by Abhi on 31-Jul-2012 for caseid 2401 Travcom data transfer rail
            'rsSelectExcelSheet2.Open "Select  [Sheet2$].* From [Sheet2$] WHERE [Sheet2$].InvoiceID='" & Val(vInvoiceID_String) & "' AND [Sheet2$].ProductGroupCode NOT IN('1','2','3','4')", CnExcel, adOpenForwardOnly, adLockReadOnly
            'by Abhi on 13-Aug-2014 for caseid 4425 Client Reference details not loading for Customer Card Transactions
            'rsSelectExcelSheet2.Open "Select  [Sheet2$].* From [Sheet2$] WHERE [Sheet2$].InvoiceID='" & Val(vInvoiceID_String) & "' AND [Sheet2$].ProductGroupCode NOT IN('1','2','3','4') AND [Sheet2$].TransactionType='1'", CnExcel, adOpenForwardOnly, adLockReadOnly
            'Do While rsSelectExcelSheet2.EOF = False
            '    vSheet2FormOfPayment_String = RemoveQuotes(rsSelectExcelSheet2.Fields("FormOfPayment"))
            '
            '    'PENO
            '    vWSPPENLINE_Type.UpLoadNo_Long = UploadNo
            '    vWSPPENLINE_Type.RecID_String = "PENO"
            '    vWSPPENLINE_Type.AC_String = ""
            '    vWSPPENLINE_Type.DES_String = ""
            '    vWSPPENLINE_Type.REFE_String = ""
            '    vWSPPENLINE_Type.SELLADT_String = ""
            '    vWSPPENLINE_Type.SELLCHD_String = ""
            '    vWSPPENLINE_Type.SCHARGE_String = ""
            '    vWSPPENLINE_Type.DEPT_String = ""
            '    vWSPPENLINE_Type.BRANCH_String = ""
            '    vWSPPENLINE_Type.CUSTCC_String = ""
            '    vWSPPENLINE_Type.PEMAIL_String = ""
            '    vWSPPENLINE_Type.PTELE_String = ""
            '    vWSPPENLINE_Type.PNO_Long = 0
            '    vWSPPENLINE_Type.FARE_Currency = 0
            '    vWSPPENLINE_Type.TAXES_Currency = 0
            '    vWSPPENLINE_Type.MARKUP_Currency = 0
            '    vWSPPENLINE_Type.SURNAMEFRSTNAMENO_String = ""
            '    vWSPPENLINE_Type.PENFARESELL_Currency = 0
            '    vWSPPENLINE_Type.PENFAREPASSTYPE_String = ""
            '    vWSPPENLINE_Type.DeliAdd_String = ""
            '    vWSPPENLINE_Type.MC_String = ""
            '    vWSPPENLINE_Type.BB_String = ""
            '    vWSPPENLINE_Type.PENFARESUPPID_String = ""
            '    vWSPPENLINE_Type.PENFAREAIRID_String = ""
            '    vWSPPENLINE_Type.PENLINKPNR_String = ""
            '    'by Abhi on 31-Jul-2012 for caseid 2401 Travcom data transfer hotel only etc
            '    'vWSPPENLINE_Type.PENOPRDID_String = RemoveQuotes(rsSelectExcelSheet2.Fields("ProductCode"))
            '    vWSPPENLINE_Type.PENOPRDID_String = RemoveQuotes(rsSelectExcelSheet2.Fields("ProductGroupCode"))
            '    vWSPPENLINE_Type.PENOQTY_Integer = 1
            '    vWSPPENLINE_Type.PENORATE_Currency = Val(RemoveQuotes(rsSelectExcelSheet2.Fields("NetFare")))
            '    vWSPPENLINE_Type.PENOSELL_Currency = Val(RemoveQuotes(rsSelectExcelSheet2.Fields("SellingFare")))
            '    vWSPPENLINE_Type.PENOSUPPID_String = RemoveQuotes(rsSelectExcelSheet2.Fields("VendorNumber"))
            '    vWSPPENLINE_Type.PENATOLTYPE_String = ""
            '    vWSPPENLINE_Type.INETREF_String = ""
            '    If Val(vSheet2FormOfPayment_String) = 1 Then
            '        vWSPPENLINE_Type.PENOPAYMETHODID_String = "CP"
            '    Else
            '        vWSPPENLINE_Type.PENOPAYMETHODID_String = ""
            '    End If
            '    vWSPPENLINE_Type.PENRT_String = ""
            '    vWSPPENLINE_Type.PENPOL_String = ""
            '    vWSPPENLINE_Type.PENPROJ_String = ""
            '    vWSPPENLINE_Type.PENEID_String = ""
            '    vWSPPENLINE_Type.PENPO_String = ""
            '    vWSPPENLINE_Type.PENHFRC_String = ""
            '    vWSPPENLINE_Type.PENLFRC_String = ""
            '    vWSPPENLINE_Type.PENHIGHF_Currency = 0
            '    vWSPPENLINE_Type.PENLOWF_Currency = 0
            '    vWSPPENLINE_Type.PENUC1_String = ""
            '    vWSPPENLINE_Type.PENUC2_String = ""
            '    vWSPPENLINE_Type.PENUC3_String = ""
            '    Call WSPPENLINE_Add(vWSPPENLINE_Type)
            '
            '    'PEN
            '    vCUSTCC_String = RemoveQuotes(rsSelectExcelSheet2.Fields("Ref 60"))
            '    vSheet2Ref63_String = RemoveQuotes(rsSelectExcelSheet2.Fields("Ref 63"))
            '    If Trim(vSheet2Ref63_String) <> "" Then
            '        If Trim(vCUSTCC_String) <> "" Then
            '            vCUSTCC_String = vCUSTCC_String & "/"
            '        End If
            '        vCUSTCC_String = vCUSTCC_String & vSheet2Ref63_String
            '    End If
            '    If Trim(vSheet1Department_String) <> "" Then
            '        If Trim(vCUSTCC_String) <> "" Then
            '            vCUSTCC_String = vCUSTCC_String & "/"
            '        End If
            '        vCUSTCC_String = vCUSTCC_String & vSheet1Department_String
            '    End If
            '
            '    vWSPPENLINE_Type.UpLoadNo_Long = UploadNo
            '    vWSPPENLINE_Type.RecID_String = "PEN"
            '    vWSPPENLINE_Type.AC_String = ""
            '    vWSPPENLINE_Type.DES_String = ""
            '    vWSPPENLINE_Type.REFE_String = ""
            '    vWSPPENLINE_Type.SELLADT_String = ""
            '    vWSPPENLINE_Type.SELLCHD_String = ""
            '    vWSPPENLINE_Type.SCHARGE_String = ""
            '    vWSPPENLINE_Type.DEPT_String = ""
            '    vWSPPENLINE_Type.BRANCH_String = RemoveQuotes(rsSelectExcelSheet1.Fields("BranchNumber"))
            '    vWSPPENLINE_Type.CUSTCC_String = vCUSTCC_String
            '    vWSPPENLINE_Type.PEMAIL_String = ""
            '    vWSPPENLINE_Type.PTELE_String = ""
            '    vWSPPENLINE_Type.PNO_Long = 0
            '    vWSPPENLINE_Type.FARE_Currency = 0
            '    vWSPPENLINE_Type.TAXES_Currency = 0
            '    vWSPPENLINE_Type.MARKUP_Currency = 0
            '    vWSPPENLINE_Type.SURNAMEFRSTNAMENO_String = ""
            '    vWSPPENLINE_Type.PENFARESELL_Currency = 0
            '    vWSPPENLINE_Type.PENFAREPASSTYPE_String = ""
            '    vWSPPENLINE_Type.DeliAdd_String = ""
            '    vWSPPENLINE_Type.MC_String = ""
            '    vWSPPENLINE_Type.BB_String = ""
            '    vWSPPENLINE_Type.PENFARESUPPID_String = ""
            '    vWSPPENLINE_Type.PENFAREAIRID_String = ""
            '    vWSPPENLINE_Type.PENLINKPNR_String = ""
            '    vWSPPENLINE_Type.PENOPRDID_String = ""
            '    vWSPPENLINE_Type.PENOQTY_Integer = 0
            '    vWSPPENLINE_Type.PENORATE_Currency = 0
            '    vWSPPENLINE_Type.PENOSELL_Currency = 0
            '    vWSPPENLINE_Type.PENOSUPPID_String = ""
            '    vWSPPENLINE_Type.PENATOLTYPE_String = ""
            '    vWSPPENLINE_Type.INETREF_String = ""
            '    vWSPPENLINE_Type.PENOPAYMETHODID_String = ""
            '    vWSPPENLINE_Type.PENRT_String = ""
            '    vWSPPENLINE_Type.PENPOL_String = ""
            '    vWSPPENLINE_Type.PENPROJ_String = ""
            '    vWSPPENLINE_Type.PENEID_String = RemoveQuotes(rsSelectExcelSheet2.Fields("Ref 11"))
            '    'by Abhi on 30-Jul-2012 for caseid 2401 Travcom data transfer Hotel rate
            '    vWSPPENLINE_Type.PENPO_String = vInvoiceID_String
            '    vWSPPENLINE_Type.PENHFRC_String = RemoveQuotes(rsSelectExcelSheet2.Fields("ExceptionCode"))
            '    vWSPPENLINE_Type.PENLFRC_String = ""
            '    vWSPPENLINE_Type.PENHIGHF_Currency = Val(RemoveQuotes(rsSelectExcelSheet2.Fields("ReferenceFare")))
            '    vWSPPENLINE_Type.PENLOWF_Currency = Val(RemoveQuotes(rsSelectExcelSheet2.Fields("LowFare")))
            '    vWSPPENLINE_Type.PENUC1_String = RemoveQuotes(rsSelectExcelSheet2.Fields("BookerName"))
            '    vWSPPENLINE_Type.PENUC2_String = ""
            '    vWSPPENLINE_Type.PENUC3_String = ""
            '    Call WSPPENLINE_Add(vWSPPENLINE_Type)
            '
            '    rsSelectExcelSheet2.MoveNext
            'Loop
            
            'WSPPENLINE PENO
            rsSelectExcelSheet2.Open "Select  [Sheet2$].* From [Sheet2$] WHERE [Sheet2$].InvoiceID='" & Val(vInvoiceID_String) & "' AND [Sheet2$].ProductGroupCode NOT IN('1','2','3','4') AND [Sheet2$].TransactionType='1'", CnExcel, adOpenForwardOnly, adLockReadOnly
            Do While rsSelectExcelSheet2.EOF = False
                vSheet2FormOfPayment_String = RemoveQuotes(rsSelectExcelSheet2.Fields("FormOfPayment"))
            
                'PENO
                vWSPPENLINE_Type.UpLoadNo_Long = UploadNo
                vWSPPENLINE_Type.RecID_String = "PENO"
                vWSPPENLINE_Type.AC_String = ""
                vWSPPENLINE_Type.DES_String = ""
                vWSPPENLINE_Type.REFE_String = ""
                vWSPPENLINE_Type.SELLADT_String = ""
                vWSPPENLINE_Type.SELLCHD_String = ""
                vWSPPENLINE_Type.SCHARGE_String = ""
                vWSPPENLINE_Type.DEPT_String = ""
                vWSPPENLINE_Type.BRANCH_String = ""
                vWSPPENLINE_Type.CUSTCC_String = ""
                vWSPPENLINE_Type.PEMAIL_String = ""
                vWSPPENLINE_Type.PTELE_String = ""
                vWSPPENLINE_Type.PNO_Long = 0
                vWSPPENLINE_Type.FARE_Currency = 0
                vWSPPENLINE_Type.TAXES_Currency = 0
                vWSPPENLINE_Type.MARKUP_Currency = 0
                vWSPPENLINE_Type.SURNAMEFRSTNAMENO_String = ""
                vWSPPENLINE_Type.PENFARESELL_Currency = 0
                vWSPPENLINE_Type.PENFAREPASSTYPE_String = ""
                vWSPPENLINE_Type.DeliAdd_String = ""
                vWSPPENLINE_Type.MC_String = ""
                vWSPPENLINE_Type.BB_String = ""
                vWSPPENLINE_Type.PENFARESUPPID_String = ""
                vWSPPENLINE_Type.PENFAREAIRID_String = ""
                vWSPPENLINE_Type.PENLINKPNR_String = ""
                'by Abhi on 31-Jul-2012 for caseid 2401 Travcom data transfer hotel only etc
                'vWSPPENLINE_Type.PENOPRDID_String = RemoveQuotes(rsSelectExcelSheet2.Fields("ProductCode"))
                vWSPPENLINE_Type.PENOPRDID_String = RemoveQuotes(rsSelectExcelSheet2.Fields("ProductGroupCode"))
                vWSPPENLINE_Type.PENOQTY_Integer = 1
                vWSPPENLINE_Type.PENORATE_Currency = Val(RemoveQuotes(rsSelectExcelSheet2.Fields("NetFare")))
                vWSPPENLINE_Type.PENOSELL_Currency = Val(RemoveQuotes(rsSelectExcelSheet2.Fields("SellingFare")))
                vWSPPENLINE_Type.PENOSUPPID_String = RemoveQuotes(rsSelectExcelSheet2.Fields("VendorNumber"))
                vWSPPENLINE_Type.PENATOLTYPE_String = ""
                vWSPPENLINE_Type.INETREF_String = ""
                If Val(vSheet2FormOfPayment_String) = 1 Then
                    vWSPPENLINE_Type.PENOPAYMETHODID_String = "CP"
                Else
                    vWSPPENLINE_Type.PENOPAYMETHODID_String = ""
                End If
                vWSPPENLINE_Type.PENRT_String = ""
                vWSPPENLINE_Type.PENPOL_String = ""
                vWSPPENLINE_Type.PENPROJ_String = ""
                vWSPPENLINE_Type.PENEID_String = ""
                vWSPPENLINE_Type.PENPO_String = ""
                vWSPPENLINE_Type.PENHFRC_String = ""
                vWSPPENLINE_Type.PENLFRC_String = ""
                vWSPPENLINE_Type.PENHIGHF_Currency = 0
                vWSPPENLINE_Type.PENLOWF_Currency = 0
                vWSPPENLINE_Type.PENUC1_String = ""
                vWSPPENLINE_Type.PENUC2_String = ""
                vWSPPENLINE_Type.PENUC3_String = ""
                Call WSPPENLINE_Add(vWSPPENLINE_Type)
                
                'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
                If Trim(vWSPPENLINE_Type.AC_String) <> "" Then
                    GIT_CUSTOMERUSERCODE_String = vWSPPENLINE_Type.AC_String
                End If
                'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
                
                rsSelectExcelSheet2.MoveNext
            Loop
            
            'WSPPENLINE PEN
            If rsSelectExcelSheet2.State = 1 Then rsSelectExcelSheet2.Close
            rsSelectExcelSheet2.Open "Select  [Sheet2$].* From [Sheet2$] WHERE [Sheet2$].InvoiceID='" & Val(vInvoiceID_String) & "' AND [Sheet2$].TransactionType='1'", CnExcel, adOpenForwardOnly, adLockReadOnly
            Do While rsSelectExcelSheet2.EOF = False

                'PEN
                vCUSTCC_String = RemoveQuotes(rsSelectExcelSheet2.Fields("Ref 60"))
                vSheet2Ref63_String = RemoveQuotes(rsSelectExcelSheet2.Fields("Ref 63"))
                If Trim(vSheet2Ref63_String) <> "" Then
                    If Trim(vCUSTCC_String) <> "" Then
                        vCUSTCC_String = vCUSTCC_String & "/"
                    End If
                    vCUSTCC_String = vCUSTCC_String & vSheet2Ref63_String
                End If
                If Trim(vSheet1Department_String) <> "" Then
                    If Trim(vCUSTCC_String) <> "" Then
                        vCUSTCC_String = vCUSTCC_String & "/"
                    End If
                    vCUSTCC_String = vCUSTCC_String & vSheet1Department_String
                End If
            
                vWSPPENLINE_Type.UpLoadNo_Long = UploadNo
                vWSPPENLINE_Type.RecID_String = "PEN"
                vWSPPENLINE_Type.AC_String = ""
                vWSPPENLINE_Type.DES_String = ""
                vWSPPENLINE_Type.REFE_String = ""
                vWSPPENLINE_Type.SELLADT_String = ""
                vWSPPENLINE_Type.SELLCHD_String = ""
                vWSPPENLINE_Type.SCHARGE_String = ""
                vWSPPENLINE_Type.DEPT_String = ""
                vWSPPENLINE_Type.BRANCH_String = RemoveQuotes(rsSelectExcelSheet1.Fields("BranchNumber"))
                vWSPPENLINE_Type.CUSTCC_String = vCUSTCC_String
                vWSPPENLINE_Type.PEMAIL_String = ""
                vWSPPENLINE_Type.PTELE_String = ""
                vWSPPENLINE_Type.PNO_Long = 0
                vWSPPENLINE_Type.FARE_Currency = 0
                vWSPPENLINE_Type.TAXES_Currency = 0
                vWSPPENLINE_Type.MARKUP_Currency = 0
                vWSPPENLINE_Type.SURNAMEFRSTNAMENO_String = ""
                vWSPPENLINE_Type.PENFARESELL_Currency = 0
                vWSPPENLINE_Type.PENFAREPASSTYPE_String = ""
                vWSPPENLINE_Type.DeliAdd_String = ""
                vWSPPENLINE_Type.MC_String = ""
                vWSPPENLINE_Type.BB_String = ""
                vWSPPENLINE_Type.PENFARESUPPID_String = ""
                vWSPPENLINE_Type.PENFAREAIRID_String = ""
                vWSPPENLINE_Type.PENLINKPNR_String = ""
                vWSPPENLINE_Type.PENOPRDID_String = ""
                vWSPPENLINE_Type.PENOQTY_Integer = 0
                vWSPPENLINE_Type.PENORATE_Currency = 0
                vWSPPENLINE_Type.PENOSELL_Currency = 0
                vWSPPENLINE_Type.PENOSUPPID_String = ""
                vWSPPENLINE_Type.PENATOLTYPE_String = ""
                vWSPPENLINE_Type.INETREF_String = ""
                vWSPPENLINE_Type.PENOPAYMETHODID_String = ""
                vWSPPENLINE_Type.PENRT_String = ""
                vWSPPENLINE_Type.PENPOL_String = ""
                vWSPPENLINE_Type.PENPROJ_String = ""
                vWSPPENLINE_Type.PENEID_String = RemoveQuotes(rsSelectExcelSheet2.Fields("Ref 11"))
                'by Abhi on 30-Jul-2012 for caseid 2401 Travcom data transfer Hotel rate
                'by Abhi on 18-Aug-2014 for caseid 4439 Change the mapping  "InvoiceNo" instead of "InvoiceID" for Legacy Loading (Travcom Mapping)
                'vWSPPENLINE_Type.PENPO_String = vInvoiceID_String
                vWSPPENLINE_Type.PENPO_String = vInvoiceNo_String
                'by Abhi on 18-Aug-2014 for caseid 4439 Change the mapping  "InvoiceNo" instead of "InvoiceID" for Legacy Loading (Travcom Mapping)
                vWSPPENLINE_Type.PENHFRC_String = RemoveQuotes(rsSelectExcelSheet2.Fields("ExceptionCode"))
                vWSPPENLINE_Type.PENLFRC_String = ""
                vWSPPENLINE_Type.PENHIGHF_Currency = Val(RemoveQuotes(rsSelectExcelSheet2.Fields("ReferenceFare")))
                vWSPPENLINE_Type.PENLOWF_Currency = Val(RemoveQuotes(rsSelectExcelSheet2.Fields("LowFare")))
                vWSPPENLINE_Type.PENUC1_String = RemoveQuotes(rsSelectExcelSheet2.Fields("BookerName"))
                vWSPPENLINE_Type.PENUC2_String = ""
                vWSPPENLINE_Type.PENUC3_String = ""
                Call WSPPENLINE_Add(vWSPPENLINE_Type)
            
                rsSelectExcelSheet2.MoveNext
            Loop
            
            'by Abhi on 13-Aug-2014 for caseid 4425 Client Reference details not loading for Customer Card Transactions
            'WSPHTLSEG
            If rsSelectExcelSheet2.State = 1 Then rsSelectExcelSheet2.Close
            'by Abhi on 31-Jul-2012 for caseid 2401 Travcom data transfer rail
            'rsSelectExcelSheet2.Open "Select  [Sheet2$].* From [Sheet2$] WHERE [Sheet2$].InvoiceID='" & Val(vInvoiceID_String) & "' AND [Sheet2$].ProductGroupCode IN('4')", CnExcel, adOpenForwardOnly, adLockReadOnly
            rsSelectExcelSheet2.Open "Select  [Sheet2$].* From [Sheet2$] WHERE [Sheet2$].InvoiceID='" & Val(vInvoiceID_String) & "' AND [Sheet2$].ProductGroupCode IN('4') AND [Sheet2$].TransactionType='1'", CnExcel, adOpenForwardOnly, adLockReadOnly
            Do While rsSelectExcelSheet2.EOF = False
                'by Abhi on 30-Jul-2012 for caseid 2401 Travcom data transfer Hotel rate
                vSheet2PassengerName_String = RemoveQuotes(rsSelectExcelSheet2.Fields("PassengerName"))
                SubSplited = SplitForce(vSheet2PassengerName_String, "/", 2)
                SubSplited2 = SplitFirstNameAndInitial(SubSplited(1))
                
                
                vWSPHTLSEG_Type.UpLoadNo_Long = UploadNo
                vWSPHTLSEG_Type.RecID_String = ""
                vWSPHTLSEG_Type.SEGNO_String = ""
                vWSPHTLSEG_Type.CHAINCODE_String = RemoveQuotes(rsSelectExcelSheet2.Fields("VendorNumber"))
                vWSPHTLSEG_Type.CHAINNAME_String = RemoveQuotes(rsSelectExcelSheet2.Fields("VendorName"))
                vWSPHTLSEG_Type.STATUSCODE_String = ""
                vWSPHTLSEG_Type.ROOMS_String = ""
                vWSPHTLSEG_Type.CTYCODE_String = RemoveQuotes(rsSelectExcelSheet2.Fields("CityCode"))
                vWSPHTLSEG_Type.INDATE_String = GetAirSegDate(RemoveQuotes(rsSelectExcelSheet2.Fields("TravelDate")))
                vWSPHTLSEG_Type.OUTDATE_String = GetAirSegDate(RemoveQuotes(rsSelectExcelSheet2.Fields("ReturnDate")))
                vWSPHTLSEG_Type.PROPCODE_String = RemoveQuotes(rsSelectExcelSheet2.Fields("ChainCode"))
                vWSPHTLSEG_Type.PROPNAME_String = ""
                vWSPHTLSEG_Type.NOPER_String = RemoveQuotes(rsSelectExcelSheet2.Fields("NumberOfPassengers"))
                vWSPHTLSEG_Type.TYPE_String = ""
                vWSPHTLSEG_Type.CURRCODE_String = RemoveQuotes(rsSelectExcelSheet2.Fields("CurrencyCode"))
                vWSPHTLSEG_Type.SYSRAMT_String = ""
                vWSPHTLSEG_Type.SYSRPLAN_String = ""
                vWSPHTLSEG_Type.ROOMDESC_String = ""
                vWSPHTLSEG_Type.RATEDESC_String = ""
                vWSPHTLSEG_Type.ROOMLOC_String = ""
                vWSPHTLSEG_Type.BSOURCE_String = ""
                'by Abhi on 30-Jul-2012 for caseid 2401 Travcom data transfer Hotel rate
                'vWSPHTLSEG_Type.GUESTNAME_String = RemoveQuotes(rsSelectExcelSheet2.Fields("PassengerName"))
                vWSPHTLSEG_Type.GUESTNAME_String = Replace(SubSplited2(0), "*CHD", "") & " " & SubSplited(0)
                vWSPHTLSEG_Type.DISCOUNTNO_String = ""
                vWSPHTLSEG_Type.FTNO_String = ""
                vWSPHTLSEG_Type.FGUESTNO_String = ""
                vWSPHTLSEG_Type.TOTALAMT_String = RemoveQuotes(rsSelectExcelSheet2.Fields("SellingFare"))
                vWSPHTLSEG_Type.BASEAMT_String = ""
                vWSPHTLSEG_Type.SCHRGEAMT_String = ""
                vWSPHTLSEG_Type.SURCHARGE_String = ""
                vWSPHTLSEG_Type.TTD_String = ""
                vWSPHTLSEG_Type.COMAMT_String = ""
                vWSPHTLSEG_Type.VCOM_String = ""
                vWSPHTLSEG_Type.CONFNO_String = ""
                vWSPHTLSEG_Type.CNCLNNO_String = ""
                vWSPHTLSEG_Type.TAXRATE_String = RemoveQuotes(rsSelectExcelSheet2.Fields("VATCode"))
                vWSPHTLSEG_Type.HOTADD1_String = ""
                vWSPHTLSEG_Type.HOTADD2_String = ""
                vWSPHTLSEG_Type.CNTRYCODE_String = ""
                vWSPHTLSEG_Type.POSTELCODE_String = ""
                vWSPHTLSEG_Type.TELENO_String = ""
                vWSPHTLSEG_Type.FAXNO_String = ""
                vWSPHTLSEG_Type.CHKINTIME_String = ""
                vWSPHTLSEG_Type.CHKOUTTIME_String = ""
                'by Abhi on 30-Jul-2012 for caseid 2401 Travcom data transfer Hotel rate
                vWSPHTLSEG_Type.RGCURRCODE_String = RemoveQuotes(rsSelectExcelSheet2.Fields("CurrencyCode"))
                'by Abhi on 30-Jul-2012 for caseid 2401 Travcom data transfer Hotel rate
                vWSPHTLSEG_Type.RGAMT_String = RemoveQuotes(rsSelectExcelSheet2.Fields("VoucherRate"))
                'by Abhi on 25-Sep-2012 for caseid 2479 Travcom data transfer Some of the cases VoucherRate in sheet 2 is zero for hotel.In this case can you please pick the amount from sheet2.NetFare to WSPHTLSEG.RGAMT
                If Val(vWSPHTLSEG_Type.RGAMT_String) = 0 Then
                    vWSPHTLSEG_Type.RGAMT_String = RemoveQuotes(rsSelectExcelSheet2.Fields("NetFare"))
                End If
                vWSPHTLSEG_Type.EXRATEDTE_String = ""
                'by Abhi on 30-Jul-2012 for caseid 2401 Travcom data transfer Hotel rate
                vWSPHTLSEG_Type.CURRCODELOC_String = RemoveQuotes(rsSelectExcelSheet2.Fields("CurrencyCode"))
                vWSPHTLSEG_Type.CNTRYCODEH_String = ""
                vWSPHTLSEG_Type.EXCHRATE_String = ""
                vWSPHTLSEG_Type.RT_String = ""
                vWSPHTLSEG_Type.RTCURRCODE_String = ""
                vWSPHTLSEG_Type.RTAMT_String = ""
                vWSPHTLSEG_Type.RQ_String = ""
                vWSPHTLSEG_Type.RQCURRCODE_String = ""
                vWSPHTLSEG_Type.RQAMT_String = ""
                vWSPHTLSEG_Type.RQ1_String = ""
                vWSPHTLSEG_Type.RQ1CURRCODE_String = ""
                vWSPHTLSEG_Type.RQ1AMT_String = ""
                vWSPHTLSEG_Type.RG1_String = ""
                vWSPHTLSEG_Type.RG1CURRCODE_String = ""
                vWSPHTLSEG_Type.RG1AMT_String = ""
                vWSPHTLSEG_Type.RT1_String = ""
                vWSPHTLSEG_Type.RT1CURRCODE_String = ""
                vWSPHTLSEG_Type.RT1AMT_String = ""
                Call WSPHTLSEG_Add(vWSPHTLSEG_Type)
                
                rsSelectExcelSheet2.MoveNext
            Loop
            'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
            FMain.stbUpload.Panels(1).Text = "Writing GDS In Tray..."
            DoEvents
            'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
            'vSQL_String = "" _
            '    & "INSERT INTO dbo.GDSINTRAYTABLE " _
            '    & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC]) " _
            '    & "SELECT     'XLS' AS GIT_ID, '' AS GIT_GDS, 'LEGACY' AS GIT_INTERFACE, dbo.WSPPNRADD.UpLoadNo AS GIT_UPLOADNO, dbo.WSPPNRADD.DLCI AS GIT_PNRDATE, " _
            '    & "                      dbo.WSPPNRADD.PNRADD AS GIT_PNR, dbo.WSPPNAME.SURNAME AS GIT_LASTNAME, dbo.WSPPNAME.FRSTNAME AS GIT_FIRSTNAME, " _
            '    & "                      dbo.WSPAGNTSINE.BAGENT AS GIT_BOOKINGAGENT, dbo.WSPPNAME.DOCNO AS GIT_TICKETNUMBER, dbo.WSPPNRADD.FNAME AS GIT_FILENAME, " _
            '    & "                      dbo.WSPPNRADD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.WSPAGNTSINE.BSID AS GIT_PCC " _
            '    & "FROM         dbo.WSPPNAME INNER JOIN " _
            '    & "                      dbo.WSPPNRADD ON dbo.WSPPNAME.UpLoadNo = dbo.WSPPNRADD.UpLoadNo INNER JOIN " _
            '    & "                      dbo.WSPAGNTSINE ON dbo.WSPPNRADD.UpLoadNo = dbo.WSPAGNTSINE.UpLoadNo " _
            '    & "WHERE     (dbo.WSPPNRADD.UpLoadNo = " & UploadNo & ") AND (dbo.WSPPNRADD.RecID = '')"
            vDateLastModified_String = fsObj.GetFile(FileName).DateLastModified
            'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
            'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
            vDateLastModified_String = DateTime12hrsFormat(vDateLastModified_String)
            vFileDateCreated_String = fsObj.GetFile(FileName).DateCreated
            vFileDateCreated_String = DateTime12hrsFormat(vFileDateCreated_String)
            'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
            'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
            GIT_PENLINEBRID_String = ""
            'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
            'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
            'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
            ErrDetails_String = " GDSINTRAYTABLE."
            'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
            vSQL_String = "" _
                & "INSERT INTO dbo.GDSINTRAYTABLE " _
                & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC],[GIT_FILEISSUEDATE],[GIT_CUSTOMERUSERCODE],[GIT_PENWAIT],[GIT_FILECREATEDDATE],[GIT_ISSUEDBY], [GIT_PENLINEBRID]) " _
                & "SELECT     'XLS' AS GIT_ID, '' AS GIT_GDS, 'LEGACY' AS GIT_INTERFACE, dbo.WSPPNRADD.UpLoadNo AS GIT_UPLOADNO, dbo.WSPPNRADD.DLCI AS GIT_PNRDATE, " _
                & "                      dbo.WSPPNRADD.PNRADD AS GIT_PNR, dbo.WSPPNAME.SURNAME AS GIT_LASTNAME, dbo.WSPPNAME.FRSTNAME AS GIT_FIRSTNAME, " _
                & "                      dbo.WSPAGNTSINE.BAGENT AS GIT_BOOKINGAGENT, dbo.WSPPNAME.DOCNO AS GIT_TICKETNUMBER, dbo.WSPPNRADD.FNAME AS GIT_FILENAME, " _
                & "                      dbo.WSPPNRADD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.WSPAGNTSINE.BSID AS GIT_PCC, '" & vDateLastModified_String & "' AS GIT_FILEISSUEDATE, " _
                & "                      '" & SkipChars(GIT_CUSTOMERUSERCODE_String) & "' AS GIT_CUSTOMERUSERCODE, '" & SkipChars(GIT_PENWAIT_String) & "' AS GIT_PENWAIT, " _
                & "                      '" & vFileDateCreated_String & "' AS GIT_FILECREATEDDATE, '' AS GIT_ISSUEDBY, '" & GIT_PENLINEBRID_String & "' AS GIT_PENLINEBRID " _
                & "FROM         dbo.WSPPNAME INNER JOIN " _
                & "                      dbo.WSPPNRADD ON dbo.WSPPNAME.UpLoadNo = dbo.WSPPNRADD.UpLoadNo INNER JOIN " _
                & "                      dbo.WSPAGNTSINE ON dbo.WSPPNRADD.UpLoadNo = dbo.WSPAGNTSINE.UpLoadNo " _
                & "WHERE     (dbo.WSPPNRADD.UpLoadNo = " & UploadNo & ") AND (dbo.WSPPNRADD.RecID = '')"
            'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
            'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
            'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
            dbCompany.Execute vSQL_String
            FMain.stbUpload.Panels(1).Text = "File moving to Destination..."
            'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
            'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
            FMain.prgUpload.value = vRecordNumber_Long
            vRecordNumber_Long = vRecordNumber_Long + 1
            'by Abhi on 01-Nov-2013 for caseid 3497 PenGDS Legacy-Excel file went to error files and not reading the
            rsSelectExcelSheet1.MoveNext
        Loop
        If rsSelectExcelSheet1.State = adStateOpen Then rsSelectExcelSheet1.Close
        Set rsSelectExcelSheet1 = Nothing
        If rsSelectExcelSheet2.State = adStateOpen Then rsSelectExcelSheet2.Close
        Set rsSelectExcelSheet2 = Nothing
        If rsSelectExcelSheet3.State = adStateOpen Then rsSelectExcelSheet3.Close
        Set rsSelectExcelSheet3 = Nothing
        If CnExcel.State = adStateOpen Then CnExcel.Close
        Set CnExcel = Nothing
        FileObj.CopyFile txtFileName.Text, txtRelLocation.Text & "\", True
        'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
        'On Error Resume Next
        FileObj.DeleteFile txtFileName.Text, True
        'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
        INIWrite App.Path & "\PenSEARCH.ini", "Legacy", "LUFPNR", LUFPNR_String
        INIWrite App.Path & "\PenSEARCH.ini", "Legacy", "LUFDate", DateFormat(Date)
        INIWrite App.Path & "\PenSEARCH.ini", "Legacy", "LUFTime", TimeFormat12HRS(time) & "(" & TimeFormat(time) & ")"
        'by Abhi on 18-May-2011 for caseid 1757 Added Events for GDSAuto
        'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
        'Call EventLog("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & LUFPNR_String & " - " & FMain.stbUpload.Panels(3).Text)
        Call EventLog("#PenGDS# " & PadR(FMain.stbUpload.Panels(2).Text, 9) & " - " & PadR(LUFPNR_String, 6) & " - " & PadR(FMain.stbUpload.Panels(3).Text, 27))
        'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        dbCompany.CommitTrans
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        PENErr_BeginTrans = False
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    End If

Note:
    Me.Hide
Exit Function
PENErr:
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'If PENErr() = True Then GoTo DeadlockRETRY
    If PENErr() = True Then Resume DeadlockRETRY
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
End Function
