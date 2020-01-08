VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FAmadeusUpload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AMADEUS Uploading"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5460
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
      Left            =   3870
      TabIndex        =   3
      Top             =   2130
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
Attribute VB_Name = "FAmadeusUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim dbCompany As New ADODB.Connection
Dim FNAME As String
Dim TKT_NoIndex As String
'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue
Dim vPSGRID_String As String
'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
Dim fPNRDTE_String As String
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date


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
'            DeleteFileToRecycleBin (txtFileName.Text)
'        Next
'    End If
'Exit Sub
'Note:
'    MsgBox Err.Description & ". Invalid Drive Location", vbCritical: Exit Sub
'
End Sub

Public Sub cmdDirUpload_Click()
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'On Error GoTo Note
    Dim Count As Integer
    
    FMain.SendStatus FMain.SSTab1.TabCaption(4)
    File1.Path = txtDirName.Text
    File1.Pattern = FMain.txtAmadeusExt
    If txtDirName.Text = "" Or File1.ListCount = 0 Then FMain.stbUpload.Panels(1).Text = "MIR Files Not Found...": Exit Sub
    'Uploading Each File
    Open (App.Path & "\_UploadingSQL_") For Random As #1
    Close #1
    FMain.cmdStop.Enabled = False
    DoEvents
    FMain.lblFName.Caption = "Uploading Amadeus..."
    FMain.stbUpload.Panels(1).Text = "Reading..."
    FMain.stbUpload.Panels(2).Text = "Amadeus"
    FMain.prgUpload.Max = IIf((File1.ListCount = 1), File1.ListCount, File1.ListCount - 1)
    Pgr.value = 5
    
'    If File1.ListCount > 0 Then
'        If MsgBox("Do You Wish To Clear First ?", vbQuestion + vbYesNo) = vbYes Then
'            ClearAll
'        End If
'    End If
    'by Abhi on 18-May-2011 for caseid 1757 Added Events for GDSAuto
    Call EventLogDelete
    'by Abhi on 06-Sep-2017 for caseid 7766 Penline and Company currency change in penair is not reflecting in pengds
    Call getPENLINEID
    COMCID_String = getFromFileTable("COMCID")
    'by Abhi on 06-Sep-2017 for caseid 7766 Penline and Company currency change in penair is not reflecting in pengds
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    FMain.OutOfBookingsDateAmadeusText.Text = DateFormat1900toBlank(getFromFileTable("OutOfBookingsDateAmadeus"))
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
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
            AddFile txtFileName.Text
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
'    Exit Sub
'    Resume
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

Private Sub Form_Load()
'dbCompany.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=sa;Initial Catalog=PENDATA001;Data Source=PEN1\PENSOFT"
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'On Error GoTo Note
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'Dim rsSelect As New ADODB.Recordset
    Me.Icon = FMain.Icon
    'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsSelect.Open "Select AMDUPLOADDIRNAME,AMDDESTDIRNAME From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsSelect.Open "Select AMDUPLOADDIRNAME,AMDDESTDIRNAME From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
    'If rsSelect.EOF = False Then
    '    txtDirName = IIf(IsNull(rsSelect!AMDUPLOADDIRNAME), "", rsSelect!AMDUPLOADDIRNAME)
    '    txtRelLocation = IIf(IsNull(rsSelect!AMDDESTDIRNAME), "", rsSelect!AMDDESTDIRNAME)
    'End If
    'rsSelect.Close
    'If rsSelect.State = 1 Then rsSelect.Close
    ''by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    ''rsSelect.Open "Select AMDSTATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    'rsSelect.Open "Select AMDSTATUS From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
    'chkGal.value = IIf((IsNull(rsSelect!AMDSTATUS) = True), 0, rsSelect!AMDSTATUS)
    txtDirName = FMain.txtAmadeusSource
    txtRelLocation = FMain.txtAmadeusDest
    chkGal.value = FMain.chkAmadeus
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    File1.Path = txtDirName.Text
    File1.Refresh
    File1.Pattern = FMain.txtAmadeusExt
    File1.Refresh
    Exit Sub
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'Note:
'    If Me.Visible = True Then Unload Me

End Sub

Private Sub Text1_Change()
Caption = Len(Text1)
End Sub



Private Function BLK(Data) As Collection
Dim temp
Dim temp2
Dim Coll As New Collection
temp = Split(Data, ";")
Coll.Add temp(0), "AIRVER"
Coll.Add temp(1), "AIROPT "
Coll.Add temp(2), "AIRTRN"
Coll.Add temp(3), "HDRSIZ"
Coll.Add temp(4), "TRANNUM"
                            temp2 = SplitTwo(temp(5), 2)
Coll.Add temp2(0), "NETID"
Coll.Add temp2(1), "ADDR"
                            temp2 = SplitTwo(temp(6), 3)
Coll.Add temp2(0), "BLKNUM"
Coll.Add temp2(0), "TOTBLK"
Set BLK = Coll
End Function







Private Function SplitTwo(Data, First As Long)
On Error Resume Next
Dim temp(1) As String
temp(0) = Left(Data, First)
If Len(Data) > 0 Then
temp(1) = Mid(Data, First + 1, Len(Data) - First)
End If
SplitTwo = temp
End Function

Private Function SplitTwoReverse(Data, First As Long)
On Error Resume Next
Dim temp(1) As String
temp(1) = Right(Data, First)
If (Len(Data) > 0) Then
temp(0) = Mid(Data, 1, SkipNegative(Len(Data) - First))
End If
SplitTwoReverse = temp
End Function

Private Function SplitFirstTwo(Data, find As String)
Dim temp
Dim tm(1) As String
temp = Split(CStr(Data), find, 2)
If UBound(temp) > -1 Then
tm(0) = temp(0)
End If
If UBound(temp) > 0 Then
tm(1) = temp(1)
End If

SplitFirstTwo = tm
End Function

Private Function SplitField(Data, Start As String, Finish As String)
Dim rStart, rMid, rEnd As String
Dim retn(2) As String
Dim temp, temp2, temp3
    temp = SplitFirstTwo(Data, "(")
    rStart = temp(0)
    temp2 = SplitFirstTwo(temp(1), ")")
    rMid = temp2(0)
    rEnd = temp2(1)
retn(0) = rStart
retn(1) = rMid
retn(2) = rEnd
SplitField = retn
End Function

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
Private Function SplitWithLengthsPlus(Data, ParamArray Lengths())
On Error Resume Next
Dim Nos As Integer
Dim pos As Integer, j As Integer
Nos = UBound(Lengths)
Dim temp() As String
ReDim temp(Nos + 1) As String
pos = 1

For j = 0 To Nos
    temp(j) = Mid(Data, pos, Lengths(j))
    pos = pos + Lengths(j)
    DoEvents
Next
If Len(Data) > 0 Then
    temp(Nos + 1) = Mid(Data, pos, SkipNull(Len(Data) - pos + 1))
End If
SplitWithLengthsPlus = temp

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

Private Function AMD(Data) As Collection
Dim Coll As New Collection
Dim Lines
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
'by Abhi on 09-Nov-2009 VLOCATOR in AirSegDetails
Dim CollAMD As New Collection
Dim CollVLOCATOR As New Collection
Dim CollVLOCATORSplited As New Collection
Dim vi_Long As Long
Dim viVL_Long As Long

Lines = SplitForce(Data, Chr(13) & Chr(10), 3)
'splited = SplitForce(Lines(0), ";", 4)
Splited = SplitForce(Lines(0), ";", 32)

'Line1
    temp = SplitWithLengths(Splited(0), 1, 2, 8)
    CollAMD.Add temp(0), "RETRANS"
    CollAMD.Add temp(1), "DOM"
    CollAMD.Add temp(2), "AIRSEQ "
    
    temp = SplitForce(Splited(1), "/", 2)
    CollAMD.Add temp(0), "AIRSEC"
    CollAMD.Add temp(1), "AIRTOTAL"
    
    temp = SplitTwo(Splited(2), 4)
    CollAMD.Add temp(0), "ACTTAG"
    CollAMD.Add Left(temp(1), 5), "DATE"
    
    CollAMD.Add Splited(3), "AGTSIN"
'Line 2
    Splited = SplitForce(Lines(1), ";", 4)
    
    temp = SplitTwo(Splited(0), 2)
    CollAMD.Add temp(0), "NETID"
    CollAMD.Add temp(1), "ADDRTKT"
    
    temp = SplitTwo(Splited(1), 2)
    CollAMD.Add temp(0), "NETID_BKO"
    CollAMD.Add temp(1), "ADDRTKT_BKO"
    
    CollAMD.Add Splited(2), "IDENTBOS"
    CollAMD.Add Splited(3), "BOSTYPE"

'Line 3
Splited = SplitForce(Lines(2), ";", 32)

'sec1
    temp = SplitWithLengths(Splited(0), 3, 3, 6, 3)
    CollAMD.Add temp(0), "CTYID"
    CollAMD.Add temp(1), "AIRCDE"
    CollAMD.Add temp(2), "RECLOC"
    'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
    LUFPNR_String = temp(2)
    CollAMD.Add temp(3), "PNRENV"
'sec2
    temp = SplitWithLengths(Splited(1), 2, 2, 1)
    CollAMD.Add temp(0), "TOTPSG"
    CollAMD.Add temp(1), "AIRPSG"
    CollAMD.Add temp(2), "NHPNR"

    CollAMD.Add Splited(2), "IDBKG"
    CollAMD.Add Splited(3), "BKGAGCY"
    CollAMD.Add Splited(4), "IDFIRST"
    CollAMD.Add Splited(5), "OWNAGY"
    CollAMD.Add Splited(6), "IDENTCO"
    CollAMD.Add Splited(7), "CURAGCY"
    CollAMD.Add Splited(8), "IDENTT"
    CollAMD.Add Splited(9), "TKTAGCY"
    CollAMD.Add Splited(10), "OFFID"
    CollAMD.Add Splited(11), "CPNID"
    CollAMD.Add Splited(12), "IDENTT2"
    CollAMD.Add Splited(13), "TKTAGCY2"
    CollAMD.Add Splited(14), "OFFID2"
    CollAMD.Add Splited(15), "CPNID2"
    CollAMD.Add Splited(16), "IDENTT3"
    CollAMD.Add Splited(17), "TKTAGCY3"
    CollAMD.Add Splited(18), "OFFID3"
    CollAMD.Add Splited(19), "CPNID3"
    CollAMD.Add Splited(20), "IDENTT4"
    CollAMD.Add Splited(21), "TKTAGCY4"
    CollAMD.Add Splited(22), "OFFID4"
    CollAMD.Add Splited(23), "CPNID4"
    CollAMD.Add Splited(24), "IDENTT5"
    CollAMD.Add Splited(25), "TKTAGCY5"
    CollAMD.Add Splited(26), "OFFID5"
    CollAMD.Add Splited(27), "CPNID5"
    CollAMD.Add Splited(28), "ACCTNUM"
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    If Trim(Splited(28)) <> "" Then
        GIT_CUSTOMERUSERCODE_String = Splited(28)
    End If
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    CollAMD.Add Splited(29), "ORDNUM"
    CollAMD.Add Splited(30), "ERSP"
'sec 32
    sTemp1 = Splited(31)
'by Abhi on 09-Nov-2009 VLOCATOR in AirSegDetails
'    sTemp2 = Right(sTemp1, 6)
'    sTemp3 = Replace(sTemp1, sTemp2, "")
'    temp = SplitWithLengths(sTemp3, 28, 3)
'
'    CollAMD.Add temp(0), "NSPNR"
'    CollAMD.Add temp(1), "NSAIR"
'    CollAMD.Add sTemp2, "NSRECLOC"
    sTemp2 = Right(sTemp1, 9)
    sTemp3 = Replace(sTemp1, sTemp2, "")
    
    CollAMD.Add sTemp3, "NSPNR"
    CollAMD.Add "", "NSAIR"
    CollAMD.Add "", "NSRECLOC"
    
    'temp = SplitWithLengths(sTemp2, 3, 6)
    'CollVLOCATOR.Add temp(0), "NSAIR"
    'CollVLOCATOR.Add temp(1), "NSRECLOC"
    
    viVL_Long = 1
    'by Abhi on 25-Nov-2009 for nothing comes it will not create a collection
    If sTemp2 <> "" Then
        Set CollVLOCATORSplited = AMDVLOCATOR(sTemp2)
        CollVLOCATOR.Add CollVLOCATORSplited, str(viVL_Long)
        viVL_Long = viVL_Long + 1
    End If
    For vi_Long = 32 To UBound(Splited)
        sTemp2 = Splited(vi_Long)
        If sTemp2 <> "" Then
            Set CollVLOCATORSplited = AMDVLOCATOR(sTemp2)
            CollVLOCATOR.Add CollVLOCATORSplited, str(viVL_Long)
            viVL_Long = viVL_Long + 1
        End If
        DoEvents
    Next
    
    Coll.Add CollAMD, "AMD"
    Coll.Add CollVLOCATOR, "VLOCATOR"

Set AMD = Coll
End Function

Private Function AMDVLOCATOR(Data) As Collection
'by Abhi on 09-Nov-2009 VLOCATOR in AirSegDetails
Dim Coll As New Collection
Dim sTemp2 As String
Dim temp

    sTemp2 = Data
    'by Abhi on 25-Nov-2009 for nothing comes it will not create a collection but it must create a collection
    'If sTemp2 <> "" Then
        temp = SplitWithLengths(sTemp2, 3, 6)
        Coll.Add temp(0), "NSAIR"
        Coll.Add temp(1), "NSRECLOC"
    'End If
Set AMDVLOCATOR = Coll
End Function


Private Function A(Data As String) As Collection
Dim Coll As New Collection
Dim Lines As Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = SplitForce(Data, ";", 2)

'sec 1
Coll.Add Splited(0), "TKTAIR"
'sec 2
temp = SplitWithLengths(Splited(1), 3, 3)
Coll.Add temp(0), "ALCDE"
Coll.Add temp(1), "NUMCDE"

Set A = Coll
End Function

Private Function B(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Splited = SplitForce(Data, ";", 2)
Coll.Add Splited(0), "ENTRY"
Set B = Coll
End Function

Private Function C(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = SplitTwo(Data, 5)
'Sec 1
temp = SplitWithLengths(Splited(0), 4, 1)
    Coll.Add temp(0), "SVCCAR"
    Coll.Add temp(1), "TKTIND"
temp = SplitForce(Splited(1), "-", 7)
temp2 = SplitTwo(temp(0), 1)
Coll.Add temp2(1), "BKAGT"
Coll.Add temp(1), "TKTAGT"
Coll.Add temp(2), "PRCDE"
Coll.Add temp(3), "FCMI"
Coll.Add temp(4), "JRNYTYPE"
Coll.Add temp(5), "JRNYCODE"
Coll.Add temp(6), "PSURCHG"
Set C = Coll
End Function

Private Function d(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = SplitForce(Data, ";", 3)
Coll.Add Splited(0), "PNRDTE"
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
If Len(Splited(0)) = 6 Then
    fPNRDTE_String = DateFormat(ToPenDateYYMMDD(Splited(0)))
End If
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
Coll.Add Splited(1), "CHGDTE"
Coll.Add Splited(2), "AIRDTE"
Set d = Coll
End Function


Private Function G(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = Split(Data, ";")
Coll.Add Splited(0), "INTIND"
Coll.Add Splited(1), "SALEIND"
Coll.Add Splited(2), "ORGDEST"
Coll.Add Splited(3), "DOMISO"
Set G = Coll
End Function


Private Function H(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = SplitForce(Data, ";", 24)

Coll.Add Splited(0), "TTONBR"

temp = SplitWithLengthsPlus(Splited(1), 3, 1)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "STOP"
    Coll.Add temp(2), "ORIGAIR"
Coll.Add Splited(2), "ORIGCTY"
Coll.Add Splited(3), "DESTAIR"
Coll.Add Splited(4), "DESTCTY"

temp = SplitWithLengths(Splited(5), 6, 5, 2, 2, 5, 5, 5, 5)
    Coll.Add temp(0), "AIRCDE"
    Coll.Add temp(1), "FLTNBR"
    Coll.Add temp(2), "CLSSVC"
    Coll.Add temp(3), "CLSBKG"
    Coll.Add temp(4), "DEPDTE"
    Coll.Add temp(5), "DEPTIM"
    Coll.Add temp(6), "ARRTIM"
    Coll.Add temp(7), "ARRDTE"
    
Coll.Add Left(Splited(6), 2), "STATUS"
Coll.Add Splited(7), "PNRSTAT"
Coll.Add Splited(9), "NBRSTOP"
Coll.Add Splited(10), "EQUIP"
Coll.Add Splited(11), "ENTER"
Coll.Add Splited(12), "SHRCOMM"
Coll.Add Splited(13), "BAGALL"
Coll.Add Splited(14), "CKINTRM"
Coll.Add Splited(15), "CKINTIM"

Coll.Add Splited(16), "TKTT"
Coll.Add Splited(17), "FLTDUR"
Coll.Add Splited(18), "NONSMOK"
Coll.Add Splited(19), "MLTPM"
Coll.Add Splited(20), "CPNNR"
Coll.Add Splited(21), "CPNIN"
Coll.Add Splited(22), "CPNSN"
Coll.Add Splited(23), "CPNTSN"
Coll.Add Splited(8), "MEAL"

Set H = Coll
End Function



Private Function getLineInfo(DataLine As String, id As Long)
Dim head  As String

'----------GENERAL---------------------
head = Left(DataLine, 2)
Select Case head
    'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation
    'Case "A-", "B-", "C-", "D-", "E-", "F-", "G-", "H-", "I-", "J-", "K-", "L-", "M-", "N-", "O-", "P-", "Q-", "R-", "S-", "T-", "U-", "V-", "W-", "X-", "Y-", "Z-", "N-", "O-", "FP", "FM", "TK", "FV", "RM"
    'by Abhi on 18-Jul-2016 for caseid 6573 Amadeus -Ticket date picking logic for farelogix file
    'Case "A-", "B-", "C-", "D-", "E-", "F-", "G-", "H-", "I-", "J-", "K-", "L-", "M-", "N-", "O-", "P-", "Q-", "R-", "S-", "T-", "U-", "V-", "W-", "X-", "Y-", "Z-", "N-", "O-", "FP", "FM", "TK", "FV", "RM", "RC"
    Case "A-", "B-", "C-", "D-", "E-", "F-", "G-", "H-", "I-", "J-", "K-", "L-", "M-", "N-", "O-", "P-", "Q-", "R-", "S-", "T-", "U-", "V-", "W-", "X-", "Y-", "Z-", "N-", "O-", "FP", "FM", "TK", "FV", "RM", "RC", "FS"
    'by Abhi on 18-Jul-2016 for caseid 6573 Amadeus -Ticket date picking logic for farelogix file
    'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation
        id = 0
        getLineInfo = head
        Exit Function
End Select

'--------AIR------------------
head = Left(DataLine, 4)
Select Case head
    Case "AIR-"
        id = 1
        getLineInfo = head
        Exit Function
End Select
'---------3------------------
head = Left(DataLine, 3)
Select Case head
    'by Abhi on 12-Mar-2014 for caseid 3791 EMD number picking from amadeus file
    'Case "TAX", "KFT", "SSR", "OSI", "TSA", "KN-", "KNT", "KS-" ', "PEN"
    'by Abhi on 14-May-2014 for caseid 3787 Amadeus - Add the ATF value to Fare Amount
    'Case "TAX", "KFT", "SSR", "OSI", "TSA", "KN-", "KNT", "KS-", "TMC" ', "PEN"
    'by Abhi on 29-May-2014 for caseid 4104 EMD Value picking for Amadeus
    'Case "TAX", "KFT", "SSR", "OSI", "TSA", "KN-", "KNT", "KS-", "TMC", "ATF" ', "PEN"
    'by Abhi on 07-Aug-2017 for caseid 7674 Form of payment mapping for EMD tickets for Amadeus
    'Case "TAX", "KFT", "SSR", "OSI", "TSA", "KN-", "KNT", "KS-", "TMC", "ATF", "EMD" ', "PEN"
    'by Abhi on 25-Apr-2019 for caseid 10192 Amadeus Refund File reading
    'Case "TAX", "KFT", "SSR", "OSI", "TSA", "KN-", "KNT", "KS-", "TMC", "ATF", "EMD", "MFP" ', "PEN"
    Case "TAX", "KFT", "SSR", "OSI", "TSA", "KN-", "KNT", "KS-", "TMC", "ATF", "EMD", "MFP", "RFD" ', "PEN"
    'by Abhi on 25-Apr-2019 for caseid 10192 Amadeus Refund File reading
    'by Abhi on 07-Aug-2017 for caseid 7674 Form of payment mapping for EMD tickets for Amadeus
    'by Abhi on 29-May-2014 for caseid 4104 EMD Value picking for Amadeus
    'by Abhi on 14-May-2014 for caseid 3787 Amadeus - Add the ATF value to Fare Amount
    'by Abhi on 12-Mar-2014 for caseid 3791 EMD number picking from amadeus file
        id = 0
        getLineInfo = head
        Exit Function
End Select
'----------AMD-----------
head = Left(DataLine, 3)
Select Case head
    Case "AMD"
        id = -2
        getLineInfo = head
        Exit Function
End Select

id = -1

End Function


Private Function insertLineGeneral(DataLine As String, Ident As String, UploadNo As Long)
Dim head  As String
Dim Info As String
Dim temp
Dim Coll As New Collection
Dim CollPen As New Collection
Dim vi As Long

    head = Ident
    Info = Right(DataLine, Len(DataLine) - 2)
Select Case head
    Case "A-"
    Set Coll = A(Info)
        InsertData Coll, "AmdLineA", "A-", UploadNo
Case "B-"
        Set Coll = B(Info)
        InsertData Coll, "AmdLineB", "B-", UploadNo
Case "C-"
        Set Coll = C(Info)
        InsertData Coll, "AmdLineC", "C-", UploadNo
Case "D-"
        Set Coll = d(Info)
        InsertData Coll, "AmdLineD", "D-", UploadNo
Case "G-"
        Set Coll = G(Info)
        InsertData Coll, "AmdLineG", "G-", UploadNo
Case "H-"
        Set Coll = H(Info)
        InsertData Coll, "AmdLineH", "H-", UploadNo
Case "I-"
        Set Coll = i(Info)
        InsertDataByFieldName Coll, "AmdLineI", "I-", UploadNo
Case "K-"
        If Trim(Info) <> "" Then
            Set Coll = k(Info)
            InsertData Coll, "AmdLineK", "K-", UploadNo
        End If
Case "Q-"
        Set Coll = Q(Info)
        InsertData Coll, "AmdLineQ", "Q-", UploadNo
Case "T-"
        Set Coll = T(Info)
        InsertData Coll, "AmdLineT", "T-", UploadNo
Case "U-"
        temp = WhichUType(Info)
        If temp = "HHL" Then
            Set Coll = U_Hotel(Info)
            InsertDataByFieldName Coll, "AMDLineU_Hotel", "U-", UploadNo
        'by Abhi on 23-Jun-2010 for caseid 1401 Amadeus U-Hotel-HTL Segment
        ElseIf temp = "HTL" Then
            Set Coll = U_HotelHTLM(Info)
            InsertDataByFieldName Coll, "AMDLineU_HotelHTLM", "U-", UploadNo
        'by Abhi on 28-Nov-2014 for caseid 4776 Amadeus Car segments
        ElseIf temp = "CCR" Then
            Set Coll = U_Car(Info)
            InsertDataByFieldName Coll, "AmdLineU_Car", "U-", UploadNo
        'by Abhi on 28-Nov-2014 for caseid 4776 Amadeus Car segments
        Else
            Set Coll = U(Info)
            If Coll("STOP") = "O" Or Coll("STOP") = "X" Then
                InsertData Coll, "AmdLineH", "U-", UploadNo
            End If
        End If
Case "X-"
        'by Abhi on 23-Dec-2014 for caseid 4838 Amdline x is not coming to penair
        'Set Coll = x(Info)
        'InsertData Coll, "AmdLineX", "X-", UploadNo
        'by Abhi on 22-Jan-2015 for caseid 4917 disable the case id-4838 (Amdline x is not coming to penair )
        'by Abhi on 04-Feb-2015 for caseid 4879 Amadeus-Operational segment Identifier in Airsegdetails table
        Set Coll = X4H(Info)
        InsertData Coll, "AmdLineH", "X-", UploadNo
        'by Abhi on 04-Feb-2015 for caseid 4879 Amadeus-Operational segment Identifier in Airsegdetails table
        'by Abhi on 22-Jan-2015 for caseid 4917 disable the case id-4838 (Amdline x is not coming to penair )
        'by Abhi on 23-Dec-2014 for caseid 4838 Amdline x is not coming to penair
Case "Y-"
        Set Coll = y(Info)
        InsertData Coll, "AmdLineY", "Y-", UploadNo
Case "M-"
        Set Coll = M(Info)
        InsertDataCollection Coll, "AmdLineM", "M-", UploadNo
Case "N-"
        Set Coll = N(Info)
        InsertData Coll, "AmdLineN", "N-", UploadNo
Case "O-"
        Set Coll = O(Info)
        InsertData Coll, "AmdLineO", "O-", UploadNo

Case "FP"
        Set Coll = FP(Info)
        InsertData Coll, "AmdLineFP", "FP", UploadNo
Case "FM"
        Set Coll = FM(Info)
        InsertData Coll, "AmdLineFM", "FM", UploadNo
Case "TK"
        Set Coll = TK(Info)
        InsertDataByFieldName Coll, "AmdLineTK", "TK", UploadNo
Case "FV"
        Set Coll = FV(Info)
        InsertData Coll, "AmdLineFV", "FV", UploadNo
Case "RM"
        Set Coll = RM(Info)
        InsertData Coll("GEN"), "AmdLineRM", "RM", UploadNo
        Set CollPen = Coll("PEN")
        For vi = 1 To CollPen.Count
            InsertDataByFieldName CollPen(vi), "AmdLinePEN", "PEN", UploadNo
            DoEvents
        Next
        Set CollPen = Coll("PENPASSENGER")
        For vi = 1 To CollPen.Count
            InsertDataByFieldName CollPen(vi), "AmdLinePEN", "PEN", UploadNo
            DoEvents
        Next
        'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Amadeus
        Set CollPen = Coll("PENFARE")
        For vi = 1 To CollPen.Count
            InsertDataByFieldName CollPen(vi), "AmdLinePEN", "PENFARE", UploadNo
            DoEvents
        Next
        
        'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
        Set CollPen = Coll("PENLINK")
        If CollPen.Count > 0 Then
            InsertDataByFieldName CollPen(1), "AmdLinePEN", "PENLINK", UploadNo
        End If
        'by Abhi on 10-Aug-2010 for caseid 1433 Penline PENAUTOOFF
        Set CollPen = Coll("PENAUTOOFF")
        If CollPen.Count > 0 Then
            InsertDataByFieldName CollPen(1), "AmdLinePEN", "PENAUTOOFF", UploadNo
        End If
        'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
        Set CollPen = Coll("PENO")
        If CollPen.Count > 0 Then
            InsertDataByFieldName CollPen(1), "AmdLinePEN", "PENO", UploadNo
        End If
        'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
        Set CollPen = Coll("PENATOL")
        If CollPen.Count > 0 Then
            InsertDataByFieldName CollPen(1), "AmdLinePEN", "PENATOL", UploadNo
        End If
        'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
        Set CollPen = Coll("PENAGENTCOM")
        If CollPen.Count > 0 Then
            InsertDataByFieldName CollPen(1), "AmdLinePEN", "PENAGENTCOM", UploadNo
        End If
        'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
        'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
        Set CollPen = Coll("PENAIRTKT")
        If CollPen.Count > 0 Then
            InsertDataByFieldName CollPen(1), "AmdLinePEN", "PENAIRTKT", UploadNo
        End If
        'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
        'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
        Set CollPen = Coll("PENAGROSS")
        If CollPen.Count > 0 Then
            InsertDataByFieldName CollPen(1), "AmdLinePEN", "PENAGROSS", UploadNo
        End If
        'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
        'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
        Set CollPen = Coll("PENBILLCUR")
        If CollPen.Count > 0 Then
            InsertDataByFieldName CollPen(1), "AmdLinePEN", "PENBILLCUR", UploadNo
        End If
        'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        Set CollPen = Coll("PENWAIT")
        If CollPen.Count > 0 Then
            InsertDataByFieldName CollPen(1), "AmdLinePEN", "PENWAIT", UploadNo
        End If
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
        Set CollPen = Coll("PENVC")
        If CollPen.Count > 0 Then
            InsertDataByFieldName CollPen(1), "AmdLinePEN", "PENVC", UploadNo
        End If
        'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
        'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
        Set CollPen = Coll("PENCS")
        If CollPen.Count > 0 Then
            InsertDataByFieldName CollPen(1), "AmdLinePEN", "PENCS", UploadNo
        End If
        'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation
Case "RC"
        Set Coll = RC(Info)
        Set CollPen = Coll("MARKUP")
        For vi = 1 To CollPen.Count
            InsertDataByFieldName CollPen(vi), "AmdLinePEN", "MARKUP", UploadNo
            DoEvents
        Next
'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation
Case "TAX"
        Info = Right(DataLine, Len(DataLine) - 4)
        Set Coll = TAX(Info)
        InsertData Coll, "AmdLineTAX", "TAX-", UploadNo
Case "SSR"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = SSR(Info)
        InsertData Coll, "AmdLineSSR", "SSR", UploadNo
Case "KFT", "KNT"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = KFT(Info)
        InsertDataByFieldName Coll, "AmdLineKFT", head, UploadNo
Case "TSA"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = TSA(Info)
        InsertData Coll, "AmdLineTSA", "TSA", UploadNo
Case "OSI"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = OSI(Info)
        InsertData Coll, "AmdLineOSI", "OSI", UploadNo
Case "KN-"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = k(Info)
        InsertData Coll, "AmdLineK", "KN-", UploadNo
Case "KS-"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = KS(Info)
        InsertData Coll, "AmdLineKS", "KS-", UploadNo
'by Abhi on 12-Mar-2014 for caseid 3791 EMD number picking from amadeus file
Case "TMC"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = TMC(Info)
        InsertData Coll, "AmdLineTMC", "TMC", UploadNo
'by Abhi on 12-Mar-2014 for caseid 3791 EMD number picking from amadeus file
'by Abhi on 27-Mar-2014 for caseid 3853 Amadeus invoice number to reference field in folder
Case "V-"
        Set Coll = v(Info)
        InsertData Coll, "AmdLineV", "V-", UploadNo
'by Abhi on 27-Mar-2014 for caseid 3853 Amadeus invoice number to reference field in folder
'by Abhi on 14-May-2014 for caseid 3787 Amadeus - Add the ATF value to Fare Amount
Case "ATF"
        Info = Right(DataLine, Len(DataLine) - 4)
        Set Coll = Q_ATF(Info)
        InsertData Coll, "AmdLineQ_ATF", "ATF", UploadNo
'by Abhi on 14-May-2014 for caseid 3787 Amadeus - Add the ATF value to Fare Amount
'by Abhi on 29-May-2014 for caseid 4104 EMD Value picking for Amadeus
Case "EMD"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = EMD(Info)
        InsertData Coll, "AmdLineEMD", "EMD", UploadNo
'by Abhi on 29-May-2014 for caseid 4104 EMD Value picking for Amadeus
'Case "PEN"
'        Info = Right(DataLine, Len(DataLine) - 3)
'        Set Coll = PEN(Info)
'        InsertDataByFieldName Coll, "AmdLinePEN", "PEN", UploadNo
'by Abhi on 18-Jul-2016 for caseid 6573 Amadeus -Ticket date picking logic for farelogix file
Case "FS"
    Set Coll = FS(Info)
    InsertDataByFieldName Coll, "AmdLineFS", head, UploadNo
'by Abhi on 18-Jul-2016 for caseid 6573 Amadeus -Ticket date picking logic for farelogix file
'by Abhi on 07-Aug-2017 for caseid 7674 Form of payment mapping for EMD tickets for Amadeus
Case "MFP"
    Info = Right(DataLine, Len(DataLine) - 3)
    Set Coll = MFP(Info)
    InsertData Coll, "AmdLineMFP", "MFP", UploadNo
'by Abhi on 07-Aug-2017 for caseid 7674 Form of payment mapping for EMD tickets for Amadeus
'by Abhi on 25-Apr-2019 for caseid 10192 Amadeus Refund File reading
Case "RFD"
    Info = Right(DataLine, Len(DataLine) - 3)
    Set Coll = RFD(Info)
    InsertData Coll, "AmdLineRFD", "RFD", UploadNo
Case "R-"
    Info = Right(DataLine, Len(DataLine) - 2)
    Set Coll = R(Info)
    InsertData Coll, "AmdLineR", "R-", UploadNo
'by Abhi on 25-Apr-2019 for caseid 10192 Amadeus Refund File reading
End Select
End Function

Private Function x(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = SplitForce(Data, ";", 23)

Coll.Add Splited(0), "TTONBR"

temp = SplitWithLengthsPlus(Splited(1), 3, 1)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "STOP"
    Coll.Add temp(2), "ORIGAIR"
Coll.Add Splited(2), "ORIGCTY"
Coll.Add Splited(3), "DESTAIR"
Coll.Add Splited(4), "DESTCTY"
temp = SplitWithLengthsPlus(Splited(5), 6, 5, 2, 2, 5, 5, 5, 5)
    Coll.Add temp(0), "AIRCDE"
    Coll.Add temp(1), "FLTNBR"
    Coll.Add temp(2), "CLSSVC"
    Coll.Add temp(3), "CLSBKG"
    Coll.Add temp(4), "DEPDTE"
    Coll.Add temp(5), "DEPTIM"
    Coll.Add temp(6), "ARRTIM"
    Coll.Add temp(7), "ARRDTE"
    
Coll.Add Splited(6), "STATUS"
Coll.Add Splited(7), "PNRSTAT"
Coll.Add Splited(8), "MEAL"
Coll.Add Splited(9), "NBRSTOP"

Coll.Add Splited(10), "EQUIP"
Coll.Add Splited(11), "ENTER"
Coll.Add Splited(12), "SHRCOMM"
Coll.Add Splited(13), "BAGALL"
Coll.Add Splited(14), "CKINTRM"
Coll.Add Splited(15), "CKINTIM"
Coll.Add Splited(16), "TKTT"
Coll.Add Splited(17), "FLTDUR"
Coll.Add Splited(18), "NONSMOK"

Set x = Coll

End Function

'by Abhi on 23-Dec-2014 for caseid 4838 Amdline x is not coming to penair
Private Function X4H(Data As String) As Collection
'Dim Coll As New Collection
'Dim splited
'Dim temp, temp2
'Dim sTemp1, sTemp2, sTemp3
'splited = SplitForce(Data, ";", 23)
'
'Coll.Add splited(0), "TTONBR"
'
'temp = SplitWithLengthsPlus(splited(1), 3, 1)
'    Coll.Add temp(0), "SEGNBR"
'    Coll.Add temp(1), "STOP"
'    Coll.Add temp(2), "ORIGAIR"
'Coll.Add splited(2), "ORIGCTY"
'Coll.Add splited(3), "DESTAIR"
'Coll.Add splited(4), "DESTCTY"
'temp = SplitWithLengthsPlus(splited(5), 6, 5, 2, 2, 5, 5, 5, 5)
'    Coll.Add temp(0), "AIRCDE"
'    Coll.Add temp(1), "FLTNBR"
'    Coll.Add temp(2), "CLSSVC"
'    Coll.Add temp(3), "CLSBKG"
'    Coll.Add temp(4), "DEPDTE"
'    Coll.Add temp(5), "DEPTIM"
'    Coll.Add temp(6), "ARRTIM"
'    Coll.Add temp(7), "ARRDTE"
'
'Coll.Add Left(splited(6), 2), "STATUS"
'Coll.Add splited(7), "PNRSTAT"
'Coll.Add splited(9), "NBRSTOP"
'
'Coll.Add splited(10), "EQUIP"
'Coll.Add splited(11), "ENTER"
'Coll.Add splited(12), "SHRCOMM"
'Coll.Add splited(13), "BAGALL"
'Coll.Add splited(14), "CKINTRM"
'Coll.Add splited(15), "CKINTIM"
'Coll.Add splited(16), "TKTT"
'Coll.Add splited(17), "FLTDUR"
'Coll.Add splited(18), "NONSMOK"
'Coll.Add "", "MLTPM"
'Coll.Add "", "CPNNR"
'Coll.Add "", "CPNIN"
'Coll.Add "", "CPNSN"
'Coll.Add "", "CPNTSN"
'Coll.Add splited(8), "MEAL"
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = SplitForce(Data, ";", 24)

Coll.Add Splited(0), "TTONBR"

temp = SplitWithLengthsPlus(Splited(1), 3, 1)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "STOP"
    Coll.Add temp(2), "ORIGAIR"
Coll.Add Splited(2), "ORIGCTY"
Coll.Add Splited(3), "DESTAIR"
Coll.Add Splited(4), "DESTCTY"

temp = SplitWithLengths(Splited(5), 6, 5, 2, 2, 5, 5, 5, 5)
    Coll.Add temp(0), "AIRCDE"
    Coll.Add temp(1), "FLTNBR"
    Coll.Add temp(2), "CLSSVC"
    Coll.Add temp(3), "CLSBKG"
    Coll.Add temp(4), "DEPDTE"
    Coll.Add temp(5), "DEPTIM"
    Coll.Add temp(6), "ARRTIM"
    Coll.Add temp(7), "ARRDTE"
    
Coll.Add Left(Splited(6), 2), "STATUS"
Coll.Add Splited(7), "PNRSTAT"
Coll.Add Splited(9), "NBRSTOP"
Coll.Add Splited(10), "EQUIP"
Coll.Add Splited(11), "ENTER"
Coll.Add Splited(12), "SHRCOMM"
Coll.Add Splited(13), "BAGALL"
Coll.Add Splited(14), "CKINTRM"
Coll.Add Splited(15), "CKINTIM"

Coll.Add Splited(16), "TKTT"
Coll.Add Splited(17), "FLTDUR"
Coll.Add Splited(18), "NONSMOK"
Coll.Add "", "MLTPM"
Coll.Add "", "CPNNR"
Coll.Add "", "CPNIN"
Coll.Add "", "CPNSN"
Coll.Add "", "CPNTSN"
Coll.Add Splited(8), "MEAL"

Set X4H = Coll
End Function
'by Abhi on 23-Dec-2014 for caseid 4838 Amdline x is not coming to penair


Private Function y(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = SplitForce(Data, ";", 14)

Coll.Add Splited(0), "TTONBR"
    temp = SplitTwo(Splited(1), 3)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "ORIGAIR"
Coll.Add Splited(2), "ORIGCTY"
Coll.Add Splited(3), "DESTAIR"
Coll.Add Splited(4), "DESTCTY"
Coll.Add Splited(5), "OWNER"
Coll.Add Splited(6), "OWNNAME"
Coll.Add Splited(7), "COCKPIT"
Coll.Add Splited(8), "CPITNAME"
Coll.Add Splited(9), "CABIN"
Coll.Add Splited(10), "CABINNAME"
Coll.Add Splited(11), "EQUIP"
Coll.Add Splited(12), "GAUGE"

Set y = Coll
End Function


Private Function k(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = SplitForce(Data, ";", 15)
    
    temp = SplitWithLengths(Splited(0), 1, 3, 11)
    Coll.Add temp(0), "ISSUE"
    Coll.Add temp(1), "CURRTYP"
    Coll.Add temp(2), "BASE"
    
    temp = SplitWithLengths(Splited(1), 3, 11)
    Coll.Add temp(0), "CURCODE"
    Coll.Add temp(1), "EQVAMT"
    
    temp = SplitWithLengths(Splited(2), 10, 1, 3, 7, 3, 2)
    Coll.Add temp(0), "TAXFLD"
    Coll.Add temp(1), "OTAXI"
    Coll.Add temp(2), "TCURR"
    Coll.Add temp(3), "TAXA"
    Coll.Add temp(4), "TAXC"
    Coll.Add temp(5), "TAXN"
    
    'by Abhi on 04-Apr-2014 for caseid 3879 Amadeus Reissue file not picking value
    'temp = SplitWithLengths(splited(3), 3, 11)
    temp = SplitWithLengths(Splited(12), 3, 11)
    'by Abhi on 04-Apr-2014 for caseid 3879 Amadeus Reissue file not picking value
    Coll.Add temp(0), "CURRCDE"
    Coll.Add temp(1), "TOTAMT"

Coll.Add Splited(4), "SELLBUY1"
Coll.Add Splited(5), "SELLBUY2"
Coll.Add Splited(6), "TRANCURR"
'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue
Coll.Add vPSGRID_String, "PSGRID"
'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue

Set k = Coll
End Function


'Private Function KFT(Data As String) As Collection
'Dim coll As New Collection
'Dim Splited
'Dim temp, temp2, j
'Dim sTemp1, sTemp2, sTemp3
'
'Splited = Split(Data, ";")
'
'For j = 0 To UBound(Splited)
'If Len(Splited(j)) > 0 Then
'    coll.Add KFT_Sub(Splited(j))
'End If
'Next
'
'Set KFT = coll
'End Function
'

Private Function KFT(Data) As Collection
Dim Coll As New Collection
Dim temp, temp2, temp3
Dim TAXA_Total As Currency

Dim aKFT(5)
Dim aKFTUB(4)
Dim vi_Long As Long

    'temp = SplitForce(Data, ";", 6) '11
    temp = SplitForce(Data, ";", 30)
    
    Coll.Add temp(0), "ISSUE"
    
    aKFT(1) = SplitWithLengths(temp(1), 4, 9, 3, 2)
    aKFT(2) = SplitWithLengths(temp(2), 4, 9, 3, 2)
    aKFT(3) = SplitWithLengths(temp(3), 4, 9, 3, 2)
    aKFT(4) = SplitWithLengths(temp(4), 4, 9, 3, 2)
    aKFT(5) = SplitWithLengths(temp(5), 4, 9, 3, 2)
            
    For vi_Long = 2 To 5
        If Right(Trim(aKFT(vi_Long)(2)), 2) = "UB" Then
            aKFTUB(0) = aKFT(1)(0)
            aKFTUB(1) = aKFT(1)(1)
            aKFTUB(2) = aKFT(1)(2)
            aKFTUB(3) = aKFT(1)(3)
            
            aKFT(1)(0) = aKFT(vi_Long)(0)
            aKFT(1)(1) = aKFT(vi_Long)(1)
            aKFT(1)(2) = aKFT(vi_Long)(2)
            aKFT(1)(3) = aKFT(vi_Long)(3)
            
            aKFT(vi_Long)(0) = aKFTUB(0)
            aKFT(vi_Long)(1) = aKFTUB(1)
            aKFT(vi_Long)(2) = aKFTUB(2)
            aKFT(vi_Long)(3) = aKFTUB(3)
        End If
        DoEvents
    Next
        Coll.Add aKFT(1)(0), "TCURR1"
        Coll.Add aKFT(1)(1), "TAXA1"
        Coll.Add aKFT(1)(2), "TAXC1"
        Coll.Add aKFT(1)(3), "TAXN1"

        Coll.Add aKFT(2)(0), "TCURR2"
        Coll.Add aKFT(2)(1), "TAXA2"
        Coll.Add aKFT(2)(2), "TAXC2"
        Coll.Add aKFT(2)(3), "TAXN2"

        Coll.Add aKFT(3)(0), "TCURR3"
        Coll.Add aKFT(3)(1), "TAXA3"
        Coll.Add aKFT(3)(2), "TAXC3"
        Coll.Add aKFT(3)(3), "TAXN3"

        Coll.Add aKFT(4)(0), "TCURR4"
        Coll.Add aKFT(4)(1), "TAXA4"
        Coll.Add aKFT(4)(2), "TAXC4"
        Coll.Add aKFT(4)(3), "TAXN4"

        Coll.Add aKFT(5)(0), "TCURR5"
        Coll.Add aKFT(5)(1), "TAXA5"
        Coll.Add aKFT(5)(2), "TAXC5"
        Coll.Add aKFT(5)(3), "TAXN5"
    
    temp2 = SplitWithLengths(temp(6), 4, 9, 3, 2)
        TAXA_Total = Val(TAXA_Total) + Val(temp2(1))
            For vi_Long = 7 To 30
                temp3 = SplitWithLengths(temp(vi_Long), 4, 9, 3, 2)
                    TAXA_Total = Val(TAXA_Total) + Val(temp3(1))
                DoEvents
            Next
            'temp3 = SplitWithLengths(temp(8), 4, 9, 3, 2)
            '    TAXA_Total = Val(TAXA_Total) + Val(temp3(1))
            'temp3 = SplitWithLengths(temp(9), 4, 9, 3, 2)
            '    TAXA_Total = Val(TAXA_Total) + Val(temp3(1))
            'temp3 = SplitWithLengths(temp(10), 4, 9, 3, 2)
            '    TAXA_Total = Val(TAXA_Total) + Val(temp3(1))
            'temp3 = SplitWithLengths(temp(11), 4, 9, 3, 2)
            '    TAXA_Total = Val(TAXA_Total) + Val(temp3(1))
        Coll.Add temp2(0), "TCURR6"
        Coll.Add CStr(TAXA_Total), "TAXA6"
        Coll.Add temp2(2), "TAXC6"
        Coll.Add temp2(3), "TAXN6"
        
    Coll.Add Data, "FAREREMARKS"
    'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue
    Coll.Add vPSGRID_String, "PSGRID"
    'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue
Set KFT = Coll
End Function

Private Function TAX(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Splited = SplitForce(Data, ";", 3)

For j = 0 To UBound(Splited)
    TAX_Sub Splited(j), j, Coll
    DoEvents
Next
Set TAX = Coll
End Function


Private Function TAX_Sub(Data, Index, Coll As Collection) As Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

If Len(Data) = 24 Then
    temp = SplitWithLengths(Data, 3, 3, 9, 3)
    Coll.Add temp(0), "TAXFLD" & Index
    Coll.Add temp(1), "TCURR" & Index
    Coll.Add temp(2), "TAXA" & Index
    Coll.Add temp(3), "TAXC" & Index
Else
    temp = SplitWithLengths(Data, 3, 9, 3)
    Coll.Add "", "TAXFLD" & Index
    Coll.Add temp(0), "TCURR" & Index
    Coll.Add temp(1), "TAXA" & Index
    Coll.Add temp(2), "TAXC" & Index
End If
    
Set TAX_Sub = Coll
End Function




Private Function U___2(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Splited = SplitForce(Data, ";", 13)
    temp = SplitWithLengths(Splited(0), 3, 1)
    Coll.Add temp(0), "TTONBR"
    Coll.Add temp(1), "DATAMODI"
    
    temp = SplitWithLengthsPlus(Splited(1), 3, 1)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "STOP"
    Coll.Add temp(2), "ORIGAIR"
    
Coll.Add Splited(2), "ORIGCTY"
Coll.Add Splited(3), "DESTAIR"
Coll.Add Splited(4), "DESTCTY"
Coll.Add Splited(5), "HVOID"
    
Set U___2 = Coll
End Function



Private Function InsertData(Data As Collection, TableName As String, LineIdent As String, UploadNo As Long) As Boolean
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
On Error GoTo PENErr
Dim ErrNumber As String
Dim ErrDescription As String
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'On Error GoTo errPara
Dim rs As New ADODB.Recordset
Dim j As Integer
'by Abhi on 23-Oct-2010 for caseid 1516 PenGDS Amadeus slow reading
'Rs.Open "Select * from " + TableName, dbCompany, adOpenDynamic, adLockPessimistic
'by Abhi on 12-Nov-2010 for caseid 1546 PenGDS Optimistic concurrency check failed
'Rs.Open "Select * from " + TableName & " WHERE UpLoadNo=" & UploadNo, dbCompany, adOpenForwardOnly, adLockOptimistic
rs.Open "Select * from " + TableName & " WHERE UpLoadNo=" & UploadNo, dbCompany, adOpenForwardOnly, adLockPessimistic
rs.AddNew
rs.Fields(0) = UploadNo
rs.Fields(1) = LineIdent
For j = 2 To rs.Fields.Count - 1
    rs.Fields(j).value = Data(j - 1)
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    'If Len(Data(j - 1)) > rs.Fields(j).DefinedSize Then
    '    ErrDetails_String = vbCrLf & "Table=" & TableName & ", Field=" & rs.Fields(j).Name & ", FieldSize=" & rs.Fields(j).DefinedSize & "<>" & Len(Data(j - 1)) & "," & vbCrLf & Data(j - 1) & vbCrLf
    'End If
    'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    If Len(Data(j - 1)) > rs.Fields(j).DefinedSize Then
        ErrDetails_String = vbCrLf & "MSTAG=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(j).Name & ", FieldSize=" & rs.Fields(j).DefinedSize & "<>" & Len(Data(j - 1)) & "," & vbCrLf & Data(j - 1) & vbCrLf
    End If
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    DoEvents
Next
rs.Update
Exit Function
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'errPara:
'MsgBox Err.Description & vbCrLf & Rs.Fields(j).Name & " : (" & Rs.Fields(j).DefinedSize & ") at " & j & " Column in Table '" & TableName & "'", vbCritical, "Error During Updation"
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
PENErr:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    'If ErrNumber = ErrNumber Then 'Subscript out of range
    If KeyExists(Data, j - 1) = False Then
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
        ErrDetails_String = vbCrLf & "MSTAG=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(j).Name & ", Field does not exist in Collection!," & vbCrLf
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    ElseIf Trim(ErrDetails_String) <> "" Then '-2147467259 Data provider or other service returned an E_FAIL status.
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    ElseIf Len(Data(j - 1)) > rs.Fields(j).DefinedSize Then '-2147217887 Multiple-step operation generated errors. Check each status value.
        ErrDetails_String = vbCrLf & "MSTAG=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(j).Name & ", FieldSize=" & rs.Fields(j).DefinedSize & "<>" & Len(Data(j - 1)) & "," & vbCrLf & Data(j - 1) & vbCrLf
    Else
        ErrDetails_String = vbCrLf & "MSTAG=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(j).Name & "," & vbCrLf & Data(j - 1) & vbCrLf
    End If
    Err.Raise ErrNumber, , ErrDescription
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
End Function

Private Function InsertDataCollection(Data As Collection, TableName As String, LineIdent As String, UploadNo As Long) As Boolean
For j = 1 To Data.Count
    InsertData Data(j), TableName, LineIdent, UploadNo
    DoEvents
Next
End Function


Private Function M_Sub(Data) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
    temp = SplitTwoReverse(Data, 15)
    Coll.Add temp(0), "FARELEM"
    temp = SplitWithLengths(temp(1), 3, 6, 6)
    Coll.Add temp(0), "PRIMCDE"
    Coll.Add temp(1), "FAREB"
    Coll.Add temp(2), "TKTDSG"
Set M_Sub = Coll
End Function

Private Function M(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = Split(Data, ";")

For j = 0 To UBound(Splited)
If Len(Splited(j)) > 0 Then
    Coll.Add M_Sub(Splited(j))
End If
DoEvents
Next

Set M = Coll
End Function

Private Function U(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = SplitForce(Data, ";", 22)

    temp = SplitTwo(Splited(0), 3)
    Coll.Add temp(0), "TTONBR"
    'Coll.Add temp(1), "DATAMODI"

    temp = SplitWithLengthsPlus(Splited(1), 3, 1)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "STOP"
    'if temp(1)="M" then
    Coll.Add temp(2), "ORIGAIR"
    
Coll.Add Splited(2), "ORIGCTY"
Coll.Add Splited(3), "DESTAIR"
Coll.Add Splited(4), "DESTCTY"

    temp = SplitWithLengthsPlus(Splited(5), 6, 5, 2, 2, 5, 5, 5, 5)
    Coll.Add temp(0), "AIRCDE"
    Coll.Add temp(1), "FLTNBR"
    Coll.Add temp(2), "CLSSVC"
    Coll.Add temp(3), "CLSBKG"
    Coll.Add temp(4), "DEPDTE"
    Coll.Add temp(5), "DEPTIM"
    Coll.Add temp(6), "ARRTIM"
    Coll.Add temp(7), "ARRDTE"

Coll.Add Left(Splited(6), 2), "STATUS"

Coll.Add Splited(7), "PNRSTAT"
Coll.Add Splited(9), "NBRSTOP"
Coll.Add Splited(10), "EQUIP"
Coll.Add Splited(11), "ENTER"
Coll.Add Splited(12), "SHRCOMM"
Coll.Add Splited(13), "BAGALL"
Coll.Add Splited(14), "CKINTRM"
Coll.Add Splited(15), "CKINTIM"
Coll.Add Splited(16), "TKTT"
Coll.Add Splited(17), "FLTDUR"
Coll.Add Splited(18), "NONSMOK"
'Coll.Add splited(19), "TTAG"
Coll.Add Splited(19), "MLTPM"
Coll.Add "", "CPNNR"
Coll.Add "", "CPNIN"
Coll.Add "", "CPNSN"
Coll.Add "", "CPNTSN"
Coll.Add Splited(8), "MEAL"

'Coll.Add splited(21), "FLNTAG"


Set U = Coll
End Function



Private Function i(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2, temp3, temp4
Dim sTemp1, sTemp2, sTemp3

Splited = SplitForce(Data, ";", 13)
Coll.Add Splited(0), "GCPSGR"
    temp = SplitWithLengthsPlus(Splited(1), 2)
    Coll.Add temp(0), "PSGRNBR"
    TKT_NoIndex = temp(0)
    temp2 = SplitField(temp(1), "(", ")")
    
    temp3 = SplitForce(temp2(0), "/", 2)
        Coll.Add temp3(0), "PSGRLNME"
    
    'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
    'temp4 = mSplitTypeAndSirName(temp3(1))
    temp4 = FindInitialAndName(temp3(1))
    'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
        Coll.Add temp4(0), "PSGRFNME"
        Coll.Add temp4(1), "PSGRTITLE"
        'by Abhi on 18-Mar-2014 for caseid 3814 Amadeus Passenger type is picking from the wrong position of gds file
        'If temp2(1) <> "" Then
        '    'by Abhi on 08-Jul-2010 for caseid 1413 pax type picking issue for amadeus
        '    If temp2(2) = "" Then
        '        Coll.Add temp2(1), "PSGRID"
        '    Else
        '        'by Abhi on 08-Jul-2010 for caseid 1413 pax type picking issue for amadeus
        '        temp2(2) = Replace(temp2(2), "(", "")
        '        temp2(2) = Replace(temp2(2), ")", "")
        '        Coll.Add temp2(2), "PSGRID"
        '    End If
        'Else
        '    Coll.Add temp4(2), "PSGRID"
        'End If
        'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue
        'If Left(temp2(1), 2) <> "ID" Then
        '    Coll.Add temp2(1), "PSGRID"
        'Else
        '    Coll.Add "", "PSGRID"
        'End If
        
        If Left(temp2(1), 2) <> "ID" Then
            'by Abhi on 18-Feb-2016 for caseid 5998 Amadeus Pax type picking logic
            'vPSGRID_String = temp2(1)
            vPSGRID_String = Left(temp2(1), 3)
            'by Abhi on 18-Feb-2016 for caseid 5998 Amadeus Pax type picking logic
        Else
            vPSGRID_String = ""
        End If
        If Trim(vPSGRID_String) = "" Then
            vPSGRID_String = "ADT"
        End If
        'by Abhi on 16-May-2017 for caseid 7439 Amadeus fare picking issue due to invalid value in Amdline1 and AmdlineK
        vPSGRID_String = GetPassengerTypefromShort(vPSGRID_String)
        'by Abhi on 16-May-2017 for caseid 7439 Amadeus fare picking issue due to invalid value in Amdline1 and AmdlineK
        
        Coll.Add vPSGRID_String, "PSGRID"
        'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue
        'by Abhi on 18-Mar-2014 for caseid 3814 Amadeus Passenger type is picking from the wrong position of gds file
    'by Abhi on 11-Jun-2014 for caseid 4138 Amadeus picking the details (email address, telephone) from I-00 line
    sTemp1 = ExtractBetween(Splited(3), "E-", "//")
        'by Abhi on 02-Oct-2017 for caseid 7925 GDS Upload
        sTemp1 = Replace(sTemp1, " AT ", "@", , , vbTextCompare)
        'by Abhi on 02-Oct-2017 for caseid 7925 GDS Upload
        Coll.Add sTemp1, "EMAIL"
    sTemp1 = ExtractBetween(Splited(3), "M-", "//")
        If Trim(sTemp1) = "" Then
            sTemp1 = ExtractBetween(Splited(3), "H-", "//")
        End If
        Coll.Add sTemp1, "PHONE"
    'by Abhi on 11-Jun-2014 for caseid 4138 Amadeus picking the details (email address, telephone) from I-00 line

'    temp3 = temp2(2)
'    temp = SplitField(temp3, "(", ")")
'    temp2 = SplitTwo(temp(1), 2)
'
'    Coll.Add temp2(0), "IDFLD"
'    Coll.Add temp2(1), "IDNBR"
'
'Coll.Add splited(2), "CRRMK"
'Coll.Add splited(3), "APINFO"
'Coll.Add splited(4), "TRARL"
'Coll.Add splited(5), "COMRL"
Set i = Coll
End Function



Private Function SSR(Data As String) As Collection

Dim Coll As New Collection
Dim Splited
Dim temp, temp2, temp3, temp4
Dim sTemp1, sTemp2, sTemp3
Splited = SplitForce(Data, ";S", 2)

    temp = SplitFirstTwo(Splited(0), "/")
    temp2 = SplitWithLengthsPlus(temp(0), 1, 4, 1, 3, 1, 2)
    Coll.Add temp2(1), "SSRTYPE"
    Coll.Add temp2(3), "AIRCDE"
    Coll.Add temp2(4), "STATCDE"
    Coll.Add temp2(5), "NBRPTY"
    Coll.Add temp(1), "FREE"
    
    temp = SplitForce(Splited(1), ";P", 2)
    temp2 = SplitForce(temp(0), "-", 2)
    temp3 = SplitTwoReverse(temp2(0), 3)
    Coll.Add temp3(0), "SEGASS"
    Coll.Add temp3(1), "SEGNBR"
    temp4 = SplitForce(temp2(1), ",", 2)
    Coll.Add temp4(0), "SEGNBR2"
    
    temp2 = SplitFirstTwo(temp(1), "-")
    temp3 = SplitTwoReverse(temp2(0), 2)
    Coll.Add temp3(0), "PSGASS"
    Coll.Add temp3(1), "PSGNBR"
    temp4 = SplitForce(temp2(1), ",", 2)
    Coll.Add temp4(0), "PSGNBR2"
    
Set SSR = Coll

End Function


Private Function Q(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

temp = SplitForce(Data, ";", 2)
    Coll.Add temp(0), "FARE"
    Coll.Add temp(1), "PRCENTRY"
Set Q = Coll
End Function



Private Function T(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Splited = SplitFirstTwo(Data, "/")
    temp = SplitForce(Splited(0), "-", 3)
    temp2 = SplitWithLengths(temp(0), 1, 3)
    Coll.Add temp2(0), "TKTTYPE"
    Coll.Add temp2(1), "NUMAIR"
    
    Coll.Add temp(1), "TKTNBR"
    Coll.Add temp(2), "DIGIT"
Coll.Add Splited(1), "FREE"
Coll.Add TKT_NoIndex, "SLNO"
    
Set T = Coll
End Function


'Private Function O(data As String) As Collection
'Dim coll As New Collection
'Dim Splited
'Dim temp, temp2
'Dim sTemp1, sTemp2, sTemp3
'
'Splited = SplitFirstTwo(data, "/")
'    temp = SplitForce(Splited(0), "-", 3)
'    temp2 = SplitWithLengths(temp(0), 1, 3)
'    coll.Add temp2(0), "TKTTYPE"
'    coll.Add temp2(1), "NUMAIR"
'
'    coll.Add temp(1), "TKTNBR"
'    coll.Add temp(2), "DIGIT"
'coll.Add Splited(1), "FREE"
'
'Set O = coll
'End Function




Private Function FP(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Splited = SplitFirstTwo(Data, ";S")
    Coll.Add Splited(0), "PAYMNT"
Splited = SplitFirstTwo(Splited(1), ";P")
    temp = SplitFirstTwo(Splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"
temp = SplitFirstTwo(Splited(1), "-")
temp2 = SplitTwoReverse(temp(0), 2)
    Coll.Add temp2(0), "PSGASS"
    Coll.Add temp2(1), "PSGNBR"
temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "PSGNBR2"
Set FP = Coll
End Function



Private Function FM(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Splited = SplitFirstTwo(Data, ";S")
    Coll.Add Splited(0), "FARECOM"
Splited = SplitFirstTwo(Splited(1), ";P")
    temp = SplitFirstTwo(Splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"

    temp = SplitFirstTwo(Splited(1), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "PSGASS"
    Coll.Add temp2(1), "PSGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "PSGNBR2"
    Coll.Add TKT_NoIndex, "SLNO"
Set FM = Coll
End Function

Private Function AddFile(FileName As String)
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
On Error GoTo PENErr
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
Dim id As Long
Dim head As String
Dim temp
Dim AMDLINE As String
Dim AMDCompleted As Boolean
AMDCompleted = True
Dim UploadNo As Long
Dim LineStr  As String
Dim tsObj As TextStream
Dim FileObj As New FileSystemObject
'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
Dim vSQL_String As String
'by Abhi on 17-Apr-2014 for caseid 3934 GDS Amadeus:If the tag AmdHeadAMD.ACTTAG contains the value "VOID" can we skip it to show from intray
Dim Splited
Dim vValidated_Boolean As Boolean
'by Abhi on 17-Apr-2014 for caseid 3934 GDS Amadeus:If the tag AmdHeadAMD.ACTTAG contains the value "VOID" can we skip it to show from intray
'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
Dim vDateLastModified_String As String
'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
'by Abhi on 19-Oct-2016 for caseid 6820 Amadues -German File reading for GDS with m as spliter
Dim vReadLine_String As String
Dim Lines
Dim j As Long
'by Abhi on 19-Oct-2016 for caseid 6820 Amadues -German File reading for GDS with m as spliter
'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
Dim vFileDateCreated_String As String
'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
Dim vOutOfBookings_Boolean As Boolean
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date

    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    GDSDeadlockRETRY_Integer = 0
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
DeadlockRETRY:
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    TKT_NoIndex = "00"   ' for Link Between I, T, TK, FM
    'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
    LUFPNR_String = ""
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    NoofPermissionDenied = 0
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    'by Abhi on 21-Apr-2010 for caseid 1320 NOFOLDER for in penline
    vValidated_Boolean = True
    'by Abhi on 01-Aug-2014 for caseid 4383 Amdeus Fare line issue
    vPSGRID_String = ""
    'by Abhi on 01-Aug-2014 for caseid 4383 Amdeus Fare line issue
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    GIT_CUSTOMERUSERCODE_String = ""
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
    GIT_PENWAIT_String = "N"
    'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
    
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
        PENFAREPNO_Long = 1
        'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue
        vPSGRID_String = ""
        'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue
        Do While tsObj.AtEndOfStream <> True
            DoEvents
            'by Abhi on 19-Oct-2016 for caseid 6820 Amadues -German File reading for GDS with m as spliter
            'LineStr = tsObj.ReadLine
            vReadLine_String = tsObj.ReadLine
            Lines = Split(vReadLine_String, "m")
            For j = 0 To UBound(Lines)
                'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
                fPNRDTE_String = ""
                'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
                LineStr = Lines(j)
            'by Abhi on 19-Oct-2016 for caseid 6820 Amadues -German File reading for GDS with m as spliter
                head = getLineInfo(LineStr, id)
                'by Abhi on 17-Apr-2014 for caseid 3934 GDS Amadeus:If the tag AmdHeadAMD.ACTTAG contains the value "VOID" can we skip it to show from intray
                If Trim(head) = "AMD" Then
                    Splited = SplitForce(LineStr, ";", 4)
                    If Left(Splited(2), 4) = "VOID" Then
                        vValidated_Boolean = False
                        Exit Do
                    End If
                End If
                'by Abhi on 17-Apr-2014 for caseid 3934 GDS Amadeus:If the tag AmdHeadAMD.ACTTAG contains the value "VOID" can we skip it to show from intray
                FMain.stbUpload.Panels(1).Text = "Reading.... " & head
                Select Case id
                Case 0
                    insertLineGeneral LineStr, head, UploadNo
                    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
                    If Trim(head) = "D-" Then
                        If Trim(FMain.OutOfBookingsDateAmadeusText.Text) <> "" Then
                            If Len(Trim(fPNRDTE_String)) > 0 Then
                                If CDate(DateFormat(fPNRDTE_String)) < CDate(DateFormat(FMain.OutOfBookingsDateAmadeusText.Text)) Then
                                    vOutOfBookings_Boolean = True
                                    vValidated_Boolean = False
                                    Exit Do
                                End If
                            End If
                        End If
                    End If
                    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
                    
                    If AMDCompleted = False Then
                        temp = SplitTwo(AMDLINE, 3)
                        InsertAMD temp(1), UploadNo, GetFileNameFromPath(FileName)
                        AMDCompleted = True
                    End If
                Case 1
                    InserAirMain LineStr, UploadNo
                Case -2
                    AMDLINE = LineStr
                    AMDCompleted = False
                Case -1
                    'by Abhi on 01-Aug-2014 for caseid 4383 Amdeus Fare line issue
                    If Left(LineStr, 3) = "END" Then
                        vSQL_String = "" _
                            & "Update dbo.AmdLineK " _
                            & "SET              PSGRID = '" & SkipChars(vPSGRID_String) & "' " _
                            & "WHERE     (PSGRID = '') AND (UpLoadNo = " & UploadNo & ") "
                        dbCompany.Execute vSQL_String
                        vSQL_String = "" _
                            & "Update dbo.AmdLineKFT " _
                            & "SET              PSGRID = '" & SkipChars(vPSGRID_String) & "' " _
                            & "WHERE     (PSGRID = '') AND (UpLoadNo = " & UploadNo & ") "
                        dbCompany.Execute vSQL_String
                        vSQL_String = "" _
                            & "Update dbo.AmdLineKS " _
                            & "SET              PSGRID = '" & SkipChars(vPSGRID_String) & "' " _
                            & "WHERE     (PSGRID = '') AND (UpLoadNo = " & UploadNo & ") "
                        dbCompany.Execute vSQL_String
                        
                        vPSGRID_String = ""
                    End If
                    'by Abhi on 01-Aug-2014 for caseid 4383 Amdeus Fare line issue
                    AMDLINE = AMDLINE & Chr(13) & Chr(10) & LineStr
                End Select
            'by Abhi on 19-Oct-2016 for caseid 6820 Amadues -German File reading for GDS with m as spliter
                DoEvents
            Next
            'by Abhi on 19-Oct-2016 for caseid 6820 Amadues -German File reading for GDS with m as spliter
            DoEvents
        Loop
        tsObj.Close
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        
        FMain.stbUpload.Panels(1).Text = "Writing GDS In Tray..."
        DoEvents
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'vSQL_String = "" _
        '    & "INSERT INTO dbo.GDSINTRAYTABLE " _
        '    & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC]) " _
        '    & "SELECT     'AIR' AS GIT_ID, 'AMADEUS' AS GIT_GDS, 'AMADEUS' AS GIT_INTERFACE, dbo.AmdHeadAMD.UploadNo AS GIT_UPLOADNO, '' AS GIT_PNRDATE, " _
        '    & "                      dbo.AmdHeadAMD.RECLOC AS GIT_PNR, dbo.AmdLineI.PSGRLNME AS GIT_LASTNAME, dbo.AmdLineI.PSGRFNME AS GIT_FIRSTNAME, " _
        '    & "                      dbo.AmdLineC.BKAGT AS GIT_BOOKINGAGENT, dbo.AmdLineT.TKTNBR AS GIT_TICKETNUMBER, dbo.AmdHeadAMD.FName AS GIT_FILENAME, " _
        '    & "                      dbo.AmdHeadAMD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.AmdHeadAMD.IDBKG AS GIT_PCC " _
        '    & "FROM         dbo.AmdHeadAMD INNER JOIN " _
        '    & "                      dbo.AmdLineI ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineI.UpLoadNo LEFT OUTER JOIN " _
        '    & "                      dbo.AmdLineT ON dbo.AmdLineI.PSGRNBR = dbo.AmdLineT.SLNO AND dbo.AmdLineI.UpLoadNo = dbo.AmdLineT.UpLoadNo LEFT OUTER JOIN " _
        '    & "                      dbo.AmdLineC ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineC.UpLoadNo " _
        '    & "Where (dbo.AmdHeadAMD.UploadNo = " & UploadNo & ")"
        If vValidated_Boolean = True Then
            vDateLastModified_String = fsObj.GetFile(FileName).DateLastModified
            'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
            vDateLastModified_String = DateTime12hrsFormat(vDateLastModified_String)
            vFileDateCreated_String = fsObj.GetFile(FileName).DateCreated
            vFileDateCreated_String = DateTime12hrsFormat(vFileDateCreated_String)
            'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
            'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
            vSQL_String = "" _
                & "SELECT     TOP (1) BRANCH " _
                & "FROM         dbo.AmdLinePEN " _
                & "WHERE     (UpLoadNo = " & UploadNo & ") AND (MSGTAG = 'PEN') AND (BRANCH <> '')"
            GIT_PENLINEBRID_String = getFromExecuted(vSQL_String, "BRANCH")
            'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
        End If
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        ErrDetails_String = " GDSINTRAYTABLE 1."
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
        vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC],[GIT_FILEISSUEDATE],[GIT_CUSTOMERUSERCODE],[GIT_PENWAIT],[GIT_FILECREATEDDATE],[GIT_ISSUEDBY], [GIT_PENLINEBRID]) " _
            & "SELECT     'AIR' AS GIT_ID, 'AMADEUS' AS GIT_GDS, 'AMADEUS' AS GIT_INTERFACE, dbo.AmdHeadAMD.UploadNo AS GIT_UPLOADNO, '' AS GIT_PNRDATE, " _
            & "                      dbo.AmdHeadAMD.RECLOC AS GIT_PNR, dbo.AmdLineI.PSGRLNME AS GIT_LASTNAME, dbo.AmdLineI.PSGRFNME AS GIT_FIRSTNAME, " _
            & "                      dbo.AmdLineC.BKAGT AS GIT_BOOKINGAGENT, dbo.AmdLineT.TKTNBR AS GIT_TICKETNUMBER, dbo.AmdHeadAMD.FName AS GIT_FILENAME, " _
            & "                      dbo.AmdHeadAMD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.AmdHeadAMD.IDBKG AS GIT_PCC, '" & vDateLastModified_String & "' AS GIT_FILEISSUEDATE, " _
            & "                      '" & SkipChars(GIT_CUSTOMERUSERCODE_String) & "' AS GIT_CUSTOMERUSERCODE, '" & SkipChars(GIT_PENWAIT_String) & "' AS GIT_PENWAIT, " _
            & "                      '" & vFileDateCreated_String & "' AS GIT_FILECREATEDDATE, dbo.AmdLineC.TKTAGT AS GIT_ISSUEDBY, '" & GIT_PENLINEBRID_String & "' AS GIT_PENLINEBRID " _
            & "FROM         dbo.AmdHeadAMD INNER JOIN " _
            & "                      dbo.AmdLineI ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineI.UpLoadNo LEFT OUTER JOIN " _
            & "                      dbo.AmdLineT ON dbo.AmdLineI.PSGRNBR = dbo.AmdLineT.SLNO AND dbo.AmdLineI.UpLoadNo = dbo.AmdLineT.UpLoadNo LEFT OUTER JOIN " _
            & "                      dbo.AmdLineC ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineC.UpLoadNo " _
            & "Where (dbo.AmdHeadAMD.UploadNo = " & UploadNo & ")"
        'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'by Abhi on 17-Apr-2014 for caseid 3934 GDS Amadeus:If the tag AmdHeadAMD.ACTTAG contains the value "VOID" can we skip it to show from intray
        'dbCompany.Execute vSQL_String
        If vValidated_Boolean = True Then
            dbCompany.Execute vSQL_String
        End If
        'by Abhi on 17-Apr-2014 for caseid 3934 GDS Amadeus:If the tag AmdHeadAMD.ACTTAG contains the value "VOID" can we skip it to show from intray
        'by Abhi on 12-Mar-2014 for caseid 3791 EMD number picking from amadeus file
        'by Abhi on 17-Mar-2014 for caseid 3812 Amadeus duplicated GDS In Tray
        'vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC]) " _
            & "SELECT     'AIR' AS GIT_ID, 'AMADEUS' AS GIT_GDS, 'AMADEUS' AS GIT_INTERFACE, dbo.AmdHeadAMD.UploadNo AS GIT_UPLOADNO, '' AS GIT_PNRDATE, " _
            & "                      dbo.AmdHeadAMD.RECLOC AS GIT_PNR, dbo.AmdLineI.PSGRLNME AS GIT_LASTNAME, dbo.AmdLineI.PSGRFNME AS GIT_FIRSTNAME, " _
            & "                      dbo.AmdLineC.BKAGT AS GIT_BOOKINGAGENT, dbo.AmdLineTMC.TKTNBR AS GIT_TICKETNUMBER, dbo.AmdHeadAMD.FName AS GIT_FILENAME, " _
            & "                      dbo.AmdHeadAMD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.AmdHeadAMD.IDBKG AS GIT_PCC " _
            & "FROM         dbo.AmdHeadAMD INNER JOIN " _
            & "                      dbo.AmdLineI ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineI.UpLoadNo LEFT OUTER JOIN " _
            & "                      dbo.AmdLineTMC ON dbo.AmdLineI.UpLoadNo = dbo.AmdLineTMC.UpLoadNo LEFT OUTER JOIN " _
            & "                      dbo.AmdLineC ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineC.UpLoadNo " _
            & "Where (dbo.AmdHeadAMD.UploadNo = " & UploadNo & ")"
        'by Abhi on 18-Dec-2015 for caseid 5894 Amadeus EMD ticket is duplicated in GDS intray
        'vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC]) " _
            & "SELECT     'AIR' AS GIT_ID, 'AMADEUS' AS GIT_GDS, 'AMADEUS' AS GIT_INTERFACE, dbo.AmdHeadAMD.UploadNo AS GIT_UPLOADNO, '' AS GIT_PNRDATE, " _
            & "                      dbo.AmdHeadAMD.RECLOC AS GIT_PNR, dbo.AmdLineI.PSGRLNME AS GIT_LASTNAME, dbo.AmdLineI.PSGRFNME AS GIT_FIRSTNAME, " _
            & "                      dbo.AmdLineC.BKAGT AS GIT_BOOKINGAGENT, dbo.AmdLineTMC.TKTNBR AS GIT_TICKETNUMBER, dbo.AmdHeadAMD.FName AS GIT_FILENAME, " _
            & "                      dbo.AmdHeadAMD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.AmdHeadAMD.IDBKG AS GIT_PCC " _
            & "FROM         dbo.AmdHeadAMD INNER JOIN " _
            & "                      dbo.AmdLineI ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineI.UpLoadNo INNER JOIN " _
            & "                      dbo.AmdLineTMC ON dbo.AmdLineI.UpLoadNo = dbo.AmdLineTMC.UpLoadNo LEFT OUTER JOIN " _
            & "                      dbo.AmdLineC ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineC.UpLoadNo " _
            & "Where (dbo.AmdHeadAMD.UploadNo = " & UploadNo & ")"
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'vSQL_String = "" _
        '    & "INSERT INTO dbo.GDSINTRAYTABLE " _
        '    & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC]) " _
        '    & "SELECT     'AIR' AS GIT_ID, 'AMADEUS' AS GIT_GDS, 'AMADEUS' AS GIT_INTERFACE, dbo.AmdHeadAMD.UploadNo AS GIT_UPLOADNO, '' AS GIT_PNRDATE, " _
        '    & "                      dbo.AmdHeadAMD.RECLOC AS GIT_PNR, dbo.AmdLineI.PSGRLNME AS GIT_LASTNAME, dbo.AmdLineI.PSGRFNME AS GIT_FIRSTNAME, " _
        '    & "                      dbo.AmdLineC.BKAGT AS GIT_BOOKINGAGENT, dbo.AmdLineTMC.TKTNBR AS GIT_TICKETNUMBER, dbo.AmdHeadAMD.FName AS GIT_FILENAME, " _
        '    & "                      dbo.AmdHeadAMD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.AmdHeadAMD.IDBKG AS GIT_PCC " _
        '    & "FROM         dbo.AmdHeadAMD INNER JOIN " _
        '    & "                      dbo.AmdLineI ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineI.UpLoadNo INNER JOIN " _
        '    & "                      dbo.AmdLineTMC ON dbo.AmdLineI.UpLoadNo = dbo.AmdLineTMC.UpLoadNo AND dbo.AmdLineI.PSGRNBR = dbo.AmdLineTMC.TSMNBR LEFT OUTER JOIN " _
        '    & "                      dbo.AmdLineC ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineC.UpLoadNo " _
        '    & "Where (dbo.AmdHeadAMD.UploadNo = " & UploadNo & ")"
        'by Abhi on 16-Aug-2016 for caseid 6673 changes in Amadeus EMD loading-Passenger number
        'vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC],[GIT_FILEISSUEDATE]) " _
            & "SELECT     'AIR' AS GIT_ID, 'AMADEUS' AS GIT_GDS, 'AMADEUS' AS GIT_INTERFACE, dbo.AmdHeadAMD.UploadNo AS GIT_UPLOADNO, '' AS GIT_PNRDATE, " _
            & "                      dbo.AmdHeadAMD.RECLOC AS GIT_PNR, dbo.AmdLineI.PSGRLNME AS GIT_LASTNAME, dbo.AmdLineI.PSGRFNME AS GIT_FIRSTNAME, " _
            & "                      dbo.AmdLineC.BKAGT AS GIT_BOOKINGAGENT, dbo.AmdLineTMC.TKTNBR AS GIT_TICKETNUMBER, dbo.AmdHeadAMD.FName AS GIT_FILENAME, " _
            & "                      dbo.AmdHeadAMD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.AmdHeadAMD.IDBKG AS GIT_PCC, '" & vDateLastModified_String & "' AS GIT_FILEISSUEDATE " _
            & "FROM         dbo.AmdHeadAMD INNER JOIN " _
            & "                      dbo.AmdLineI ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineI.UpLoadNo INNER JOIN " _
            & "                      dbo.AmdLineTMC ON dbo.AmdLineI.UpLoadNo = dbo.AmdLineTMC.UpLoadNo AND dbo.AmdLineI.PSGRNBR = dbo.AmdLineTMC.TSMNBR LEFT OUTER JOIN " _
            & "                      dbo.AmdLineC ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineC.UpLoadNo " _
            & "Where (dbo.AmdHeadAMD.UploadNo = " & UploadNo & ")"
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        ErrDetails_String = " GDSINTRAYTABLE 2."
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
        vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC],[GIT_FILEISSUEDATE],[GIT_CUSTOMERUSERCODE],[GIT_PENWAIT],[GIT_FILECREATEDDATE],[GIT_ISSUEDBY], [GIT_PENLINEBRID]) " _
            & "SELECT     'AIR' AS GIT_ID, 'AMADEUS' AS GIT_GDS, 'AMADEUS' AS GIT_INTERFACE, dbo.AmdHeadAMD.UploadNo AS GIT_UPLOADNO, '' AS GIT_PNRDATE, " _
            & "                      dbo.AmdHeadAMD.RECLOC AS GIT_PNR, dbo.AmdLineI.PSGRLNME AS GIT_LASTNAME, dbo.AmdLineI.PSGRFNME AS GIT_FIRSTNAME, " _
            & "                      dbo.AmdLineC.BKAGT AS GIT_BOOKINGAGENT, dbo.AmdLineTMC.TKTNBR AS GIT_TICKETNUMBER, dbo.AmdHeadAMD.FName AS GIT_FILENAME, " _
            & "                      dbo.AmdHeadAMD.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.AmdHeadAMD.IDBKG AS GIT_PCC, '" & vDateLastModified_String & "' AS GIT_FILEISSUEDATE, " _
            & "                      '" & SkipChars(GIT_CUSTOMERUSERCODE_String) & "' AS GIT_CUSTOMERUSERCODE, '" & SkipChars(GIT_PENWAIT_String) & "' AS GIT_PENWAIT, " _
            & "                      '" & vFileDateCreated_String & "' AS GIT_FILECREATEDDATE, dbo.AmdLineC.TKTAGT AS GIT_ISSUEDBY, '" & GIT_PENLINEBRID_String & "' AS GIT_PENLINEBRID " _
            & "FROM         dbo.AmdHeadAMD INNER JOIN " _
            & "                      dbo.AmdLineI ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineI.UpLoadNo INNER JOIN " _
            & "                      dbo.AmdLineTMC ON dbo.AmdLineI.UpLoadNo = dbo.AmdLineTMC.UpLoadNo AND dbo.AmdLineI.PSGRNBR = dbo.AmdLineTMC.PSGRNBR LEFT OUTER JOIN " _
            & "                      dbo.AmdLineC ON dbo.AmdHeadAMD.UploadNo = dbo.AmdLineC.UpLoadNo " _
            & "Where (dbo.AmdHeadAMD.UploadNo = " & UploadNo & ")"
        'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        'by Abhi on 16-Aug-2016 for caseid 6673 changes in Amadeus EMD loading-Passenger number
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'by Abhi on 18-Dec-2015 for caseid 5894 Amadeus EMD ticket is duplicated in GDS intray
        'by Abhi on 17-Mar-2014 for caseid 3812 Amadeus duplicated GDS In Tray
        'by Abhi on 17-Apr-2014 for caseid 3934 GDS Amadeus:If the tag AmdHeadAMD.ACTTAG contains the value "VOID" can we skip it to show from intray
        'dbCompany.Execute vSQL_String
        If vValidated_Boolean = True Then
            dbCompany.Execute vSQL_String
        End If
        'by Abhi on 17-Apr-2014 for caseid 3934 GDS Amadeus:If the tag AmdHeadAMD.ACTTAG contains the value "VOID" can we skip it to show from intray
        'by Abhi on 12-Mar-2014 for caseid 3791 EMD number picking from amadeus file
        'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
        'FMain.stbUpload.Panels(1).Text = "File moving to Destination..."
        'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    End If
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    If vOutOfBookings_Boolean = True Then
        FMain.stbUpload.Panels(1).Text = "File moving to Outofbookings..."
        FileObj.CopyFile txtFileName.Text, txtDirName.Text & "\Outofbookings\", True
    Else
        FMain.stbUpload.Panels(1).Text = "File moving to Destination..."
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
        FileObj.CopyFile txtFileName.Text, txtRelLocation.Text & "\", True
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    End If
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
'On Error Resume Next
    FileObj.DeleteFile txtFileName.Text, True
    'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
    INIWrite App.Path & "\PenSEARCH.ini", "Amadeus", "LUFPNR", LUFPNR_String
    INIWrite App.Path & "\PenSEARCH.ini", "Amadeus", "LUFDate", DateFormat(Date)
    INIWrite App.Path & "\PenSEARCH.ini", "Amadeus", "LUFTime", TimeFormat12HRS(time) & "(" & TimeFormat(time) & ")"
    'by Abhi on 18-May-2011 for caseid 1757 Added Events for GDSAuto
    'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
    'Call EventLog("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & LUFPNR_String & " - " & FMain.stbUpload.Panels(3).Text)
    Call EventLog("#PenGDS# " & PadR(FMain.stbUpload.Panels(2).Text, 9) & " - " & PadR(LUFPNR_String, 6) & " - " & PadR(FMain.stbUpload.Panels(3).Text, 27))
    'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
    'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    'by Abhi on 17-Apr-2014 for caseid 3934 GDS Amadeus:If the tag AmdHeadAMD.ACTTAG contains the value "VOID" can we skip it to show from intray
    'dbCompany.CommitTrans
    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
    If PENErr_BeginTrans = True Then
    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        If vValidated_Boolean = True Then
            dbCompany.CommitTrans
        Else
            dbCompany.RollbackTrans
        End If
    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        PENErr_BeginTrans = False
    End If
    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
    'by Abhi on 17-Apr-2014 for caseid 3934 GDS Amadeus:If the tag AmdHeadAMD.ACTTAG contains the value "VOID" can we skip it to show from intray
    'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
Exit Function
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
End Function

Private Sub ClearTable(TableName As String)
dbCompany.Execute "Delete  from " & TableName
End Sub

Public Function ClearAll()
ClearTable "AmdHeadAMD"
ClearTable "AmdHeadBLK"
ClearTable "AmdLineA"
ClearTable "AmdLineB"
ClearTable "AmdLineC"
ClearTable "AmdLineD"
ClearTable "AmdLineG"
ClearTable "AmdLineH"
ClearTable "AmdLineI"
ClearTable "AmdLineK"
ClearTable "AmdLineL"
ClearTable "AmdLineQ"
ClearTable "AmdLineT"
ClearTable "AmdLineU"
ClearTable "AmdLineX"
ClearTable "AmdLineY"
ClearTable "AmdLineM"
ClearTable "AmdLineN"
ClearTable "AmdLineO"
ClearTable "AmdLineFP"
ClearTable "AmdLineFM"
ClearTable "AmdLineTK"
ClearTable "AmdLineFV"
ClearTable "AmdLineRM"

ClearTable "AmdLineTAX"
ClearTable "AmdLineSSR"
ClearTable "AmdLineKFT"
ClearTable "AmdLineTSA"
ClearTable "AmdLineOSI"
End Function

Private Function InserAirMain(DataLine As String, UploadNo As Long)
Dim aa, Coll As Collection
aa = SplitTwo(DataLine, 7)
Set Coll = BLK(aa(1))
InsertData Coll, "AmdHeadBLK", "BLK", UploadNo
End Function

Private Function InsertAMD(Data, UploadNo As Long, Optional FileName As String = "")
Dim Coll As Collection
'by Abhi on 09-Nov-2009 VLOCATOR in AirSegDetails
Dim vi As Long
Dim CollAMD As New Collection
Dim CollVLOCATOR As New Collection

Set Coll = AMD(Data)
'by Abhi on 09-Nov-2009 VLOCATOR in AirSegDetails
Set CollVLOCATOR = Coll("VLOCATOR")
For vi = 1 To CollVLOCATOR.Count
    InsertDataByFieldName CollVLOCATOR(vi), "AmdHeadAMD_VLOCATOR", "AMD", UploadNo
    DoEvents
Next

Set CollAMD = Coll("AMD")
CollAMD.Add FileName, "FName"
CollAMD.Add 0, "GDSAutoFailed"
InsertData CollAMD, "AmdHeadAMD", "AMD", UploadNo
End Function

Private Function TSA(Data As String) As Collection
Dim Coll As New Collection
Dim Splited, MainSplit
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
MainSplit = SplitForce(Data, ";S", 3)
Splited = SplitForce(MainSplit(0), ";", 3)

'temp = SplitForce(splited(0), "+", 3) 'TSA prob for ucanfly
temp = SplitForce(Splited(0), "+", 4)
    temp2 = SplitWithLengths(temp(0), 1, 3)
        Coll.Add temp2(0), "SALEINFO"
        Coll.Add temp2(1), "INTIND"
    Coll.Add temp(1), "SALEIND"
    Coll.Add temp(2), "ORGDEST"
    Coll.Add temp(3), "DOMISO"
    
    
temp = SplitForce(Splited(1), "+", 3)
    temp2 = SplitWithLengths(temp(0), 1, 1)
        Coll.Add temp2(0), "PRCINFO"
        Coll.Add temp2(1), "PRCDE"
    Coll.Add temp(1), "FCMI"
    Coll.Add temp(2), "JRNYTYPE"
    Coll.Add temp(3), "JRNYCODE"
    Coll.Add temp(3), "PSURCHG"
'---------------
'main split 2
'---------------
temp = SplitFirstTwo(MainSplit(1), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
   temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"
Set TSA = Coll
End Function

Private Function TK(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Splited = SplitFirstTwo(Data, ";S")
    Coll.Add Splited(0), "TKTARR"
    temp = SplitWithLengths(Splited(0), 2, 5)
    Coll.Add temp(1), "TKTDATE"
Splited = SplitFirstTwo(Splited(1), ";P")
    temp = SplitFirstTwo(Splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"

    temp = SplitFirstTwo(Splited(1), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "PSGASS"
    Coll.Add temp2(1), "PSGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "PSGNBR2"
    Coll.Add TKT_NoIndex, "SLNO"
Set TK = Coll
End Function

Private Function FV(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Splited = SplitFirstTwo(Data, ";S")
    Coll.Add Splited(0), "TKTCARR"
Splited = SplitFirstTwo(Splited(1), ";P")
    temp = SplitFirstTwo(Splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"

    temp = SplitFirstTwo(Splited(1), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "PSGASS"
    Coll.Add temp2(1), "PSGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "PSGNBR2"
Set FV = Coll
End Function



Private Function RM(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Dim CollGEN As New Collection
Dim CollPenline As New Collection
Dim CollPenSplited As New Collection
Dim CollPenPass As New Collection
'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Amadeus
Dim CollPenPENFARE As New Collection


Dim TempPEN, TempPENSplited
Dim TempPENi As Long, TempPENCount As Long
Dim TempPENPNO As Long
Dim TempPENP1E1 As String
'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
Dim CollPenPENLINK As New Collection
'by Abhi on 10-Aug-2010 for caseid 1433 Penline PENAUTOOFF
Dim CollPenPENAUTOOFF As New Collection
'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
Dim CollPenPENO As New Collection
'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
Dim CollPenPENATOL As New Collection
'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
Dim CollPenPENAGENTCOM As New Collection
'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
Dim CollPenPENAIRTKT As New Collection
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
Dim CollPenPENAGROSS As New Collection
'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
Dim CollPenPENBILLCUR As New Collection
'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
Dim CollPenPENWAIT As New Collection
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
Dim CollPenPENVC As New Collection
'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
Dim CollPenPENCS As New Collection
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Temp1 = Data
TempPEN = Data
TempPENP1E1 = Data

'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
'For TempPENPNO = 1 To 99
'    'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
'    'If InStr(1, TempPEN, "P" & TempPENPNO & "E-", vbTextCompare) > 0 Or InStr(1, TempPEN, "P" & TempPENPNO & "T-", vbTextCompare) > 0 Then
'    If InStr(1, TempPEN, PENLINEID_String & "E" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, TempPEN, PENLINEID_String & "T" & TempPENPNO & "-", vbTextCompare) > 0 Then
'        Set CollPenSplited = RMPEN_PENLINEPassenger(TempPEN, TempPENPNO)
'        CollPenPass.Add CollPenSplited, str(TempPENPNO)
'    End If
'    DoEvents
'Next

TempPENSplited = Split(TempPEN, "\")
TempPENCount = UBound(TempPENSplited)
For TempPENi = 0 To TempPENCount
    'If UCase(Left(TempPENSplited(TempPENi), 4)) = UCase("PEN/") Then
    'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
    'If InStr(1, UCase(TempPENSplited(TempPENi)), UCase("PEN/"), vbTextCompare) > 0 Then
    If InStr(1, UCase(TempPENSplited(TempPENi)), UCase(PENLINEID_String & "PEN/"), vbTextCompare) > 0 Then
        'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
        TempPENP1E1 = TempPENSplited(TempPENi)
        For TempPENPNO = 1 To 99
            'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
            'If InStr(1, TempPENP1E1, "E" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, TempPENP1E1, "T" & TempPENPNO & "-", vbTextCompare) > 0 Then
            If InStr(1, TempPENP1E1, "E" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, TempPENP1E1, "T" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, TempPENP1E1, "DOB" & TempPENPNO & "-", vbTextCompare) > 0 Then
            'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
                Set CollPenSplited = RMPEN_PENLINEPassenger(TempPENP1E1, TempPENPNO)
                CollPenPass.Add CollPenSplited, str(TempPENPNO)
            End If
            DoEvents
        Next

        'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
        If Trim(UCase(PENLINEID_String & "PEN/")) <> Trim(UCase(TempPENP1E1)) Then
            'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
            'Set CollPenSplited = RMPEN_PENLINE(TempPENSplited(TempPENi))
            Set CollPenSplited = RMPEN_PENLINE(TempPENP1E1)
            CollPenline.Add CollPenSplited, str(TempPENi)
        End If
    End If
    'by Abhi on 30-Apr-2014 for caseid 3974 Customer COde Picking Amadeus file
    If InStr(1, TempPEN, "SUBAGENCY_CODE:", vbTextCompare) > 0 Then
        Set CollPenSplited = RMSUBAGENCY(TempPENP1E1)
        CollPenline.Add CollPenSplited, str(TempPENi)
    End If
    'by Abhi on 30-Apr-2014 for caseid 3974 Customer COde Picking Amadeus file
    DoEvents
Next

'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
'If InStr(1, UCase(TempPEN), UCase("PENFARE/"), vbTextCompare) > 0 Then
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENFARE/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENFARE(TempPEN, PENFAREPNO_Long)
    CollPenPENFARE.Add CollPenSplited, str(CollPenSplited("PNO"))
    PENFAREPNO_Long = PENFAREPNO_Long + 1
End If

'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENLINK/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENLINK(TempPEN)
    If CollPenSplited.Count > 0 Then
        CollPenPENLINK.Add CollPenSplited, "PENLINK"
    End If
End If

'by Abhi on 10-Aug-2010 for caseid 1433 Penline PENAUTOOFF
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENAUTOOFF"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENAUTOOFF(TempPEN)
    If CollPenSplited.Count > 0 Then
        CollPenPENAUTOOFF.Add CollPenSplited, "PENAUTOOFF"
    End If
End If

'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENO/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENO(TempPEN)
    CollPenPENO.Add CollPenSplited, "PENO"
End If

'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENATOL/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENATOL(TempPEN)
    If CollPenSplited.Count > 0 Then
        CollPenPENATOL.Add CollPenSplited, "PENATOL"
    End If
End If

'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENAGENTCOM/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENAGENTCOM(TempPEN)
    CollPenPENAGENTCOM.Add CollPenSplited, "PENAGENTCOM"
End If
'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)

'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENAIRTKT/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENAIRTKT(TempPEN)
    CollPenPENAIRTKT.Add CollPenSplited, "PENAIRTKT"
End If
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENAGROSS/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENAGROSS(TempPEN)
    CollPenPENAGROSS.Add CollPenSplited, "PENAGROSS"
End If
'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus

'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENBILLCUR"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENBILLCUR(TempPEN)
    CollPenPENBILLCUR.Add CollPenSplited, "PENBILLCUR"
End If
'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENWAIT"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENWAIT(TempPEN)
    If CollPenSplited.Count > 0 Then
        CollPenPENWAIT.Add CollPenSplited, "PENWAIT"
    End If
End If
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENVC"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENVC(TempPEN)
    If CollPenSplited.Count > 0 Then
        CollPenPENVC.Add CollPenSplited, "PENVC"
    End If
End If
'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENCS"), vbTextCompare) > 0 Then
    Set CollPenSplited = RMPEN_PENLINEPENCS(TempPEN)
    If CollPenSplited.Count > 0 Then
        CollPenPENCS.Add CollPenSplited, "PENCS"
    End If
End If
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Splited = SplitFirstTwo(Data, ";S")
    temp = SplitTwo(Splited(0), 1)
    CollGEN.Add temp(0), "RMTYPE"
    CollGEN.Add temp(0), "PNRRMK"
Splited = SplitFirstTwo(Splited(1), ";P")
    temp = SplitFirstTwo(Splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    CollGEN.Add temp2(0), "SEGASS"
    CollGEN.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    CollGEN.Add temp2(0), "SEGNBR2"

    temp = SplitFirstTwo(Splited(1), "-")
    temp2 = SplitTwoReverse(temp(0), 2)
    CollGEN.Add temp2(0), "PSGASS"
    CollGEN.Add temp2(1), "PSGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    CollGEN.Add temp2(0), "PSGNBR2"

    Coll.Add CollGEN, "GEN"
    Coll.Add CollPenline, "PEN"
    Coll.Add CollPenPass, "PENPASSENGER"
    'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
    Coll.Add CollPenPENFARE, "PENFARE"
    'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
    Coll.Add CollPenPENLINK, "PENLINK"
    'by Abhi on 10-Aug-2010 for caseid 1433 Penline PENAUTOOFF
    Coll.Add CollPenPENAUTOOFF, "PENAUTOOFF"
    'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
    Coll.Add CollPenPENO, "PENO"
    'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
    Coll.Add CollPenPENATOL, "PENATOL"
    'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
    Coll.Add CollPenPENAGENTCOM, "PENAGENTCOM"
    'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add CollPenPENAIRTKT, "PENAIRTKT"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    Coll.Add CollPenPENAGROSS, "PENAGROSS"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add CollPenPENBILLCUR, "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
    Coll.Add CollPenPENWAIT, "PENWAIT"
    'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
    'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
    Coll.Add CollPenPENVC, "PENVC"
    'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add CollPenPENCS, "PENCS"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    
Set RM = Coll
End Function

'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation
Private Function RC(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Dim CollGEN As New Collection
Dim CollPenline As New Collection
Dim CollPenSplited As New Collection
Dim CollPenPass As New Collection
'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Amadeus
Dim CollPenPENFARE As New Collection
'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation
Dim CollMARKUP As New Collection
'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation

Dim TempPEN, TempPENSplited
Dim TempPENi As Long, TempPENCount As Long
Dim TempPENPNO As Long
Dim TempPENP1E1 As String
Dim CollPenPENLINK As New Collection
Dim CollPenPENAUTOOFF As New Collection
Dim CollPenPENO As New Collection
Dim CollPenPENATOL As New Collection
Dim CollPenPENAGENTCOM As New Collection
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
Dim CollPenPENWAIT As New Collection
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT

Temp1 = Data
TempPEN = Data
TempPENP1E1 = Data

'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation
'RC LONDZ2100-W/FLIGHT_ADT_MARKUP:GBP 26.00
'RC LONDZ2100-W/FLIGHT_CH_MARKUP:GBP 19.50
'RC LONDZ2100-W/FLIGHT_INF_MARKUP:GBP 2.60
'RM PENFARE/445.00/78.06/543.50/ADT
'PENFARE/netfare/tax/sell/passenger type
If InStr(1, TempPEN, "_MARKUP:") > 0 Then
    'Debug.Print TempPEN
    If InStr(1, TempPEN, "_ADT_MARKUP:") > 0 Then
        Splited = SplitForce(TempPEN, ":GBP", 2)
        TempPEN = " " & PENLINEID_String & "MARKUP/0.00/0.00/" & Splited(1) & "/ADT"
    End If
    
    'by Abhi on 25-May-2017 for caseid 7484 Passenger type reading from Amadeus internal mark up line
    'TempPEN = Replace(TempPEN, "_CH_MARKUP:", "_CHD_MARKUP:", , , vbTextCompare)
    If (InStr(1, TempPEN, "_C") > 0 Or InStr(1, TempPEN, "_B") > 0) And InStr(1, TempPEN, "_MARKUP:") > 0 Then
        TempPEN = RC_CHD_YTH_MARKUP(TempPEN)
    End If
    'by Abhi on 25-May-2017 for caseid 7484 Passenger type reading from Amadeus internal mark up line
    
    'If InStr(1, TempPEN, "_CH_MARKUP:") > 0 Then
    If InStr(1, TempPEN, "_CHD_MARKUP:") > 0 Then
        Splited = SplitForce(TempPEN, ":GBP", 2)
        TempPEN = " " & PENLINEID_String & "MARKUP/0.00/0.00/" & Splited(1) & "/CHD"
    End If
    
    'by Abhi on 25-May-2017 for caseid 7484 Passenger type reading from Amadeus internal mark up line
    If InStr(1, TempPEN, "_YTH_MARKUP:") > 0 Then
        Splited = SplitForce(TempPEN, ":GBP", 2)
        TempPEN = " " & PENLINEID_String & "MARKUP/0.00/0.00/" & Splited(1) & "/YTH"
    End If
    'by Abhi on 25-May-2017 for caseid 7484 Passenger type reading from Amadeus internal mark up line
    
    If InStr(1, TempPEN, "_INF_MARKUP:") > 0 Then
        Splited = SplitForce(TempPEN, ":GBP", 2)
        TempPEN = " " & PENLINEID_String & "MARKUP/0.00/0.00/" & Splited(1) & "/INF"
    End If
End If
'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation

'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
'If InStr(1, UCase(TempPEN), UCase("PENFARE/"), vbTextCompare) > 0 Then
If InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "MARKUP/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RC_MARKUP(TempPEN, PENFAREPNO_Long)
    CollMARKUP.Add CollPenSplited, str(CollPenSplited("PNO"))
    PENFAREPNO_Long = PENFAREPNO_Long + 1
End If

Splited = SplitFirstTwo(Data, ";S")
    temp = SplitTwo(Splited(0), 1)
    CollGEN.Add temp(0), "RMTYPE"
    CollGEN.Add temp(0), "PNRRMK"
Splited = SplitFirstTwo(Splited(1), ";P")
    temp = SplitFirstTwo(Splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    CollGEN.Add temp2(0), "SEGASS"
    CollGEN.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    CollGEN.Add temp2(0), "SEGNBR2"

    temp = SplitFirstTwo(Splited(1), "-")
    temp2 = SplitTwoReverse(temp(0), 2)
    CollGEN.Add temp2(0), "PSGASS"
    CollGEN.Add temp2(1), "PSGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    CollGEN.Add temp2(0), "PSGNBR2"

    Coll.Add CollGEN, "GEN"
    Coll.Add CollPenline, "PEN"
    Coll.Add CollPenPass, "PENPASSENGER"
    Coll.Add CollPenPENFARE, "PENFARE"
    Coll.Add CollPenPENLINK, "PENLINK"
    Coll.Add CollPenPENAUTOOFF, "PENAUTOOFF"
    Coll.Add CollPenPENO, "PENO"
    Coll.Add CollPenPENATOL, "PENATOL"
    Coll.Add CollPenPENAGENTCOM, "PENAGENTCOM"
    'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation
    Coll.Add CollMARKUP, "MARKUP"
    'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation
    'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
    Coll.Add CollPenPENWAIT, "PENWAIT"
    'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
Set RC = Coll
End Function

'by Abhi on 25-May-2017 for caseid 7484 Passenger type reading from Amadeus internal mark up line
Private Function RC_CHD_YTH_MARKUP(ByVal Data As String) As String
Dim vi_Integer As Integer
    Data = Replace(Data, "_CH_MARKUP:", "_CHD_MARKUP:", , , vbTextCompare)
    For vi_Integer = 0 To 17
        If vi_Integer >= 12 And vi_Integer <= 16 Then '"C12-C16" YTH
            Data = Replace(Data, "_C" & Format(vi_Integer, "00") & "_MARKUP:", "_YTH_MARKUP:", , , vbTextCompare)
            Data = Replace(Data, "_B" & Format(vi_Integer, "00") & "_MARKUP:", "_YTH_MARKUP:", , , vbTextCompare)
        Else
            Data = Replace(Data, "_C" & Format(vi_Integer, "00") & "_MARKUP:", "_CHD_MARKUP:", , , vbTextCompare)
            Data = Replace(Data, "_B" & Format(vi_Integer, "00") & "_MARKUP:", "_CHD_MARKUP:", , , vbTextCompare)
        End If
        DoEvents
    Next
    
RC_CHD_YTH_MARKUP = Data
End Function
'by Abhi on 25-May-2017 for caseid 7484 Passenger type reading from Amadeus internal mark up line

Private Function RMPEN_PENLINEPassenger(Data1, ByVal PNO As Long) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
        
    'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
    'Temp1 = ExtractBetween(Data, "P" & PNO & "E-", "\")
    'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
    'Temp1 = ExtractBetween(Data, PENLINEID_String & "E" & PNO & "-", "\")
    Temp1 = ExtractBetween(Data, "E" & PNO & "-", "\")
        'by Abhi on 13-Apr-2010 for caseid 1303 remove slash (/) from penline email
        'by Abhi on 08-Mar-2018 for caseid 8331 Amadeus,Worldspan and Sabre - Pengds penline email replacement
        'Temp1 = Replace(Temp1, "/", "")
        Temp1 = PenlineEmailReplacement(Temp1)
        Temp1 = Replace(Temp1, "/", "")
        'by Abhi on 08-Mar-2018 for caseid 8331 Amadeus,Worldspan and Sabre - Pengds penline email replacement
        'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
        'Coll.Add Temp1, "PEMAIL"
        'by Abhi on 11-Oct-2017 for caseid 7862 Passenger Email id validation in folder should also save multiple email id with coma separator
        'Coll.Add Left(Temp1, 50), "PEMAIL"
        Coll.Add Left(Temp1, 300), "PEMAIL"
        'by Abhi on 11-Oct-2017 for caseid 7862 Passenger Email id validation in folder should also save multiple email id with coma separator
    'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
    'Temp1 = ExtractBetween(Data, "P" & PNO & "T-", "\")
    'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
    'Temp1 = ExtractBetween(Data, PENLINEID_String & "T" & PNO & "-", "\")
    Temp1 = ExtractBetween(Data, "T" & PNO & "-", "\")
        'by Abhi on 13-Apr-2010 for caseid 1303 remove slash (/) from penline email
        Temp1 = Replace(Temp1, "/", "")
        'by Abhi on 18-May-2011 for caseid 1756 Error file amadeus error in penline TELE
        'Coll.Add Temp1, "PTELE"
        Coll.Add Left(Temp1, 50), "PTELE"
        Coll.Add PNO, "PNO"
        Coll.Add "", "SELLINF"
        'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Amadeus
        Coll.Add "0", "FARE"
        Coll.Add "0", "TAXES"
        Coll.Add "0", "MARKUP"
        'by Abhi on 06-Apr-2010 for caseid 1293 PENFARE modification for Worldspan and Amadeus
        Coll.Add "0", "PENFARESELL"
        Coll.Add "", "PENFAREPASSTYPE"
        'by Abhi on 13-Apr-2010 for caseid 1301 Amadeus Delivery Address from penline
        Coll.Add "", "DeliAdd"
        'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
        Coll.Add "", "MC"
        Coll.Add "", "BB"
        'by Abhi on 18-Jun-2010 for caseid 1399 Supplier of ticket for PENFARE
        Coll.Add "", "PENFARESUPPID"
        'by Abhi on 27-Jul-2010 for caseid 1439 Penline PENFARE Airline
        Coll.Add "", "PENFAREAIRID"
        'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
        Coll.Add "", "PENLINKPNR"
        'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
        Coll.Add "", "PENOPRDID"
        Coll.Add 0, "PENOQTY"
        Coll.Add 0, "PENORATE"
        Coll.Add 0, "PENOSELL"
        Coll.Add "", "PENOSUPPID"
        'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
        Coll.Add "", "PENATOLTYPE"
        'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
        Coll.Add "", "INETREF"
        'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
        Coll.Add "", "PENOPAYMETHODID"
        'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
        Coll.Add 0, "PENAGENTCOMSUM"
        Coll.Add 0, "PENAGENTCOMVAT"
        'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
        Coll.Add "", "TKTDEADLINE"
        'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
        'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
        Coll.Add "", "AIRTKTPAX"
        Coll.Add "", "AIRTKTTKT"
        Coll.Add "", "AIRTKTDATE"
        'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
        'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
        Coll.Add 0, "PENAGROSSADULT"
        Coll.Add 0, "PENAGROSSCHILD"
        Coll.Add 0, "PENAGROSSINFANT"
        Coll.Add 0, "PENAGROSSPACKAGE"
        'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
        'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
        Coll.Add 0, "PENAGROSSYOUTH"
        'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Temp1 = ExtractBetween(Data, "DOB" & PNO & "-", "\")
        Temp1 = Replace(Temp1, "/", "")
        Temp1 = Trim(Temp1)
        Coll.Add Left(Temp1, 9), "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Data1 = Data
Set RMPEN_PENLINEPassenger = Coll
End Function

Private Function RMPEN_PENLINE(ByVal Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
Dim vi As Long
    Splited = SplitForce(Data, "/", 9)
    
    Temp1 = ExtractBetween(Data, "AC-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "AC"
    Coll.Add Left(Temp1, 50), "AC"
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    If Trim(Left(Temp1, 50)) <> "" Then
        GIT_CUSTOMERUSERCODE_String = Left(Temp1, 50)
    End If
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    
    Temp1 = ExtractBetween(Data, "DES-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "DES"
    Coll.Add Left(Temp1, 50), "DES"
    
    Temp1 = ExtractBetween(Data, "REFE-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "REFE"
    Coll.Add Left(Temp1, 50), "REFE"
    
    Temp1 = ExtractBetween(Data, "SELL-ADT-", "/")
    Coll.Add Temp1, "SELLADT"
    
    Temp1 = ExtractBetween(Data, "SELL-CHD-", "/")
    Coll.Add Temp1, "SELLCHD"
    
    Temp1 = ExtractBetween(Data, "SELL-INF-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "SELLINF"
    Coll.Add Left(Temp1, 50), "SELLINF"
    
    Temp1 = ExtractBetween(Data, "SCHARGE-", "/")
    Coll.Add Temp1, "SCHARGE"
    
    Temp1 = ExtractBetween(Data, "DEPT-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "DEPT"
    Coll.Add Left(Temp1, 50), "DEPT"
    
    Temp1 = ExtractBetween(Data, "BRANCH-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "BRANCH"
    Coll.Add Left(Temp1, 50), "BRANCH"
    
    Temp1 = ExtractBetween(Data, "CC-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "CUSTCC"
    Coll.Add Left(Temp1, 200), "CUSTCC"
            
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Amadeus
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add "0", "MARKUP"
    'by Abhi on 06-Apr-2010 for caseid 1293 PENFARE modification for Worldspan and Amadeus
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    'by Abhi on 13-Apr-2010 for caseid 1301 Amadeus Delivery Address from penline
    Temp1 = ExtractBetween(Data, "ADD-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "DeliAdd"
    Coll.Add Left(Temp1, 500), "DeliAdd"
    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
    Temp1 = ExtractBetween(Data, "MC-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "MC"
    Coll.Add Left(Temp1, 50), "MC"
    Temp1 = ExtractBetween(Data, "BB-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "BB"
    Coll.Add Left(Temp1, 50), "BB"
    'by Abhi on 18-Jun-2010 for caseid 1399 Supplier of ticket for PENFARE
    Coll.Add "", "PENFARESUPPID"
    'by Abhi on 27-Jul-2010 for caseid 1439 Penline PENFARE Airline
    Coll.Add "", "PENFAREAIRID"
    'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
    Coll.Add "", "PENLINKPNR"
    'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
    Coll.Add "", "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Temp1 = ExtractBetween(Data, "INETREF-", "/")
    Coll.Add Left(Temp1, 100), "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add "", "PENOPAYMETHODID"
    'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Temp1 = ExtractBetween(Data, "TKTDEADLINE-", "/")
    Coll.Add Left(Temp1, 9), "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Temp1 = ExtractBetween(Data, "DEPOSITAMT-", "/")
    Coll.Add Val(Temp1), "DEPOSITAMT"
    Temp1 = ExtractBetween(Data, "DEPOSITDUEDATE-", "/")
    Coll.Add Left(Temp1, 9), "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Data1 = Data
Set RMPEN_PENLINE = Coll
End Function

'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Amadeus
Private Function RMPEN_PENLINEPENFARE(Data1, ByVal pPNO_Long As Long) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add pPNO_Long, "PNO"
    Coll.Add "", "SELLINF"
    'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Amadeus
    'by Abhi on 18-Jun-2010 for caseid 1399 Supplier of ticket for PENFARE
    'splited = SplitForce(Data, "/", 5)
    'by Abhi on 27-Jul-2010 for caseid 1439 Penline PENFARE Airline
    'splited = SplitForce(Data, "/", 6)
    Splited = SplitForce(Data, "/", 7)
    Coll.Add Val(Splited(1)), "FARE"
    Coll.Add Val(Splited(2)), "TAXES"
    'by Abhi on 06-Apr-2010 for caseid 1293 PENFARE modification for Worldspan and Amadeus
    'Coll.Add Val(splited(3)), "MARKUP"
    Coll.Add "0", "MARKUP"
    'by Abhi on 06-Apr-2010 for caseid 1293 PENFARE modification for Worldspan and Amadeus
    Coll.Add Val(Splited(3)), "PENFARESELL"
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add splited(4), "PENFAREPASSTYPE"
    Coll.Add Left(Splited(4), 50), "PENFAREPASSTYPE"
    'by Abhi on 13-Apr-2010 for caseid 1301 Amadeus Delivery Address from penline
    Coll.Add "", "DeliAdd"
    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    'by Abhi on 18-Jun-2010 for caseid 1399 Supplier of ticket for PENFARE
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add splited(5), "PENFARESUPPID"
    Coll.Add Left(Splited(5), 50), "PENFARESUPPID"
    'by Abhi on 27-Jul-2010 for caseid 1439 Penline PENFARE Airline
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add splited(6), "PENFAREAIRID"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'Coll.Add Left(Splited(6), 50), "PENFAREAIRID"
    Coll.Add "", "PENFAREAIRID"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
    Coll.Add "", "PENLINKPNR"
    'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
    Coll.Add "", "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Coll.Add "", "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add "", "PENOPAYMETHODID"
    'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add Left(Splited(6), 50), "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Data1 = Data
Set RMPEN_PENLINEPENFARE = Coll
End Function

'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
Private Function RMPEN_PENLINEPENLINK(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
    Data = ExtractBetween(Data1, "PENLINK/", "\")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Data, "PENLINKPNR"
    Coll.Add Left(Data, 50), "PENLINKPNR"
    'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
    Coll.Add "", "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Coll.Add "", "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add "", "PENOPAYMETHODID"
    'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

'Data1 = Data
Set RMPEN_PENLINEPENLINK = Coll
End Function

'by Abhi on 10-Aug-2010 for caseid 1433 Penline PENAUTOOFF
Private Function RMPEN_PENLINEPENAUTOOFF(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    'by Abhi on 09-Oct-2010 for caseid 1511 Penline PENATOL
    Coll.Add "", "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Coll.Add "", "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add "", "PENOPAYMETHODID"
    'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

'Data1 = Data
Set RMPEN_PENLINEPENAUTOOFF = Coll
End Function

'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
Private Function RMPEN_PENLINEPENO(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    'splited = SplitForce(Data, "/", 6)
    Splited = SplitForce(Data, "/", 7)
    Coll.Add "", "PENLINKPNR"
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add splited(1), "PENOPRDID"
    Coll.Add Left(Splited(1), 50), "PENOPRDID"
    Coll.Add Val(Splited(2)), "PENOQTY"
    Coll.Add Val(Splited(3)), "PENORATE"
    Coll.Add Val(Splited(4)), "PENOSELL"
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add splited(5), "PENOSUPPID"
    Coll.Add Left(Splited(5), 50), "PENOSUPPID"
    'by Abhi on 09-Oct-2010 for caseid 1511 Penline PENATOL
    Coll.Add "", "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Coll.Add "", "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add Left(Splited(6), 50), "PENOPAYMETHODID"
    'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Data1 = Data
Set RMPEN_PENLINEPENO = Coll
End Function

'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
Private Function RMPEN_PENLINEPENATOL(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
    Data = ExtractBetween(Data1, "PENATOL/", "\")
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Data, "PENATOLTYPE"
    Coll.Add Left(Data, 50), "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Coll.Add "", "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add "", "PENOPAYMETHODID"
    'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

'Data1 = Data
Set RMPEN_PENLINEPENATOL = Coll
End Function

'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
Private Function RMPEN_PENLINEPENAGENTCOM(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
    Splited = SplitForce(Data, "/", 3)
    Coll.Add Val(Splited(1)), "PENAGENTCOMSUM"
    Coll.Add Val(Splited(2)), "PENAGENTCOMVAT"
    'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Data1 = Data
Set RMPEN_PENLINEPENAGENTCOM = Coll
End Function
'by Abhi on 16-Apr-2014 for caseid 2600 Penline PENAGENTCOM for all GDS(Amadeus)


Private Function OSI(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Splited = SplitFirstTwo(Data, ";S")
    temp = SplitWithLengthsPlus(Splited(0), 1, 3, 1)
    Coll.Add temp(1), "AIRCDE"
    Coll.Add temp(3), "FFTEXT"
Splited = SplitFirstTwo(Splited(1), ";P")
    temp = SplitFirstTwo(Splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"

    temp = SplitFirstTwo(Splited(1), "-")
    temp2 = SplitTwoReverse(temp(0), 2)
    Coll.Add temp2(0), "PSGASS"
    Coll.Add temp2(1), "PSGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "PSGNBR2"
Set OSI = Coll
End Function


Private Function N(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
    temp = SplitWithLengths(Data, 3, 28, 11)
    Coll.Add temp(0), "CURRTYP"
    Coll.Add temp(1), "FARESEG"
    Coll.Add temp(2), "BASE"
Set N = Coll
End Function
Private Function O(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
    temp = SplitWithLengths(Data, 28, 10)
    Coll.Add temp(0), "VALELEM"
    Coll.Add temp(1), "VALDTE"
Set O = Coll
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
rs.Fields("MSGTAG") = LineIdent
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
        ErrDetails_String = vbCrLf & "MSGTAG=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(temp).Name & ", FieldSize=" & rs.Fields(temp).DefinedSize & "<>" & Len(Data(temp)) & "," & vbCrLf & Data(temp) & vbCrLf
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
        ErrDetails_String = vbCrLf & "MSGTAG=" & LineIdent & ", Table=" & TableName & ", Field=" & temp & ", Field does not exist in Collection!," & vbCrLf
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    ElseIf Trim(ErrDetails_String) <> "" Then '-2147467259 Data provider or other service returned an E_FAIL status.
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    ElseIf Len(Data(temp)) > rs.Fields(temp).DefinedSize Then '-2147217887 Multiple-step operation generated errors. Check each status value.
        ErrDetails_String = vbCrLf & "MSGTAG=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(temp).Name & ", FieldSize=" & rs.Fields(temp).DefinedSize & "<>" & Len(Data(temp)) & "," & vbCrLf & Data(temp) & vbCrLf
    Else
        ErrDetails_String = vbCrLf & "MSGTAG=" & LineIdent & ", Table=" & TableName & ", Field=" & temp & "," & vbCrLf & Data(temp) & vbCrLf
    End If
    Err.Raise ErrNumber, , ErrDescription
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
End Function
Private Function InsertDataCollectionKey(Coll As Collection, TableNae As String, LineID, UploadNo As Long)
For j = 1 To Coll.Count
    InsertDataByFieldName Coll(j), TableNae, CStr(LineID), UploadNo
    DoEvents
Next
End Function


Private Function mSplitTypeAndSirName(mData)
    Dim Data
    Dim typ, retn(3)
    Data = Trim(mData)
    typ = Right(Data, 2)
    If typ = "MR" Then
        retn(0) = Left(Data, Len(Data) - 2)
        retn(1) = typ
        retn(2) = "Adult"
        GoTo returnPara
    ElseIf typ = "MS" Then
        retn(0) = Left(Data, Len(Data) - 2)
        retn(1) = typ
        retn(2) = "Adult"
        GoTo returnPara
    End If
    
    typ = Right(Data, 3)
    If typ = "MRS" Or typ = "MIS" Then
        retn(0) = Left(Data, Len(Data) - 3)
        retn(1) = typ
        retn(2) = "Adult"
        GoTo returnPara
    End If
    typ = Right(Data, 4)
    If typ = "MSTR" Then
        retn(0) = Left(Data, Len(Data) - 4)
        retn(1) = typ
        retn(2) = "Child"
        GoTo returnPara
    ElseIf typ = "MISS" Then
        retn(0) = Left(Data, Len(Data) - 4)
        retn(1) = typ
        retn(2) = "Adult"
        GoTo returnPara
    End If
    
    retn(0) = Data
    retn(1) = ""
    retn(2) = ""
    mSplitTypeAndSirName = retn
    Exit Function
returnPara:
    mSplitTypeAndSirName = retn
End Function

Public Function WhichUType(Line) As String
Dim temp
temp = SplitWithLengths(Line, 8, 3)
WhichUType = temp(1)
End Function


Private Function U_Hotel(Data As String) As Collection
Dim Coll As New Collection
Dim Splited, Index
Dim temp, temp2, temp3
Dim sTemp1, sTemp2, sTemp3
Index = 0
Splited = SplitForce(Data, ";", 22)

    temp = SplitTwo(Splited(0), 3)
    Coll.Add temp(0), "TTONBR"
    Coll.Add temp(1), "DATAMODI"
    
    temp2 = SplitForce(Splited(1), " ", 4)
    
    temp = SplitWithLengths(temp2(0), 3, 3)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "SEGID"
    
    temp = temp2(1)
    Coll.Add temp, "AIRCDE"
    
    temp = SplitWithLengths(temp2(2), 2, 2)
    Coll.Add temp(0), "STATUS"
    Coll.Add temp(1), "NBRRMS"
    
    temp = SplitWithLengths(Splited(2), 4, 1, 5)
       Coll.Add temp(0), "PNRSTAT"
       Coll.Add temp(2), "INDTE"
    
    temp = Splited(3)
    Coll.Add temp, "OUTDTE"

    temp = Splited(4)
    Coll.Add temp, "AIRPT"

    temp = Splited(5)
    Coll.Add temp, "CDETRAN"

    temp = Splited(6)
    Coll.Add temp, "SIPPRM"

    temp = Splited(7)
    Coll.Add temp, "RATEIND"

    temp = SplitWithLengths(Splited(8), 3, 3)
    Coll.Add temp(0), "CURRCDE"
    Coll.Add temp(1), "LCURCDE"
';
    temp2 = SplitForce(Splited(9), "+", 4)
    
    temp = SplitWithLengthsPlus(temp2(0), 10)
    Coll.Add "", "PRICEINFO"
    Coll.Add temp(0), "ROOMRTE"
    
    Coll.Add temp2(1), "DATERTE"
    Coll.Add temp2(2), "RATEPER"
    Coll.Add temp2(3), "RATELOC"
    Index = 9 ' change
    Coll.Add Splited(10 + Index), "RATETXT"

    Coll.Add Splited(11 + Index), "RATETYP"

    temp = SplitWithLengthsPlus(Splited(12 + Index), 4)
    Coll.Add temp(0), "RTTID"
    Coll.Add temp(1), "RTTOPT"

    
    Coll.Add Splited(13 + Index), "PROPID"

    temp = SplitWithLengthsPlus(Splited(14 + Index), 3)
    Coll.Add temp(0), "RTBKID"
    Coll.Add temp(1), "BOOKCDE"

    temp = SplitWithLengthsPlus(Splited(15 + Index), 3)
    Coll.Add temp(0), "BKGID"
    Coll.Add temp(1), "AIRIATA"

    temp = SplitWithLengthsPlus(Splited(16 + Index), 3)
    Coll.Add temp(0), "CONFID"
    Coll.Add temp(1), "CONFNBR"

    temp = SplitWithLengthsPlus(Splited(17 + Index), 3)
    Coll.Add temp(0), "APID"
    Coll.Add temp(1), "APOPT"

    temp = SplitWithLengthsPlus(Splited(18 + Index), 3)
    Coll.Add temp(0), "CDID"
    Coll.Add temp(1), "CDOPT"

    temp = SplitWithLengthsPlus(Splited(19 + Index), 3)
    Coll.Add temp(0), "CRID"
    Coll.Add temp(1), "CROPT"

    temp = SplitWithLengthsPlus(Splited(20 + Index), 3)
    Coll.Add temp(0), "DPID"
    Coll.Add temp(1), "DPOPT"

    temp = SplitWithLengthsPlus(Splited(21 + Index), 3)
    Coll.Add temp(0), "EXID"
    Coll.Add temp(1), "EXOPT"

    temp = SplitWithLengthsPlus(Splited(22 + Index), 3)
    Coll.Add temp(0), "FAID"
    Coll.Add temp(1), "FAOPT"

    temp = SplitWithLengthsPlus(Splited(23 + Index), 3)
    Coll.Add temp(0), "FMID"
    Coll.Add temp(1), "FMOPT"

    temp = SplitWithLengthsPlus(Splited(24 + Index), 2)
    Coll.Add temp(0), "GUID"
    Coll.Add temp(1), "GUOPT"

    temp = SplitWithLengthsPlus(Splited(25 + Index), 3)
    Coll.Add temp(0), "IDID"
    Coll.Add temp(1), "IDOPT"

    temp = SplitWithLengthsPlus(Splited(26 + Index), 3)
    Coll.Add temp(0), "MAID"
    Coll.Add temp(1), "MAOPT"

    temp = SplitWithLengthsPlus(Splited(27 + Index), 3)
    Coll.Add temp(0), "RAID"
    Coll.Add temp(1), "RAOPT"



    temp = SplitWithLengthsPlus(Splited(28 + Index), 3)
    Coll.Add temp(0), "RCID"
    Coll.Add temp(1), "RCOPT"

    temp = SplitWithLengthsPlus(Splited(29 + Index), 3)
    Coll.Add temp(0), "ROID"
    Coll.Add temp(1), "ROOPT"

    temp = SplitWithLengthsPlus(Splited(30 + Index), 3)
    Coll.Add temp(0), "SIID"
    Coll.Add temp(1), "SIOPT"

    temp = SplitWithLengthsPlus(Splited(31 + Index), 3)
    Coll.Add temp(0), "SFID"
    Coll.Add temp(1), "SFOPT"

    temp = SplitWithLengthsPlus(Splited(32 + Index), 3)
    Coll.Add temp(0), "ACID"
    Coll.Add temp(1), "ACTOPT"

Index = 8


    Coll.Add Splited(34 + Index), "HTLNME"
    Coll.Add Splited(35 + Index), "HFLA1"
    Coll.Add Splited(36 + Index), "HFLA2"
    Coll.Add Splited(37 + Index), "HFLCT"
    Coll.Add Splited(38 + Index), "HFLST"
    Coll.Add Splited(39 + Index), "HFLAR"
    Coll.Add Splited(40 + Index), "HFLCY"
    Coll.Add Splited(41 + Index), "HFLZP"
    Coll.Add Splited(42 + Index), "HFLPH"
    Coll.Add Splited(43 + Index), "HFLFX"
    Coll.Add Splited(44 + Index), "HFLTX9"
    Coll.Add Splited(45 + Index), "ACCESS"
    
    temp = SplitWithLengthsPlus(Splited(46 + Index), 4)
    Coll.Add temp(0), "COMID"
    Coll.Add temp(1), "AGNTCOM"
    
    Coll.Add Splited(47 + Index), "COMMAN"
    
    temp = SplitWithLengthsPlus(Splited(48 + Index), 4)
    Coll.Add temp(0), "DESID"
    Coll.Add temp(1), "RATE1"
    
    Coll.Add Splited(49 + Index), "RATE2"
    Coll.Add Splited(50 + Index), "RATE3"
    Coll.Add Splited(51 + Index), "RATE4"
    Coll.Add Splited(52 + Index), "RATE5"

    temp = SplitWithLengthsPlus(Splited(53 + Index), 4)
    Coll.Add temp(0), "TAXID"
    Coll.Add temp(1), "TAX"

    Coll.Add Splited(54 + Index), "TAXMAN"
    
    temp = SplitWithLengthsPlus(Splited(55 + Index), 4)
    Coll.Add temp(0), "SVCID"
    Coll.Add temp(1), "SERVICE"
    
    Coll.Add Splited(56 + Index), "SVCMAN"
    
    temp = SplitWithLengthsPlus(Splited(57 + Index), 4)
    Coll.Add temp(0), "MEAID"
    Coll.Add temp(1), "MEALS"
    
    Coll.Add Splited(58 + Index), "MEAMAN"
    
    temp = SplitWithLengthsPlus(Splited(59 + Index), 4)
    Coll.Add temp(0), "STAID"
    Coll.Add temp(1), "STAY"
    
    temp = SplitWithLengthsPlus(Splited(60 + Index), 4)
    Coll.Add temp(0), "HLDID"
    Coll.Add temp(1), "HOLD"
    
    temp = SplitWithLengthsPlus(Splited(61 + Index), 4)
    Coll.Add temp(0), "CNXID"
    Coll.Add temp(1), "CANCEL1"
    
    Coll.Add Splited(62 + Index), "CANCEL2"
    Coll.Add Splited(63 + Index), "CANCEL3"
    Coll.Add Splited(64 + Index), "CANCEL4"
    Coll.Add Splited(65 + Index), "CANCEL5"
    Coll.Add Splited(66 + Index), "ARRDEP"
    
    
    
    Set U_Hotel = Coll
    Exit Function
    ' Not checked
    'from here to rest
    
    
    
    
    
    temp = SplitWithLengthsPlus(Splited(67 + Index), 4)
    Coll.Add temp(0), "ADVID"
    Coll.Add temp(1), "ADVBK"
    
    temp = SplitWithLengthsPlus(Splited(68 + Index), 4)
    Coll.Add temp(0), "TTLID"
    Coll.Add temp(1), "TOTRT"
    
    Coll.Add Splited(69 + Index), "TTLMAN"
    
    temp = SplitWithLengthsPlus(Splited(70 + Index), 4)
    Coll.Add temp(0), "NGTID"
    Coll.Add temp(1), "NONIGHTS"
    
    Coll.Add Splited(71 + Index), "BOOKVIA"
    
    temp = SplitWithLengthsPlus(Splited(72 + Index), 3)
    Coll.Add temp(0), "VVID"
    Coll.Add temp(1), "VVALUE"
    
    temp2 = SplitForce(Splited(73 + Index), "+", 2)
    
    temp = SplitWithLengthsPlus(temp2(0), 1, 3)
    Coll.Add temp(0), "EQVAMNT"
    Coll.Add temp(1), "ACCCUR"
    Coll.Add temp(2), "ACAMNT"
    
    temp = SplitWithLengthsPlus(temp2(1), 3)
    Coll.Add temp(0), "ACCCUR"
    Coll.Add temp(1), "ACAMNT"

    temp = SplitWithLengthsPlus(Splited(76 + Index), 3)
    Coll.Add temp(0), "BILLID"
    Coll.Add temp(1), "BILLNO"

    Coll.Add Splited(77 + Index), "BILNAME"
    Coll.Add Splited(78 + Index), "BILADRS1"
    Coll.Add Splited(79 + Index), "BILADRS2"
    Coll.Add Splited(80 + Index), "BILADRS3"
    Coll.Add Splited(81 + Index), "SAVAMNT"
    
    temp = SplitWithLengthsPlus(Splited(82 + Index), 4, 3)
    Coll.Add temp(0), "RROID"
    Coll.Add temp(1), "RROCURR"
    Coll.Add temp(2), "RROAMT"
    
    Coll.Add Splited(83 + Index), "PAXTEL"
    Coll.Add Splited(84 + Index), "PAXFFT"
    
    
    temp3 = SplitForce(Splited(85 + Index), "-", 2)
    
    temp = SplitTwoReverse(temp3(0), 2)
    temp2 = SplitWithLengthsPlus(temp(0), 4, 2)
        Coll.Add temp2(0), "FLNTAG"
        Coll.Add temp2(1), "PSGID"
        Coll.Add temp2(2), "PSGASS"
    Coll.Add temp(1), "PSGNBR"
    
    Coll.Add temp3(1), "PSGNBR"

Set U_Hotel = Coll
End Function

Private Function ToPenDate(TheDate As String, Optional default = Empty) As Date
On Error GoTo errPara
Dim Day, Month, Year, TESTDATE
Day = Left(TheDate, 2)
Month = Right(TheDate, 3)
'by Abhi on 26-Aug-2009 for Date picking logic without year
'Year = mYearID
Year = Year(Date)
TESTDATE = CDate(Day & "/" & Month & "/" & Year)
If (TESTDATE < CDate(Format(Date, "dd-MMM-yyyy"))) Then
    Year = Year + 1
    TESTDATE = CDate(Day & "/" & Month & "/" & Year)
End If
ToPenDate = TESTDATE
Exit Function
errPara:
ToPenDate = default
End Function
Private Function ToPenDateX(TheDate, Optional default = Empty) As Date
On Error GoTo errPara
Dim Day, Month, Year, TESTDATE
Day = Left(TheDate, 2)
Month = Mid(TheDate, 3, 3)
Year = Right(TheDate, 2)
TESTDATE = CDate(Day & "/" & Month & "/" & Year)
ToPenDateX = TESTDATE
Exit Function
errPara:
    ToPenDateX = default
End Function


Private Function PEN(Data1 As String) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Splited = SplitForce(Data, "\", 9)
    
    Temp1 = ExtractBetween(Data, "AC-", "\")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "AC"
    Coll.Add Left(Temp1, 50), "AC"
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    If Trim(Left(Temp1, 50)) <> "" Then
        GIT_CUSTOMERUSERCODE_String = Left(Temp1, 50)
    End If
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    
    Temp1 = ExtractBetween(Data, "DES-", "\")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "DES"
    Coll.Add Left(Temp1, 50), "DES"
        
    Temp1 = ExtractBetween(Data, "REFE-", "\")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "REFE"
    Coll.Add Left(Temp1, 50), "REFE"
    
    Temp1 = ExtractBetween(Data, "SELL-ADT-", "\")
    Coll.Add Temp1, "SELLADT"
    
    Temp1 = ExtractBetween(Data, "SELL-CHD-", "\")
    Coll.Add Temp1, "SELLCHD"
    
    Temp1 = ExtractBetween(Data, "SCHARGE-", "\")
    Coll.Add Temp1, "SCHARGE"
    
    Temp1 = ExtractBetween(Data, "DEPT-", "\")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "DEPT"
    Coll.Add Left(Temp1, 50), "DEPT"
    
    Temp1 = ExtractBetween(Data, "BRANCH-", "\")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "BRANCH"
    Coll.Add Left(Temp1, 50), "BRANCH"
    
    Temp1 = ExtractBetween(Data, "CC-", "\")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "CUSTCC"
    Coll.Add Left(Temp1, 200), "CUSTCC"
    
    Set PEN = Coll
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

Private Function KS(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = SplitForce(Data, ";", 15)
    
    temp = SplitWithLengths(Splited(0), 1, 3, 11)
    Coll.Add temp(0), "ISSUE"
    Coll.Add temp(1), "CURRTYP"
    Coll.Add temp(2), "BASE"
    
    temp = SplitWithLengths(Splited(1), 3, 11)
    Coll.Add temp(0), "CURCODE"
    Coll.Add temp(1), "EQVAMT"
    
    temp = SplitWithLengths(Splited(2), 10, 1, 3, 7, 3, 2)
    Coll.Add temp(0), "TAXFLD"
    Coll.Add temp(1), "OTAXI"
    Coll.Add temp(2), "TCURR"
    Coll.Add temp(3), "TAXA"
    Coll.Add temp(4), "TAXC"
    Coll.Add temp(5), "TAXN"
    
    temp = SplitWithLengths(Splited(12), 3, 11) ' temp = SplitWithLengths(splited(3), 3, 11) 'removed by abhi on 25-Sep-2007 for sell amt KS
    Coll.Add temp(0), "CURRCDE"
    Coll.Add temp(1), "TOTAMT"

Coll.Add Splited(4), "SELLBUY1"
Coll.Add Splited(5), "SELLBUY2"
Coll.Add Splited(6), "TRANCURR"
'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue
Coll.Add vPSGRID_String, "PSGRID"
'by Abhi on 24-Jun-2014 for caseid 3667 Amadeus Fare line picking issue

Set KS = Coll
End Function


'by Abhi on 23-Jun-2010 for caseid 1401 Amadeus U-Hotel-HTL Segment
Private Function U_HotelHTLM(Data As String) As Collection
Dim Coll As New Collection
Dim Splited, Index
Dim temp, temp2, temp3
Dim sTemp1, sTemp2, sTemp3
Index = 0
Splited = SplitForce(Data, ";", 12)

    temp = SplitTwo(Splited(0), 3)
    Coll.Add temp(0), "TTONBR"
    Coll.Add temp(1), "DATAMODI"
    
    temp2 = SplitForce(Splited(1), " ", 4)
    
    temp = SplitWithLengths(temp2(0), 3, 3)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "SEGID"
    
    temp = temp2(1)
    Coll.Add temp, "AIRCDE"
    
    temp = SplitWithLengths(temp2(2), 2, 2)
    Coll.Add temp(0), "STATUS"
    Coll.Add temp(1), "NBRRMS"
    
    temp = SplitWithLengths(Splited(2), 4, 1, 5)
       Coll.Add temp(0), "PNRSTAT"
       Coll.Add temp(2), "INDTE"
    
    temp = Splited(3)
    Coll.Add temp, "OUTDTE"

    temp = Splited(4)
    Coll.Add temp, "AIRPT"

    temp = Splited(5)
    Coll.Add temp, "CDETRAN"
    temp = Splited(6)
    Coll.Add temp, "FREE"
    Coll.Add "", "FLNTAG"
    Coll.Add "", "PSGID"
    Coll.Add "", "PSGASS"
    Coll.Add "", "PSGNBR"

Set U_HotelHTLM = Coll
End Function

'by Abhi on 12-Mar-2014 for caseid 3791 EMD number picking from amadeus file
Private Function TMC(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
'by Abhi on 18-Dec-2015 for caseid 5894 Amadeus EMD ticket is duplicated in GDS intray
Dim vLastItem_Integer As Integer
'by Abhi on 18-Dec-2015 for caseid 5894 Amadeus EMD ticket is duplicated in GDS intray

Splited = SplitForce(Data, ";", 2)
    
    temp = SplitWithLengths(Splited(0), 1, 3, 1, 10)
    Coll.Add temp(0), "TKTTYPE"
    Coll.Add temp(1), "NUMAIR"
    'Coll.Add temp(2), "-"
    Coll.Add temp(3), "TKTNBR"
    
    temp = SplitWithLengths(Splited(1), 13, 1, 50, 2)
    Coll.Add temp(0), "ICWNBR"
    'Coll.Add temp(1), "/"
    Coll.Add temp(2), "FREE"
    Coll.Add temp(3), "LINEID"

    'by Abhi on 18-Dec-2015 for caseid 5894 Amadeus EMD ticket is duplicated in GDS intray
    vLastItem_Integer = UBound(Splited)
    If Left(Splited(vLastItem_Integer), 1) = "D" Then
        temp = SplitWithLengths(Splited(vLastItem_Integer), 1, 3)
        Coll.Add temp(0), "TSMID"
        Coll.Add Format(temp(1), "00"), "TSMNBR"
    'by Abhi on 24-Feb-2017 for caseid 7225 Amadeus Field does not exist in Collection TSMNBR
    Else
        Coll.Add "", "TSMID"
        Coll.Add "", "TSMNBR"
    'by Abhi on 24-Feb-2017 for caseid 7225 Amadeus Field does not exist in Collection TSMNBR
    End If
    'by Abhi on 18-Dec-2015 for caseid 5894 Amadeus EMD ticket is duplicated in GDS intray
    'by Abhi on 16-Aug-2016 for caseid 6673 changes in Amadeus EMD loading-Passenger number
    Coll.Add TKT_NoIndex, "PSGRNBR"
    'by Abhi on 16-Aug-2016 for caseid 6673 changes in Amadeus EMD loading-Passenger number

Set TMC = Coll
End Function
'by Abhi on 12-Mar-2014 for caseid 3791 EMD number picking from amadeus file

'by Abhi on 27-Mar-2014 for caseid 3853 Amadeus invoice number to reference field in folder
Private Function v(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

Splited = SplitForce(Data, "-", 2)
    Coll.Add Splited(0), "INVNBR"
    Coll.Add Splited(1), "CINVNUM"
    
Set v = Coll
End Function
'by Abhi on 27-Mar-2014 for caseid 3853 Amadeus invoice number to reference field in folder

'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation
Private Function RC_MARKUP(Data1, ByVal pPNO_Long As Long) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add pPNO_Long, "PNO"
    Coll.Add "", "SELLINF"
    Splited = SplitForce(Data, "/", 7)
    Coll.Add Val(Splited(1)), "FARE"
    Coll.Add Val(Splited(2)), "TAXES"
    Coll.Add Val(Splited(3)), "MARKUP"
    Coll.Add "0", "PENFARESELL"
    Coll.Add Left(Splited(4), 50), "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add Left(Splited(5), 50), "PENFARESUPPID"
    Coll.Add Left(Splited(6), 50), "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Data1 = Data
Set RC_MARKUP = Coll
End Function
'by Abhi on 29-Apr-2014 for caseid 3972 Sell value or commission calculation

'by Abhi on 30-Apr-2014 for caseid 3974 Customer COde Picking Amadeus file
Private Function RMSUBAGENCY(ByVal Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
Dim vi As Long
    Splited = SplitForce(Data, "SUBAGENCY_CODE:", 2)
    Coll.Add Left(Splited(1), 50), "AC"
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    If Trim(Left(Splited(1), 50)) <> "" Then
        GIT_CUSTOMERUSERCODE_String = Left(Splited(1), 50)
    End If
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Data1 = Data
Set RMSUBAGENCY = Coll
End Function
'by Abhi on 30-Apr-2014 for caseid 3974 Customer COde Picking Amadeus file

'by Abhi on 14-May-2014 for caseid 3787 Amadeus - Add the ATF value to Fare Amount
Private Function Q_ATF(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

temp = SplitForce(Data, ";", 2)
    Coll.Add Right(temp(0), Len(temp(0)) - 3), "FEE"
Set Q_ATF = Coll
End Function
'by Abhi on 14-May-2014 for caseid 3787 Amadeus - Add the ATF value to Fare Amount

'by Abhi on 29-May-2014 for caseid 4104 EMD Value picking for Amadeus
Private Function EMD(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Splited = SplitForce(Data, ";", 2)
        
    Coll.Add Splited(0), "TTONBR"
    temp = SplitWithLengths(Splited(1), 3, 3)
        Coll.Add temp(0), "ELMNBR"
        Coll.Add temp(1), "AIRCDE"
    Coll.Add Splited(2), "NUMAIR"
    Coll.Add Splited(3), "ISSAIR"
    Coll.Add Splited(4), "PRCNDATE"
    temp = SplitWithLengths(Splited(5), 1, 1)
        Coll.Add temp(0), "TSDID"
        Coll.Add temp(1), "TSDNBR"
    Coll.Add Splited(6), "EMDCPN"
    Coll.Add Splited(7), "OPECARM"
    Coll.Add Splited(8), "OPECARS"
    Coll.Add Splited(9), "FEEOWN"
    Coll.Add Splited(10), "ORIGIN"
    Coll.Add Splited(11), "DESTIN"
    temp = SplitWithLengths(Splited(12), 3, 62)
        Coll.Add temp(0), "TOTAG"
        Coll.Add temp(1), "TOINFO"
    temp = SplitWithLengths(Splited(13), 3, 62)
        Coll.Add temp(0), "ATTAG"
        Coll.Add temp(1), "ATINFO"
    Coll.Add Splited(14), "EMDTYPE"
    Coll.Add Splited(15), "REASON"
    Coll.Add Splited(16), "RFICDESC"
    Coll.Add Splited(17), "RFISC"
    Coll.Add Splited(18), "RFISCDESC"
    Coll.Add Splited(19), "DOCRMKS"
    Coll.Add Splited(20), "SVCRMKS"
    Coll.Add Splited(21), "NVBDATE"
    Coll.Add Splited(22), "NVADATE"
    Coll.Add Splited(23), "XSBTOTAL"
    Coll.Add Splited(24), "XSBQUALIF"
    temp = SplitWithLengths(Splited(25), 3, 11)
        Coll.Add temp(0), "CURCODEXSB"
        Coll.Add temp(1), "XSBRATE"
    temp = SplitWithLengths(Splited(26), 3, 9)
        Coll.Add temp(0), "CPVALTAG"
        Coll.Add temp(1), "CPNVALUE"
    Coll.Add Splited(27), "ISSUE"
    temp = SplitWithLengths(Splited(28), 3, 11)
        Coll.Add temp(0), "CURCODEBASE"
        Coll.Add temp(1), "BASE"
    Coll.Add Splited(29), "TAX"
    temp = SplitWithLengths(Splited(30), 3, 11)
        Coll.Add temp(0), "CURCODEEQVAMNT"
        Coll.Add temp(1), "EQVAMNT"
    temp = SplitWithLengths(Splited(31), 3, 11)
        Coll.Add temp(0), "CURCODEEXCHVAL"
        Coll.Add temp(1), "EXCHVAL"
    temp = SplitWithLengths(Splited(32), 3, 11)
        Coll.Add temp(0), "CURCODERFNDVAL"
        Coll.Add temp(1), "RFNDVAL"
    temp = SplitWithLengths(Splited(33), 2, 1, 3, 9, 4)
        Coll.Add temp(0), "TAXTAG"
        Coll.Add temp(1), "OTAXI"
        Coll.Add temp(2), "TCURR"
        Coll.Add temp(3), "TAXA"
        Coll.Add temp(4), "TAXT"

Set EMD = Coll
End Function
'by Abhi on 29-May-2014 for caseid 4104 EMD Value picking for Amadeus

'by Abhi on 28-Nov-2014 for caseid 4776 Amadeus Car segments
Private Function U_Car(Data As String) As Collection
Dim Coll As New Collection
Dim Splited, Index
Dim temp, temp2, temp3
Dim sTemp1, sTemp2, sTemp3
Index = 0
Splited = SplitForce(Data, ";", 32) 'upto ESNBDY

    temp = SplitTwo(Splited(0), 3)
    Coll.Add temp(0), "TTONBR"
    Coll.Add temp(1), "DATAMODI"
    
    temp2 = SplitForce(Splited(1), " ", 3)
    
    temp = SplitWithLengths(temp2(0), 3, 3)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "SEGID"
    
    temp = temp2(1)
    Coll.Add temp, "AIRCDE"
    
    temp = SplitWithLengths(temp2(2), 2, 2)
    Coll.Add temp(0), "STATUS"
    Coll.Add temp(1), "NBRCARS"
    
    temp = SplitWithLengths(Splited(2), 4, 1, 5)
       Coll.Add temp(0), "PNRSTAT"
       Coll.Add temp(2), "PICKUP"
    

    temp = Splited(3)
    Coll.Add temp, "DROPOFF"

    temp = Splited(4)
    Coll.Add temp, "AIRPT"

    temp = Splited(5)
    Coll.Add temp, "CDETRAN"

    temp = Splited(6)
    Coll.Add temp, "CAR"

    temp = Splited(7)
    Coll.Add temp, "CARTRANS"

    temp = SplitWithLengths(Splited(8), 4, 36)
    Coll.Add temp(0), "ARRID"
    Coll.Add temp(1), "ARROPT"

    temp = SplitWithLengths(Splited(9), 3, 6)
    Coll.Add temp(0), "RTID"
    Coll.Add temp(1), "RTTME"

    temp = SplitWithLengths(Splited(10), 3, 12)
    Coll.Add temp(0), "BAID"
    Coll.Add temp(1), "BAOPT"

    temp = SplitWithLengths(Splited(11), 3, 25)
    Coll.Add temp(0), "BNID"
    Coll.Add temp(1), "BNOPT"

    temp = SplitWithLengths(Splited(12), 3, 24)
    Coll.Add temp(0), "BRID"
    Coll.Add temp(1), "BROPT"

    temp = SplitWithLengths(Splited(13), 3, 8)
    Coll.Add temp(0), "BKGID"
    Coll.Add temp(1), "AIRIATA"

    temp = SplitWithLengths(Splited(14), 3, 20)
    Coll.Add temp(0), "CDID"
    Coll.Add temp(1), "CDOPT"

    temp = SplitWithLengths(Splited(15), 3, 16)
    Coll.Add temp(0), "CONFID"
    Coll.Add temp(1), "CONFNBR"

    temp = SplitWithLengths(Splited(16), 3, 20)
    Coll.Add temp(0), "CKID"
    Coll.Add temp(1), "CKOPT"

    temp = SplitWithLengths(Splited(17), 3, 16)
    Coll.Add temp(0), "DCID"
    Coll.Add temp(1), "DCOPT"

    temp = SplitWithLengths(Splited(18), 3, 60)
    Coll.Add temp(0), "DOID"
    Coll.Add temp(1), "DOOPT"

    temp = SplitWithLengths(Splited(19), 3, 44)
    Coll.Add temp(0), "FPID"
    Coll.Add temp(1), "FPOPT"

    temp = SplitWithLengths(Splited(20), 3, 28)
    Coll.Add temp(0), "FTID"
    Coll.Add temp(1), "FTOPT"

    temp = SplitWithLengths(Splited(21), 2, 42)
    Coll.Add temp(0), "GID"
    Coll.Add temp(1), "GOPT"

    temp = SplitWithLengths(Splited(22), 3, 20)
    Coll.Add temp(0), "IDID"
    Coll.Add temp(1), "IDOPT"

    temp = SplitWithLengths(Splited(23), 3, 20)
    Coll.Add temp(0), "ITID"
    Coll.Add temp(1), "ITOPT"

    temp = SplitWithLengths(Splited(24), 3, 6)
    Coll.Add temp(0), "LCID"
    Coll.Add temp(1), "LCOPT"

    temp = SplitWithLengths(Splited(25), 3, 6)
    Coll.Add temp(0), "MKID"
    Coll.Add temp(1), "MKOPT"

    temp = SplitWithLengths(Splited(26), 3, 25)
    Coll.Add temp(0), "NMID"
    Coll.Add temp(1), "NMOPT"

    temp = SplitWithLengths(Splited(27), 4, 60)
    Coll.Add temp(0), "PUPID"
    Coll.Add temp(1), "PUPOPT"

    temp = SplitWithLengths(Splited(28), 3, 24)
    Coll.Add temp(0), "RBID"
    Coll.Add temp(1), "RBOPT"

    temp = SplitWithLengths(Splited(29), 3, 9)
    Coll.Add temp(0), "RCID"
    Coll.Add temp(1), "RCOPT"

    temp = SplitWithLengths(Splited(30), 3, 59)
    Coll.Add temp(0), "RGID"
    Coll.Add temp(1), "RGOPT"

    temp = SplitWithLengths(Splited(31), 3, 59)
    Coll.Add temp(0), "RQID"
    Coll.Add temp(1), "RQOPT"

    temp2 = SplitForce(Splited(32), "+", 3)
    temp = SplitWithLengths(temp2(0), 3, 3, 12)
    Coll.Add temp(0), "ESID"
    Coll.Add temp(1), "ESCUCDE"
    Coll.Add temp(2), "ESTTLAMT"
    Coll.Add temp2(1), "ESNBDY"
    
Set U_Car = Coll
End Function
'by Abhi on 28-Nov-2014 for caseid 4776 Amadeus Car segments

'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
Private Function RMPEN_PENLINEPENAIRTKT(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
Dim vi As Long
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'splited = SplitForce(Data, "/", 3)
    'Temp1 = Replace(splited(1), "PAX-", "")
    '    Coll.Add Left(Temp1, 150), "AIRTKTPAX"
    'Temp1 = Replace(splited(2), "TKT-", "")
    '    Coll.Add Left(Temp1, 50), "AIRTKTTKT"
    'Temp1 = Replace(splited(3), "DATE-", "")
    '    Coll.Add Left(Temp1, 9), "AIRTKTDATE"
    Temp1 = ExtractBetween(Data, "PAX-", "/")
        Coll.Add Left(Temp1, 150), "AIRTKTPAX"
    Temp1 = ExtractBetween(Data, "TKT-", "/")
        Coll.Add Left(Temp1, 50), "AIRTKTTKT"
    Temp1 = ExtractBetween(Data, "DATE-", "/")
        Coll.Add Left(Temp1, 9), "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Data1 = Data
Set RMPEN_PENLINEPENAIRTKT = Coll
End Function
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS

'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
Private Function RMPEN_PENLINEPENAGROSS(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    Coll.Add "", "TKTDEADLINE"
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'splited = SplitForce(Data, "/", 5)
    Splited = SplitForce(Data, "/", 6)
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add Val(Splited(1)), "PENAGROSSADULT"
    Coll.Add Val(Splited(2)), "PENAGROSSCHILD"
    Coll.Add Val(Splited(3)), "PENAGROSSINFANT"
    Coll.Add Val(Splited(4)), "PENAGROSSPACKAGE"
    'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add Val(Splited(5)), "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Data1 = Data
Set RMPEN_PENLINEPENAGROSS = Coll
End Function
'by Abhi on 15-Sep-2015 for caseid 3917 Penline for PENAGROSS for Amadeus

'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
Private Function RMPEN_PENLINEPENBILLCUR(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    Coll.Add "", "TKTDEADLINE"
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    Coll.Add 0, "PENAGROSSYOUTH"
    Coll.Add "", "PDOB"
    Splited = SplitForce(Data, "-", 2)
    Coll.Add Left(Splited(1), 4), "PENBILLCUR"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Data1 = Data
Set RMPEN_PENLINEPENBILLCUR = Coll
End Function
'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency

'by Abhi on 18-Jul-2016 for caseid 6573 Amadeus -Ticket date picking logic for farelogix file
Private Function FS(Data As String) As Collection
Dim Coll As New Collection
Dim Splited

Splited = SplitTwo(Data, 127)
    Coll.Add Splited(0), "MISC"
Set FS = Coll
End Function
'by Abhi on 18-Jul-2016 for caseid 6573 Amadeus -Ticket date picking logic for farelogix file

'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
Private Function RMPEN_PENLINEPENWAIT(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    Coll.Add "", "TKTDEADLINE"
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    Coll.Add 0, "PENAGROSSYOUTH"
    Coll.Add "", "PDOB"
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
        
    GIT_PENWAIT_String = "Y"
'Data1 = Data
Set RMPEN_PENLINEPENWAIT = Coll
End Function
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT

'by Abhi on 07-Aug-2017 for caseid 7674 Form of payment mapping for EMD tickets for Amadeus
Private Function MFP(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
'Dim sTemp1, sTemp2, sTemp3

    Splited = SplitForce(Data, ";", 4)
    
    Coll.Add Left(Splited(0), 500), "PAYMNT"
    Coll.Add Left("", 99), "LINEASS"
    Coll.Add Left("", 50), "LINENBR"
    Coll.Add Left("", 50), "LINENBR2"
    Coll.Add Left("", 99), "PSGASS"
    Coll.Add Left("", 50), "PSGNBR"
    Coll.Add Left("", 50), "PSGNBR2"
    temp = SplitWithLengths(Splited(3), 1, 3)
    Coll.Add Left(temp(0), 2), "TSMID"
    Coll.Add Left(Format(temp(1), "00"), 3), "TSMNBR"

Set MFP = Coll
End Function
'by Abhi on 07-Aug-2017 for caseid 7674 Form of payment mapping for EMD tickets for Amadeus

'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
Private Function RMPEN_PENLINEPENVC(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    Coll.Add "", "TKTDEADLINE"
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    Coll.Add 0, "PENAGROSSYOUTH"
    Coll.Add "", "PDOB"
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

'Data1 = Data
Set RMPEN_PENLINEPENVC = Coll
End Function
'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus

'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
Private Function ToPenDateYYMMDD(ByVal TheDate As String, Optional default = Empty) As Date
On Error GoTo errPara
Dim Day, Month, Year, TESTDATE As Date, TempDate
Dim mDate As Date

If IsDate(default) = True Then
    mDate = CDate(default)
Else
    mDate = Date
End If

TempDate = Trim(TheDate)
Day = Right(TempDate, 2)
Month = Mid(TempDate, 3, 2)
Year = Left(DateTime.Year(Date), 2) & Left(TempDate, 2) '  mYearID
'Year = Left(DateTime.Year(DateAppTimeZone), 2) & Left(TempDate, 2) '  mYearID
TESTDATE = CDate(Day & "/" & Month & "/" & Year)
'If (TESTDATE < CDate(Format(Date, "dd-MMM-yyyy"))) Then
'    Year = Year + 1
'    TESTDATE = CDate(Day & "/" & Month & "/" & Year)
'End If
ToPenDateYYMMDD = TESTDATE
Exit Function
errPara:
ToPenDateYYMMDD = IIf(default = "", Date, default)
End Function
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date

'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
Private Function RMPEN_PENLINEPENCS(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "", "SELLINF"
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add "0", "MARKUP"
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add 0, "PENAGENTCOMSUM"
    Coll.Add 0, "PENAGENTCOMVAT"
    Coll.Add "", "TKTDEADLINE"
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    Coll.Add 0, "PENAGROSSYOUTH"
    Coll.Add "", "PDOB"
    Coll.Add "", "PENBILLCUR"
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    Coll.Add "", "PENFARETICKETTYPE"
    Splited = SplitForce(Data, "/", 3)
    Coll.Add Left(Trim(Splited(1)), 20), "PENCSLABELID"
    Coll.Add Left(Trim(Splited(2)), 20), "PENCSLISTID"

Data1 = Data
Set RMPEN_PENLINEPENCS = Coll
End Function
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

'by Abhi on 25-Apr-2019 for caseid 10192 Amadeus Refund File reading
Private Function RFD(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
'Dim sTemp1, sTemp2, sTemp3

    Splited = SplitForce(Data, ";", 14)
    
    Coll.Add Left(Splited(0), 1), "DSOURCE"
    Coll.Add Left(Splited(1), 7), "DOCDATE"
    Coll.Add Left(Splited(2), 1), "ITNTYP"
    temp = SplitWithLengths(Splited(3), 3, 11)
        Coll.Add Left(temp(0), 3), "CURRTYP"
        Coll.Add Left(temp(1), 11), "FPAID"
    Coll.Add Left(Splited(4), 11), "FUSED"
    Coll.Add Left(Splited(5), 11), "FRFD"
    Coll.Add Left(Splited(6), 2), "NRFDID"
    Coll.Add Left(Splited(7), 11), "NRFD"
    Coll.Add Left(Splited(8), 11), "CXFEE"
    Coll.Add Left(Splited(9), 11), "CXFEECOM"
    Coll.Add Left(Splited(10), 11), "MSFEE"
    temp = SplitWithLengths(Splited(11), 2, 11)
        Coll.Add Left(temp(0), 2), "TAXCDE"
        Coll.Add Left(temp(1), 11), "TAXRFD"
    Coll.Add Left(Splited(12), 11), "TOTRFD"
    Coll.Add Left(Splited(13), 7), "DEPDTE"

Set RFD = Coll
End Function

Private Function R(Data As String) As Collection
Dim Coll As New Collection
Dim Splited
Dim temp, temp2
'Dim sTemp1, sTemp2, sTemp3

    Splited = SplitForce(Data, ";", 2)
    
    Coll.Add Left(Splited(0), 14), "RFDNBR"
    Coll.Add Left(Splited(1), 7), "RFDDATE"

Set R = Coll
End Function
'by Abhi on 25-Apr-2019 for caseid 10192 Amadeus Refund File reading
