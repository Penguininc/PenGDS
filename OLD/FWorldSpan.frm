VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
         MaxLength       =   80
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
         MaxLength       =   80
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


Private Sub CmdClear_Click()
ClearAll
End Sub

Private Sub CmdDelete_Click()
On Error GoTo Note
    Dim Count As Integer
    File1.Path = txtDirName.Text
    File1.Pattern = "*.mir"
       
    If txtDirName.Text = "" Or File1.ListCount = 0 Then MsgBox "AIR Files Not Found", vbInformation: Exit Sub
    If MsgBox("Do you want to Delete the " & File1.ListCount & " MIR Files ?", vbYesNo) = vbYes Then
        For Count = 0 To File1.ListCount - 1
            txtFileName.Text = File1.Path & "\" & File1.List(Count)
            DeleteFileToRecycleBin (txtFileName)
        Next
    End If
Exit Sub
Note:
    MsgBox Err.Description & ". Invalid Drive Location", vbCritical: Exit Sub

End Sub

Public Sub cmdDirUpload_Click()
'On Error GoTo Note
    Dim Count As Integer
    
    
    File1.Path = txtDirName.Text
    File1.Pattern = "*.prt"
    If txtDirName.Text = "" Or File1.ListCount = 0 Then FMain.stbUpload.Panels(1).Text = "MIR Files Not Found...": Exit Sub
    'Uploading Each File
    Open (App.Path & "\_UploadingSQL_") For Random As #1
    Close #1
    FMain.cmdStop.Enabled = False
    DoEvents
    FMain.lblFName.Caption = "Uploading Galilieo..."
    FMain.stbUpload.Panels(1).Text = "Reading...."
    FMain.prgUpload.Max = IIf((File1.ListCount = 1), File1.ListCount, File1.ListCount - 1)
    Pgr.value = 5
    
    For Count = 0 To File1.ListCount - 1

        DoEvents
        txtFileName.Text = File1.Path & "\" & File1.List(Count)
        FMain.stbUpload.Panels(2).Text = File1.List(Count)
        FNAME = File1.List(Count)
        AddFile txtFileName.Text, FNAME
        If Pgr.value < 90 Then
            Pgr.value = Pgr.value + 10
        End If
        FMain.prgUpload.value = Count
    Next
    Pgr.value = 100
    File1.Refresh
    FMain.prgUpload.value = 0
    FMain.stbUpload.Panels(1).Text = "Process Completed."
    FMain.cmdStop.Enabled = True
    FMain.stbUpload.Panels(2).Text = ""
    FMain.stbUpload.Panels(1).Text = "Waiting...."
    'FMain.txtGalilieoSource = ""
    'FMain.txtGalilieoDest = ""
    FMain.lblFName.Caption = ""
    fsObj.DeleteFile (App.Path & "\_UploadingSQL_"), True
    Me.Hide
Exit Sub
Note:
    Me.Hide
    Exit Sub
    Resume
End Sub

Private Sub Command1_Click()
CommonDialog1.fileName = ""
CommonDialog1.ShowOpen

If CommonDialog1.fileName = "" Then
Exit Sub
End If
Text1.Text = CommonDialog1.fileName
End Sub

Private Sub Command2_Click()
AddFile Text1.Text
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
'dbCompany.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=sa;Initial Catalog=PENDATA001;Data Source=PEN1\PENSOFT"
On Error GoTo Note
    Dim rsSelect As New ADODB.Recordset
    Me.Icon = FMain.Icon
    rsSelect.Open "Select WSPUPLOADDIRNAME,WSPDESTDIRNAME From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    If rsSelect.EOF = False Then
        txtDirName = IIf(IsNull(rsSelect!WSPUPLOADDIRNAME), "", rsSelect!WSPUPLOADDIRNAME)
        txtRelLocation = IIf(IsNull(rsSelect!WSPDESTDIRNAME), "", rsSelect!WSPDESTDIRNAME)
    End If
    rsSelect.Close
    If rsSelect.State = 1 Then rsSelect.Close
    rsSelect.Open "Select WSPSTATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    chkGal.value = IIf((IsNull(rsSelect!WSPSTATUS) = True), 0, rsSelect!WSPSTATUS)
    File1.Path = txtDirName.Text
    File1.Refresh
    File1.Pattern = "*.prt"
    File1.Refresh
    Exit Sub
Note:
    If Me.Visible = True Then Unload Me

End Sub

Private Sub Text1_Change()
Caption = Len(Text1)
End Sub



Public Function AddFile(fileName As String, Optional fTitle = "")
Dim id As Long
Dim head As String
Dim temp
Dim UploadNo As Long
Dim LineStr  As String
Dim tsObj As TextStream
Dim Lines, crrItem As String
Dim FileObj As New FileSystemObject
    WorldSpan_AirFareSLNO = 0
    Set tsObj = fsObj.OpenTextFile(fileName, ForReading)
    UploadNo = mSeqNumberGen("UPLOADNO")
    Do While tsObj.AtEndOfStream <> True
        DoEvents
        LineStr = tsObj.ReadLine
        Lines = Split(LineStr, Chr(13))
        For j = 0 To UBound(Lines)
            crrItem = Lines(j)
            WorldSpan.PostLine crrItem, UploadNo, fTitle
        Next
    Loop
    tsObj.Close
    
    FileObj.CopyFile txtFileName.Text, txtRelLocation.Text & "\", True
On Error Resume Next
    FileObj.DeleteFile txtFileName.Text, True
End Function




