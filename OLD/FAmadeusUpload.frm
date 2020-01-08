VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
            DeleteFileToRecycleBin (txtFileName.Text)
        Next
    End If
Exit Sub
Note:
    MsgBox Err.Description & ". Invalid Drive Location", vbCritical: Exit Sub

End Sub

Public Sub cmdDirUpload_Click()
On Error GoTo Note
    Dim Count As Integer
    
    
    File1.Path = txtDirName.Text
    File1.Pattern = "*.air;*.txt"
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
    
'    If File1.ListCount > 0 Then
'        If MsgBox("Do You Wish To Clear First ?", vbQuestion + vbYesNo) = vbYes Then
'            ClearAll
'        End If
'    End If
    For Count = 0 To File1.ListCount - 1

        DoEvents
        txtFileName.Text = File1.Path & "\" & File1.List(Count)
        FMain.stbUpload.Panels(2).Text = File1.List(Count)
        FNAME = File1.List(Count)
        AddFile txtFileName.Text
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

Private Sub Form_Load()
'dbCompany.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=sa;Initial Catalog=PENDATA001;Data Source=PEN1\PENSOFT"
On Error GoTo Note
    Dim rsSelect As New ADODB.Recordset
    Me.Icon = FMain.Icon
    rsSelect.Open "Select AMDUPLOADDIRNAME,AMDDESTDIRNAME From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    If rsSelect.EOF = False Then
        txtDirName = IIf(IsNull(rsSelect!AMDUPLOADDIRNAME), "", rsSelect!AMDUPLOADDIRNAME)
        txtRelLocation = IIf(IsNull(rsSelect!AMDDESTDIRNAME), "", rsSelect!AMDDESTDIRNAME)
    End If
    rsSelect.Close
    If rsSelect.State = 1 Then rsSelect.Close
    rsSelect.Open "Select AMDSTATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    chkGal.value = IIf((IsNull(rsSelect!AMDSTATUS) = True), 0, rsSelect!AMDSTATUS)
    File1.Path = txtDirName.Text
    File1.Refresh
    File1.Pattern = "*.air;*.txt"
    File1.Refresh
    Exit Sub
Note:
    If Me.Visible = True Then Unload Me

End Sub

Private Sub Text1_Change()
Caption = Len(Text1)
End Sub



Public Function BLK(Data) As Collection
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







Public Function SplitTwo(Data, First As Long)
On Error Resume Next
Dim temp(1) As String
temp(0) = Left(Data, First)
If Len(Data) > 0 Then
temp(1) = Mid(Data, First + 1, Len(Data) - First)
End If
SplitTwo = temp
End Function

Public Function SplitTwoReverse(Data, First As Long)
On Error Resume Next
Dim temp(1) As String
temp(1) = Right(Data, First)
If (Len(Data) > 0) Then
temp(0) = Mid(Data, 1, SkipNegative(Len(Data) - First))
End If
SplitTwoReverse = temp
End Function

Public Function SplitFirstTwo(Data, find As String)
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

Public Function SplitField(Data, Start As String, Finish As String)
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

Public Function SplitWithLengths(Data, ParamArray Lengths())
Dim Nos As Integer
Dim pos As Integer, j As Integer
Nos = UBound(Lengths)
Dim temp() As String
ReDim temp(Nos) As String
pos = 1

For j = 0 To Nos
    temp(j) = Mid(Data, pos, Lengths(j))
    pos = pos + Lengths(j)
Next
SplitWithLengths = temp
End Function
Public Function SplitWithLengthsPlus(Data, ParamArray Lengths())
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
Next
If Len(Data) > 0 Then
temp(Nos + 1) = Mid(Data, pos, SkipNull(Len(Data) - pos + 1))
End If
SplitWithLengthsPlus = temp

End Function

Public Function SplitForce(Data, delimiter As String, minNos As Integer)
Dim temp() As String
Dim Nos As Integer
temp = Split(Data, delimiter)
Nos = UBound(temp) + 1
If Nos < minNos Then
ReDim Preserve temp(minNos)
End If
SplitForce = temp
End Function

Public Function AMD(Data) As Collection
Dim Coll As New Collection
Dim Lines
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
Lines = SplitForce(Data, Chr(13) & Chr(10), 3)
splited = SplitForce(Lines(0), ";", 4)

'Line1
    temp = SplitWithLengths(splited(0), 1, 2, 8)
    Coll.Add temp(0), "RETRANS"
    Coll.Add temp(1), "DOM"
    Coll.Add temp(2), "AIRSEQ "
    
    temp = SplitForce(splited(1), "/", 2)
    Coll.Add temp(0), "AIRSEC"
    Coll.Add temp(1), "AIRTOTAL"
    
    temp = SplitTwo(splited(2), 4)
    Coll.Add temp(0), "ACTTAG"
    Coll.Add Left(temp(1), 5), "DATE"
    
    Coll.Add splited(3), "AGTSIN"
'Line 2
    splited = SplitForce(Lines(1), ";", 4)
    
    temp = SplitTwo(splited(0), 2)
    Coll.Add temp(0), "NETID"
    Coll.Add temp(1), "ADDRTKT"
    
    temp = SplitTwo(splited(1), 2)
    Coll.Add temp(0), "NETID_BKO"
    Coll.Add temp(1), "ADDRTKT_BKO"
    
    Coll.Add splited(2), "IDENTBOS"
    Coll.Add splited(3), "BOSTYPE"

'Line 3
splited = SplitForce(Lines(2), ";", 32)

'sec1
    temp = SplitWithLengths(splited(0), 3, 3, 6, 3)
    Coll.Add temp(0), "CTYID"
    Coll.Add temp(1), "AIRCDE"
    Coll.Add temp(2), "RECLOC"
    Coll.Add temp(3), "PNRENV"
'sec2
    temp = SplitWithLengths(splited(1), 2, 2, 1)
    Coll.Add temp(0), "TOTPSG"
    Coll.Add temp(1), "AIRPSG"
    Coll.Add temp(2), "NHPNR"

    Coll.Add splited(2), "IDBKG"
    Coll.Add splited(3), "BKGAGCY"
    Coll.Add splited(4), "IDFIRST"
    Coll.Add splited(5), "OWNAGY"
    Coll.Add splited(6), "IDENTCO"
    Coll.Add splited(7), "CURAGCY"
    Coll.Add splited(8), "IDENTT"
    Coll.Add splited(9), "TKTAGCY"
    Coll.Add splited(10), "OFFID"
    Coll.Add splited(11), "CPNID"
    Coll.Add splited(12), "IDENTT2"
    Coll.Add splited(13), "TKTAGCY2"
    Coll.Add splited(14), "OFFID2"
    Coll.Add splited(15), "CPNID2"
    Coll.Add splited(16), "IDENTT3"
    Coll.Add splited(17), "TKTAGCY3"
    Coll.Add splited(18), "OFFID3"
    Coll.Add splited(19), "CPNID3"
    Coll.Add splited(20), "IDENTT4"
    Coll.Add splited(21), "TKTAGCY4"
    Coll.Add splited(22), "OFFID4"
    Coll.Add splited(23), "CPNID4"
    Coll.Add splited(24), "IDENTT5"
    Coll.Add splited(25), "TKTAGCY5"
    Coll.Add splited(26), "OFFID5"
    Coll.Add splited(27), "CPNID5"
    Coll.Add splited(28), "ACCTNUM"
    Coll.Add splited(29), "ORDNUM"
    Coll.Add splited(30), "ERSP"
'sec 32
    sTemp1 = splited(31)
    sTemp2 = Right(sTemp1, 6)
    sTemp3 = Replace(sTemp1, sTemp2, "")
    temp = SplitWithLengths(sTemp3, 28, 3)
    
    Coll.Add temp(0), "NSPNR"
    Coll.Add temp(1), "NSAIR"
    Coll.Add sTemp2, "NSRECLOC"
Set AMD = Coll
End Function


Public Function A(Data As String) As Collection
Dim Coll As New Collection
Dim Lines As Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
splited = SplitForce(Data, ";", 2)

'sec 1
Coll.Add splited(0), "TKTAIR"
'sec 2
temp = SplitWithLengths(splited(1), 3, 3)
Coll.Add temp(0), "ALCDE"
Coll.Add temp(1), "NUMCDE"

Set A = Coll
End Function

Public Function B(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

splited = SplitForce(Data, ";", 2)
Coll.Add splited(0), "ENTRY"
Set B = Coll
End Function

Public Function C(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
splited = SplitTwo(Data, 5)
'Sec 1
temp = SplitWithLengths(splited(0), 4, 1)
    Coll.Add temp(0), "SVCCAR"
    Coll.Add temp(1), "TKTIND"
temp = SplitForce(splited(1), "-", 7)
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

Public Function D(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
splited = SplitForce(Data, ";", 3)
Coll.Add splited(0), "PNRDTE"
Coll.Add splited(1), "CHGDTE"
Coll.Add splited(2), "AIRDTE"
Set D = Coll
End Function


Public Function G(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
splited = Split(Data, ";")
Coll.Add splited(0), "INTIND"
Coll.Add splited(1), "SALEIND"
Coll.Add splited(2), "ORGDEST"
Coll.Add splited(3), "DOMISO"
Set G = Coll
End Function


Public Function H(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
splited = SplitForce(Data, ";", 23)

Coll.Add splited(0), "TTONBR"

temp = SplitWithLengthsPlus(splited(1), 3, 1)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "STOP"
    Coll.Add temp(2), "ORIGAIR"
Coll.Add splited(2), "ORIGCTY"
Coll.Add splited(3), "DESTAIR"
Coll.Add splited(4), "DESTCTY"

temp = SplitWithLengths(splited(5), 6, 5, 2, 2, 5, 5, 5, 5)
    Coll.Add temp(0), "AIRCDE"
    Coll.Add temp(1), "FLTNBR"
    Coll.Add temp(2), "CLSSVC"
    Coll.Add temp(3), "CLSBKG"
    Coll.Add temp(4), "DEPDTE"
    Coll.Add temp(5), "DEPTIM"
    Coll.Add temp(6), "ARRTIM"
    Coll.Add temp(7), "ARRDTE"
    
Coll.Add splited(6), "STATUS"
Coll.Add splited(7), "PNRSTAT"
Coll.Add splited(8), "MEAL"
Coll.Add splited(9), "NBRSTOP"
Coll.Add splited(10), "EQUIP"
Coll.Add splited(11), "ENTER"
Coll.Add splited(12), "SHRCOMM"
Coll.Add splited(13), "BAGALL"
Coll.Add splited(14), "CKINTRM"
Coll.Add splited(15), "CKINTIM "

Coll.Add splited(16), "TKTT"
Coll.Add splited(17), "FLTDUR"
Coll.Add splited(18), "NONSMOK"
Coll.Add splited(19), "MLTPM"
Coll.Add splited(20), "CPNNR"
Coll.Add splited(21), "CPNIN"
Coll.Add splited(22), "CPNSN"
Coll.Add splited(23), "CPNTSN"

Set H = Coll
End Function



Public Function getLineInfo(DataLine As String, id As Long)
Dim head  As String

'----------GENERAL---------------------
head = Left(DataLine, 2)
Select Case head
    Case "A-", "B-", "C-", "D-", "E-", "F-", "G-", "H-", "I-", "J-", "K-", "L-", "M-", "N-", "O-", "P-", "Q-", "R-", "S-", "T-", "U-", "V-", "W-", "X-", "Y-", "Z-", "N-", "O-", "FP", "FM", "TK", "FV", "RM"
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
    Case "TAX", "KFT", "SSR", "OSI", "TSA"
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


Public Function insertLineGeneral(DataLine As String, Ident As String, UploadNo As Long)
Dim head  As String
Dim Info As String
Dim Coll As New Collection
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
        Set Coll = D(Info)
        InsertData Coll, "AmdLineD", "D-", UploadNo
Case "G-"
        Set Coll = G(Info)
        InsertData Coll, "AmdLineG", "G-", UploadNo
Case "H-"
        Set Coll = H(Info)
        InsertData Coll, "AmdLineH", "H-", UploadNo
Case "I-"
        Set Coll = i(Info)
        InsertData Coll, "AmdLineI", "I-", UploadNo
Case "K-"
        Set Coll = k(Info)
        InsertData Coll, "AmdLineK", "K-", UploadNo
Case "Q-"
        Set Coll = Q(Info)
        InsertData Coll, "AmdLineQ", "Q-", UploadNo
Case "T-"
        Set Coll = T(Info)
        InsertData Coll, "AmdLineT", "T-", UploadNo
Case "U-"
        Set Coll = U(Info)
        InsertData Coll, "AmdLineU", "U-", UploadNo
Case "X-"
        Set Coll = X(Info)
        InsertData Coll, "AmdLineX", "X-", UploadNo
Case "Y-"
        Set Coll = Y(Info)
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
        InsertData Coll, "AmdLineTK", "TK", UploadNo
Case "FV"
        Set Coll = FV(Info)
        InsertData Coll, "AmdLineFV", "FV", UploadNo
Case "RM"
        Set Coll = RM(Info)
        InsertData Coll, "AmdLineRM", "RM", UploadNo
        
Case "TAX"
        Info = Right(DataLine, Len(DataLine) - 4)
        Set Coll = TAX(Info)
        InsertData Coll, "AmdLineTAX", "TAX-", UploadNo
Case "SSR"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = SSR(Info)
        InsertData Coll, "AmdLineSSR", "SSR", UploadNo
Case "KFT"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = KFT(Info)
        InsertData Coll, "AmdLineKFT", "KFT", UploadNo
Case "TSA"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = TSA(Info)
        InsertData Coll, "AmdLineTSA", "TSA", UploadNo
Case "OSI"
        Info = Right(DataLine, Len(DataLine) - 3)
        Set Coll = OSI(Info)
        InsertData Coll, "AmdLineOSI", "OSI", UploadNo
End Select
End Function




Public Function X(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
splited = SplitForce(Data, ";", 23)

Coll.Add splited(0), "TTONBR"

temp = SplitWithLengthsPlus(splited(1), 3, 1)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "STOP"
    Coll.Add temp(2), "ORIGAIR"
Coll.Add splited(2), "ORIGCTY"
Coll.Add splited(3), "DESTAIR"
Coll.Add splited(4), "DESTCTY"
temp = SplitWithLengthsPlus(splited(5), 6, 5, 2, 2, 5, 5, 5, 5)
    Coll.Add temp(0), "AIRCDE"
    Coll.Add temp(1), "FLTNBR"
    Coll.Add temp(2), "CLSSVC"
    Coll.Add temp(3), "CLSBKG"
    Coll.Add temp(4), "DEPDTE"
    Coll.Add temp(5), "DEPTIM"
    Coll.Add temp(6), "ARRTIM"
    Coll.Add temp(7), "ARRDTE"
    
Coll.Add splited(6), "STATUS"
Coll.Add splited(7), "PNRSTAT"
Coll.Add splited(8), "MEAL"
Coll.Add splited(9), "NBRSTOP"

Coll.Add splited(10), "EQUIP"
Coll.Add splited(11), "ENTER"
Coll.Add splited(12), "SHRCOMM"
Coll.Add splited(13), "BAGALL"
Coll.Add splited(14), "CKINTRM"
Coll.Add splited(15), "CKINTIM"
Coll.Add splited(16), "TKTT"
Coll.Add splited(17), "FLTDUR"
Coll.Add splited(18), "NONSMOK"

Set X = Coll

End Function

Public Function Y(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
splited = SplitForce(Data, ";", 14)

Coll.Add splited(0), "TTONBR"
    temp = SplitTwo(splited(1), 3)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "ORIGAIR"
Coll.Add splited(2), "ORIGCTY"
Coll.Add splited(3), "DESTAIR"
Coll.Add splited(4), "DESTCTY"
Coll.Add splited(5), "OWNER"
Coll.Add splited(6), "OWNNAME"
Coll.Add splited(7), "COCKPIT"
Coll.Add splited(8), "CPITNAME"
Coll.Add splited(9), "CABIN"
Coll.Add splited(10), "CABINNAME"
Coll.Add splited(11), "EQUIP"
Coll.Add splited(12), "GAUGE"

Set Y = Coll
End Function


Public Function k(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
splited = SplitForce(Data, ";", 15)
    
    temp = SplitWithLengths(splited(0), 1, 3, 11)
    Coll.Add temp(0), "ISSUE"
    Coll.Add temp(1), "CURRTYP"
    Coll.Add temp(2), "BASE"
    
    temp = SplitWithLengths(splited(1), 3, 11)
    Coll.Add temp(0), "CURCODE"
    Coll.Add temp(1), "EQVAMT"
    
    temp = SplitWithLengths(splited(2), 10, 1, 3, 7, 3, 2)
    Coll.Add temp(0), "TAXFLD"
    Coll.Add temp(1), "OTAXI"
    Coll.Add temp(2), "TCURR"
    Coll.Add temp(3), "TAXA"
    Coll.Add temp(4), "TAXC"
    Coll.Add temp(5), "TAXN"
    
    temp = SplitWithLengths(splited(3), 3, 11)
    Coll.Add temp(0), "CURRCDE"
    Coll.Add temp(1), "TOTAMT"

Coll.Add splited(4), "SELLBUY1"
Coll.Add splited(5), "SELLBUY2"
Coll.Add splited(6), "TRANCURR"

Set k = Coll
End Function


'Public Function KFT(Data As String) As Collection
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

Public Function KFT(Data) As Collection
Dim Coll As New Collection
Dim temp, temp2
    temp = SplitWithLengths(Data, 1, 1, 60, 1, 3, 9, 3, 2)
    Coll.Add temp(0), "ISSUE"
    Coll.Add temp(2), "TAXFLD"
    Coll.Add temp(3), "OTAXI"
    Coll.Add temp(4), "TCURR"
    Coll.Add temp(5), "TAXA"
    Coll.Add temp(6), "TAXC"
    Coll.Add temp(7), "TAXN"
Set KFT = Coll
End Function

Public Function TAX(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

splited = SplitForce(Data, ";", 3)

For j = 0 To UBound(splited)
    TAX_Sub splited(j), j, Coll
Next
Set TAX = Coll
End Function


Public Function TAX_Sub(Data, index, Coll As Collection) As Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

If Len(Data) = 24 Then
    temp = SplitWithLengths(Data, 3, 3, 9, 3)
    Coll.Add temp(0), "TAXFLD" & index
    Coll.Add temp(1), "TCURR" & index
    Coll.Add temp(2), "TAXA" & index
    Coll.Add temp(3), "TAXC" & index
Else
    temp = SplitWithLengths(Data, 3, 9, 3)
    Coll.Add "", "TAXFLD" & index
    Coll.Add temp(0), "TCURR" & index
    Coll.Add temp(1), "TAXA" & index
    Coll.Add temp(2), "TAXC" & index
End If
    
Set TAX_Sub = Coll
End Function




Public Function U___2(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

splited = SplitForce(Data, ";", 13)
    temp = SplitWithLengths(splited(0), 3, 1)
    Coll.Add temp(0), "TTONBR"
    Coll.Add temp(1), "DATAMODI"
    
    temp = SplitWithLengthsPlus(splited(1), 3, 1)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "STOP"
    Coll.Add temp(2), "ORIGAIR"
    
Coll.Add splited(2), "ORIGCTY"
Coll.Add splited(3), "DESTAIR"
Coll.Add splited(4), "DESTCTY"
Coll.Add splited(5), "HVOID"
    
Set U___2 = Coll
End Function



Public Function InsertData(Data As Collection, TableName As String, LineIdent As String, UploadNo As Long) As Boolean
On Error GoTo errPara
Dim Rs As New ADODB.Recordset
Dim j As Integer
Rs.Open "Select * from " + TableName, dbCompany, adOpenDynamic, adLockPessimistic
Rs.AddNew
Rs.Fields(0) = UploadNo
Rs.Fields(1) = LineIdent
For j = 2 To Rs.Fields.Count - 1
    Rs.Fields(j).value = Data(j - 1)
Next
Rs.Update
Exit Function
errPara:
'MsgBox Err.Description & vbCrLf & Rs.Fields(j).Name & " : (" & Rs.Fields(j).DefinedSize & ") at " & j & " Column in Table '" & TableName & "'", vbCritical, "Error During Updation"
End Function

Public Function InsertDataCollection(Data As Collection, TableName As String, LineIdent As String, UploadNo As Long) As Boolean
For j = 1 To Data.Count
    InsertData Data(j), TableName, LineIdent, UploadNo
Next
End Function


Public Function M_Sub(Data) As Collection
Dim Coll As New Collection
Dim splited
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

Public Function M(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
splited = Split(Data, ";")

For j = 0 To UBound(splited)
If Len(splited(j)) > 0 Then
    Coll.Add M_Sub(splited(j))
End If
Next

Set M = Coll
End Function

Public Function U(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
splited = SplitForce(Data, ";", 22)

    temp = SplitTwo(splited(0), 3)
    Coll.Add temp(0), "TTONBR"
    Coll.Add temp(1), "DATAMODI"

    temp = SplitWithLengthsPlus(splited(1), 3, 1)
    Coll.Add temp(0), "SEGNBR"
    Coll.Add temp(1), "STOP"
    Coll.Add temp(2), "ORIGAIR"
    
Coll.Add splited(2), "ORIGCTY"
Coll.Add splited(3), "DESTAIR"
Coll.Add splited(4), "DESTCTY"

    temp = SplitWithLengthsPlus(splited(5), 6, 5, 2, 2, 5, 5, 5, 5)
    Coll.Add temp(0), "AIRCDE"
    Coll.Add temp(1), "FLTNBR"
    Coll.Add temp(2), "CLSSVC"
    Coll.Add temp(3), "CLSBKG"
    Coll.Add temp(4), "DEPDTE"
    Coll.Add temp(5), "DEPTIM"
    Coll.Add temp(6), "ARRTIM"
    Coll.Add temp(7), "ARRDTE"

Coll.Add splited(6), "STATUS"
Coll.Add splited(7), "PNRSTAT"
Coll.Add splited(8), "MEAL"
Coll.Add splited(9), "NBRSTOP"
Coll.Add splited(10), "EQUIP"
Coll.Add splited(11), "ENTER"
Coll.Add splited(12), "SHRCOMM"
Coll.Add splited(13), "BAGALL"
Coll.Add splited(14), "CKINTRM"
Coll.Add splited(15), "CKINTIM"
Coll.Add splited(16), "TKTT"
Coll.Add splited(17), "FLTDUR"
Coll.Add splited(18), "NONSMOK"
Coll.Add splited(19), "TTAG"
Coll.Add splited(20), "MLTPM"
Coll.Add splited(21), "FLNTAG"


Set U = Coll
End Function



Public Function i(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2, temp3
Dim sTemp1, sTemp2, sTemp3
splited = SplitForce(Data, ";", 13)
Coll.Add splited(0), "GCPSGR"
    temp = SplitWithLengthsPlus(splited(1), 2)
    Coll.Add temp(0), "PSGRNBR"
    temp2 = SplitField(temp(1), "(", ")")
    temp3 = SplitForce(temp2(0), "/", 2)
    Coll.Add temp3(0), "PSGRFNME"
    Coll.Add temp3(1), "PSGRLNME"
    Coll.Add temp2(1), "PSGRID"
    temp3 = temp2(2)
    temp = SplitField(temp3, "(", ")")
    temp2 = SplitTwo(temp(1), 2)
    
    Coll.Add temp2(0), "IDFLD"
    Coll.Add temp2(1), "IDNBR"
    
Coll.Add splited(2), "CRRMK"
Coll.Add splited(3), "APINFO"
Coll.Add splited(4), "TRARL"
Coll.Add splited(5), "COMRL"

Set i = Coll
End Function



Public Function SSR(Data As String) As Collection

Dim Coll As New Collection
Dim splited
Dim temp, temp2, temp3, temp4
Dim sTemp1, sTemp2, sTemp3
splited = SplitForce(Data, ";S", 2)

    temp = SplitFirstTwo(splited(0), "/")
    temp2 = SplitWithLengthsPlus(temp(0), 1, 4, 1, 3, 1, 2)
    Coll.Add temp2(1), "SSRTYPE"
    Coll.Add temp2(3), "AIRCDE"
    Coll.Add temp2(4), "STATCDE"
    Coll.Add temp2(5), "NBRPTY"
    Coll.Add temp(1), "FREE"
    
    temp = SplitForce(splited(1), ";P", 2)
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


Public Function Q(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

temp = SplitForce(Data, ";", 2)
    Coll.Add temp(0), "FARE"
    Coll.Add temp(1), "PRCENTRY"
Set Q = Coll
End Function



Public Function T(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

splited = SplitFirstTwo(Data, "/")
    temp = SplitForce(splited(0), "-", 3)
    temp2 = SplitWithLengths(temp(0), 1, 3)
    Coll.Add temp2(0), "TKTTYPE"
    Coll.Add temp2(1), "NUMAIR"
    
    Coll.Add temp(1), "TKTNBR"
    Coll.Add temp(2), "DIGIT"
Coll.Add splited(1), "FREE"
    
Set T = Coll
End Function


'Public Function O(data As String) As Collection
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




Public Function FP(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

splited = SplitFirstTwo(Data, ";S")
    Coll.Add splited(0), "PAYMNT"
splited = SplitFirstTwo(splited(1), ";P")
    temp = SplitFirstTwo(splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"
temp = SplitFirstTwo(splited(1), "-")
temp2 = SplitTwoReverse(temp(0), 2)
    Coll.Add temp2(0), "PSGASS"
    Coll.Add temp2(1), "PSGNBR"
temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "PSGNBR2"
Set FP = Coll
End Function



Public Function FM(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

splited = SplitFirstTwo(Data, ";S")
    Coll.Add splited(0), "FARECOM"
splited = SplitFirstTwo(splited(1), ";P")
    temp = SplitFirstTwo(splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"

    temp = SplitFirstTwo(splited(1), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "PSGASS"
    Coll.Add temp2(1), "PSGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "PSGNBR2"
Set FM = Coll
End Function



Public Function AddFile(fileName As String)
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

    Set tsObj = fsObj.OpenTextFile(fileName, ForReading)
    UploadNo = mSeqNumberGen("UPLOADNO")
    Do While tsObj.AtEndOfStream <> True
        DoEvents
        LineStr = tsObj.ReadLine
head = getLineInfo(LineStr, id)

Select Case id
Case 0
    insertLineGeneral LineStr, head, UploadNo
    If AMDCompleted = False Then
        temp = SplitTwo(AMDLINE, 3)
        InsertAMD temp(1), UploadNo, GetFileNameFromPath(fileName)
        AMDCompleted = True
    End If
Case 1
    InserAirMain LineStr, UploadNo
Case -2
    AMDLINE = LineStr
    AMDCompleted = False
Case -1
    AMDLINE = AMDLINE & Chr(13) & Chr(10) & LineStr
End Select
        
        
        
    Loop
    tsObj.Close
    
    FileObj.CopyFile txtFileName.Text, txtRelLocation.Text & "\", True
On Error Resume Next
    FileObj.DeleteFile txtFileName.Text, True

End Function

Public Sub ClearTable(TableName As String)
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

Public Function InserAirMain(DataLine As String, UploadNo As Long)
Dim aa, Coll As Collection
aa = SplitTwo(DataLine, 7)
Set Coll = BLK(aa(1))
InsertData Coll, "AmdHeadBLK", "BLK", UploadNo
End Function

Public Function InsertAMD(Data, UploadNo As Long, Optional fileName As String = "")
Dim Coll As Collection
Set Coll = AMD(Data)
Coll.Add fileName, "FName"
InsertData Coll, "AmdHeadAMD", "AMD", UploadNo
End Function

Public Function TSA(Data As String) As Collection
Dim Coll As New Collection
Dim splited, MainSplit
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
MainSplit = SplitForce(Data, ";S", 3)
splited = SplitForce(MainSplit(0), ";", 3)

temp = SplitForce(splited(0), "+", 3)

    temp2 = SplitWithLengths(temp(0), 1, 3)
        Coll.Add temp2(0), "SALEINFO"
        Coll.Add temp2(1), "INTIND"
    Coll.Add temp(1), "SALEIND"
    Coll.Add temp(2), "ORGDEST"
    Coll.Add temp(3), "DOMISO"
    
    
temp = SplitForce(splited(1), "+", 3)
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

Public Function TK(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

splited = SplitFirstTwo(Data, ";S")
    Coll.Add splited(0), "TKTARR"
splited = SplitFirstTwo(splited(1), ";P")
    temp = SplitFirstTwo(splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"

    temp = SplitFirstTwo(splited(1), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "PSGASS"
    Coll.Add temp2(1), "PSGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "PSGNBR2"
Set TK = Coll
End Function

Public Function FV(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

splited = SplitFirstTwo(Data, ";S")
    Coll.Add splited(0), "TKTCARR"
splited = SplitFirstTwo(splited(1), ";P")
    temp = SplitFirstTwo(splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"

    temp = SplitFirstTwo(splited(1), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "PSGASS"
    Coll.Add temp2(1), "PSGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "PSGNBR2"
Set FV = Coll
End Function



Public Function RM(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

splited = SplitFirstTwo(Data, ";S")
    temp = SplitTwo(splited(0), 1)
    Coll.Add temp(0), "RMTYPE"
    Coll.Add temp(0), "PNRRMK"
splited = SplitFirstTwo(splited(1), ";P")
    temp = SplitFirstTwo(splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"

    temp = SplitFirstTwo(splited(1), "-")
    temp2 = SplitTwoReverse(temp(0), 2)
    Coll.Add temp2(0), "PSGASS"
    Coll.Add temp2(1), "PSGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "PSGNBR2"
Set RM = Coll
End Function


Public Function OSI(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3

splited = SplitFirstTwo(Data, ";S")
    temp = SplitWithLengthsPlus(splited(0), 1, 3, 1)
    Coll.Add temp(1), "AIRCDE"
    Coll.Add temp(3), "FFTEXT"
splited = SplitFirstTwo(splited(1), ";P")
    temp = SplitFirstTwo(splited(0), "-")
    temp2 = SplitTwoReverse(temp(0), 3)
    Coll.Add temp2(0), "SEGASS"
    Coll.Add temp2(1), "SEGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "SEGNBR2"

    temp = SplitFirstTwo(splited(1), "-")
    temp2 = SplitTwoReverse(temp(0), 2)
    Coll.Add temp2(0), "PSGASS"
    Coll.Add temp2(1), "PSGNBR"
    temp2 = SplitFirstTwo(temp(1), ",")
    Coll.Add temp2(0), "PSGNBR2"
Set OSI = Coll
End Function


Public Function N(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
    temp = SplitWithLengths(Data, 3, 28, 11)
    Coll.Add temp(0), "CURRTYP"
    Coll.Add temp(1), "FARESEG"
    Coll.Add temp(2), "BASE"
Set N = Coll
End Function
Public Function O(Data As String) As Collection
Dim Coll As New Collection
Dim splited
Dim temp, temp2
Dim sTemp1, sTemp2, sTemp3
    temp = SplitWithLengths(Data, 28, 10)
    Coll.Add temp(0), "VALELEM"
    Coll.Add temp(1), "VALDTE"
Set O = Coll
End Function

