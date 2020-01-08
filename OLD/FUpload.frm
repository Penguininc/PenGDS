VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
      MaxLength       =   80
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
      MaxLength       =   80
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

Private Sub chkGal_Click()
'If chkGal.Value = 0 Then
'    dbCompany.Execute "Update File set STATUS=0"
'Else
'    dbCompany.Execute "Update File set STATUS=1"
'End If
End Sub

Private Sub CmdDelete_Click()
On Error GoTo Note
    Dim Count As Integer
    File1.Path = txtDirName.Text
    File1.Pattern = "*.mir"
       
    If txtDirName.Text = "" Or File1.ListCount = 0 Then MsgBox "MIR Files Not Found", vbInformation: Exit Sub
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
On Error GoTo Note
    Dim Count As Integer
    File1.Path = txtDirName.Text
    File1.Pattern = "*.mir"
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
        cmdLoad_Click
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
End Sub
Private Sub cmdLoad_Click()
'On Error GoTo Note
    Dim FileSystem As FileSystemObject
    Dim CopyStatus As Boolean
    Dim LineItem As String
    Dim FileObj, LineDet
    If txtFileName = "" Then Exit Sub
    
    NoteError = False
    CopyStatus = False
    
    Set FileObj = CreateObject("Scripting.FileSystemObject")
    Set LineDet = FileObj.OpenTextFile(txtFileName.Text, ForReading, TristateUseDefault)
    Do While LineDet.AtEndOfLine <> True
        DoEvents
        LineItem = LineDet.ReadLine
        'If fnChecktheDatabase(LineItem) = True Then CopyStatus = True: Exit Do
        SelectDetails (LineItem)
    Loop
    LineDet.Close
    
    If CopyStatus = False And NoteError = False Then
        'FileObj.MoveFile txtFileName.Text, txtRelLocation.Text & "\"
        FileObj.CopyFile txtFileName.Text, txtRelLocation.Text & "\", True
On Error Resume Next
        FileObj.DeleteFile txtFileName.Text, True
    End If
    Exit Sub
Note:
    Me.Hide
End Sub

Public Sub SelectDetails(LineStr As String)
Dim Item As String
i = 1
    UploadNo = mSeqNumberGen("UPLOADNO")
    '''CnCompany.BeginTrans
Step1:
    If InStr(1, LineStr, Chr(13)) > 0 Then
        Item = Mid(LineStr, 1, InStr(1, LineStr, Chr(13)))
        LineStr = Mid(LineStr, InStr(1, LineStr, Chr(13)) + 1, Len(LineStr))
        fnLineDivision (Item)
        i = i + 1
        If NoteError = True Then Exit Sub '''CnCompany.RollbackTrans:
        GoTo Step1
    End If
   ''''''CnCompany.CommitTrans
    Pgr.value = 100
End Sub
    Public Sub fnLineDivision(LStr As String)
    
    If Asc(LStr) = 13 Then Exit Sub
    On Error GoTo NoteError
    
            'T5 Table
            If fnMidValue(LStr, 1, 2) = "T5" Then
                'Header1
                LastSec = "T5"
                SQL = "INSERT INTO GalHeader1 (UPLOADNO,T50BID,T50TRC,T50TRC1,T50MIR,T50SIZE,T50SEQ,T50DTE,T50TME," & _
                    "T50ISC,T50ISA,T50ISN,T50DFT,T50INP,T50OUT,FNAME) values (" & Val(UploadNo) & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 2)) & "','" & Trim(fnMidValue(LStr, 3, 2)) & "'," & _
                    "" & Val(fnMidValue(LStr, 5, 4)) & "," & Val(fnMidValue(LStr, 9, 2)) & "," & _
                    "" & Val(fnMidValue(LStr, 11, 5)) & "," & Val(fnMidValue(LStr, 16, 5)) & "," & _
                    "'" & Format(fnMidDate(LStr, 21, 7), "dd/mmm/yyyy") & "','" & Trim(fnMidValue(LStr, 28, 5)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 33, 2)) & "'," & Val(fnMidValue(LStr, 35, 3)) & "," & _
                    "'" & Trim(fnMidValue(LStr, 38, 24)) & "','" & Format(fnMidDate(LStr, 62, 7), "dd/mmm/yyyy") & "'," & _
                    "'" & Trim(fnMidValue(LStr, 69, 6)) & "','" & Trim(fnMidValue(LStr, 75, 6)) & "','" & FNAME & "')"
                dbCompany.Execute SQL
            
            'A02 Passenger Table
            ElseIf fnMidValue(LStr, 1, 3) = "A02" Then
                LastSec = "A02"
                SQL = "INSERT INTO GalPassenger (UPLOADNO,A02SEC,A02NMC,A02TRN,A02YIN,A02TKT,A02NBK," & _
                    "A02INV,A02PIC,A02FIN,A02EIN) values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 4, 33)) & "'," & Val(fnMidValue(LStr, 37, 11)) & "," & _
                    "" & Val(fnMidValue(LStr, 48, 1)) & "," & Val(fnMidValue(LStr, 49, 10)) & "," & _
                    "" & Val(fnMidValue(LStr, 59, 2)) & ",'" & Trim(fnMidValue(LStr, 61, 9)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 70, 6)) & "'," & Val(fnMidValue(LStr, 76, 2)) & "," & _
                    "" & Val(fnMidValue(LStr, 78, 2)) & ")"
                dbCompany.Execute SQL
                
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
                SQL = "INSERT INTO GalA04 (UPLOADNO,A04SEC,A04ITN,A04CDE,A04NUM,A04NME,A04FLT,A04CLS," & _
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
                SQL = "INSERT INTO GalA07 (UPLOADNO,A07SEC,A07FSI,A07CRB,A07TBF,A07CRT,A07TTA,A07CRE," & _
                    "A07EQV,A07CUR,A07TI1,A07TT1,A07TC1,A07TI2,A07TT2,A07TC2,A07TI3,A07TT3,A07TC3," & _
                    "A07TI4,A07TT4,A07TC4,A07TI5,A07TT5,A07TC5,A07XFI,A07XFT) values (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & Val(fnMidValue(LStr, 4, 2)) & "," & _
                    "'" & Trim(fnMidValue(LStr, 6, 3)) & "','" & Trim(fnMidValue(LStr, 9, 12)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 21, 3)) & "','" & Trim(fnMidValue(LStr, 24, 12)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 36, 3)) & "','" & Trim(fnMidValue(LStr, 39, 12)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 51, 3)) & "','" & Trim(fnMidValue(LStr, 54, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 57, 8)) & "','" & Trim(fnMidValue(LStr, 65, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 67, 3)) & "','" & Trim(fnMidValue(LStr, 70, 8)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 78, 2)) & "','" & Trim(fnMidValue(LStr, 80, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 83, 8)) & "','" & Trim(fnMidValue(LStr, 91, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 93, 3)) & "','" & Trim(fnMidValue(LStr, 96, 8)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 104, 2)) & "','" & Trim(fnMidValue(LStr, 106, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 109, 8)) & "','" & Trim(fnMidValue(LStr, 117, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 119, 3)) & "','" & Trim(fnMidValue(LStr, 122, 16)) & "')"
                dbCompany.Execute SQL
            
            'A08 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A08" Then
                LastSec = "A08"
                SQL = "INSERT INTO GalA08 (UPLOADNO,A08SEC,A08FSI,A08ITN,A08FBC,A08VAL,A08NVBC," & _
                    "A08NVAC,A08TDGC,A08ENDI,A08END) values (" & UploadNo & ",'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & _
                    "" & Val(fnMidValue(LStr, 4, 2)) & "," & Val(fnMidValue(LStr, 6, 2)) & "," & _
                    "'" & Trim(fnMidValue(LStr, 8, 8)) & "','" & Trim(fnMidValue(LStr, 16, 8)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 24, 7)) & "','" & Trim(fnMidValue(LStr, 31, 7)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 38, 6)) & "','" & Trim(fnMidValue(LStr, 44, 2)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 46, 60)) & "')"
                    
                dbCompany.Execute SQL
                
            'A09 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A09" Then
                LastSec = "A09"
                SQL = "INSERT INTO GalA09 (UPLOADNO,A09SEC,A09FSI,A09TY5,A09L51) Values (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "'," & Val(fnMidValue(LStr, 4, 2)) & "," & _
                    "" & Val(fnMidValue(LStr, 6, 1)) & ",'" & Trim(fnMidValue(LStr, 7, 61)) & "')"
                
                dbCompany.Execute SQL
                
                
            'A12 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A12" Then
                LastSec = "A12"
                SQL = "INSERT INTO GalA12 (UPLOADNO,A12SEC,A12CTY,A12LOC,A12PHA) VALUES (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 3)) & "'," & _
                    "'" & Trim(fnMidValue(LStr, 7, 2)) & "','" & Trim(fnMidValue(LStr, 9, 61)) & "')"
                dbCompany.Execute SQL
            
            'A14 Table
            ElseIf fnMidValue(LStr, 1, 3) = "A14" Then
                LastSec = "A14"
                SQL = "INSERT INTO GalA14 (UPLOADNO,A14SEC,A14RMK) values (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 64)) & "')"
                dbCompany.Execute SQL
            
            'A15 Table by Abhi
            ElseIf fnMidValue(LStr, 1, 3) = "A15" Then
                LastSec = "A15"
                SQL = "INSERT INTO GalA15 (UPLOADNO,A15SEC,A15SEG,A15RMK) values (" & UploadNo & "," & _
                    "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 2)) & "','" & Trim(fnMidValue(LStr, 6, 87)) & "')"
                dbCompany.Execute SQL
            
            'A16 Table
'            ElseIf fnMidValue(LStr, 1, 3) = "A16" Then
'                LastSec = "A16"
'                sql = "INSERT INTO GalA16 (UPLOADNO,A16SEC,A16TYP,A16NUM,A16DTE,A16PRP,A16HCC,A16CTY,A16MUL,A16STT,A16OUT,A16DAY," & _
'                      "A16NME,A16FON,A16FAX,A16RTT,A16RMS,A16LOC) values (" & UploadNo & "," & _
'                      "'" & Trim(fnMidValue(LStr, 1, 3)) & "','" & Trim(fnMidValue(LStr, 4, 1)) & "','" & Trim(fnMidValue(LStr, 5, 2)) & "'," & _
'                      "'" & Trim(fnMidValue(LStr, 7, 7)) & "','" & Trim(fnMidValue(LStr, 14, 6)) & "','" & Trim(fnMidValue(LStr, 20, 2)) & "'," & _
'                      "'" & Trim(fnMidValue(LStr, 22, 4)) & "','" & Trim(fnMidValue(LStr, 26, 1)) & "','" & Trim(fnMidValue(LStr, 27, 2)) & "'," & _
'                      "'" & Trim(fnMidValue(LStr, 29, 5)) & "','" & Trim(fnMidValue(LStr, 34, 3)) & "','" & Trim(fnMidValue(LStr, 37, 20)) & "'," & _
'                      "'" & Trim(fnMidValue(LStr, 57, 17)) & "','" & Trim(fnMidValue(LStr, 74, 17)) & "','" & Trim(fnMidValue(LStr, 91, 1)) & "'," & _
'                      "'" & Trim(fnMidValue(LStr, 92, 8)) & "','" & Trim(fnMidValue(LStr, 100, 20)) & "')"
'                CnCompany.Execute sql
            
            Else
                If LastSec = "T5" Then
                    If i = 2 Then
                        'Header2
                        SQL = "INSERT INTO GalHeader2 (UPLOADNO,T50BPC,T50TPC,T50AAN,T50RCL,T50ORL,T50OCC,T50OAM,T50AGS," & _
                            "T50SBI,T50SIN,T50DTY,T50PNR,T50EHT,T50DTL,T50NMC) Values (" & Val(UploadNo) & "," & _
                            "'" & Trim(fnMidValue(LStr, 1, 4)) & "','" & Trim(fnMidValue(LStr, 5, 4)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 9, 9)) & "','" & Trim(fnMidValue(LStr, 18, 6)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 24, 6)) & "','" & Trim(fnMidValue(LStr, 30, 2)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 32, 1)) & "','" & Trim(fnMidValue(LStr, 33, 6)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 39, 1)) & "','" & Trim(fnMidValue(LStr, 40, 2)) & "'," & _
                            "'" & Trim(fnMidValue(LStr, 42, 2)) & "','" & Format(fnMidDate(LStr, 44, 7), "dd/mmm/yyyy") & "'," & _
                            "" & Val(fnMidValue(LStr, 51, 3)) & ",'" & Format(fnMidDate(LStr, 54, 7), "DD/MMM/YYYY") & "'," & _
                            "" & Val(fnMidValue(LStr, 61, 3)) & ")"
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
                            "" & Val(fnMidValue(LStr, 70, 8)) & "," & Val(fnMidValue(LStr, 78, 4)) & "," & _
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
                
                    SQL = "INSERT INTO GalPassenger1 (UPLOADNO,A02PNI,A02PNR,A02PT1,A02PTL,A02PA1,A02PA2," & _
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
                    
                End If
            End If
    Exit Sub
NoteError:
        Resume Next
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
On Error GoTo Note
    Dim rsSelect As New ADODB.Recordset
    Me.Icon = FMain.Icon
    rsSelect.Open "Select UPLOADDIRNAME,DESTDIRNAME From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    If rsSelect.EOF = False Then
        txtDirName = IIf(IsNull(rsSelect!UPLOADDIRNAME), "", rsSelect!UPLOADDIRNAME)
        txtRelLocation = IIf(IsNull(rsSelect!DESTDIRNAME), "", rsSelect!DESTDIRNAME)
    End If
    rsSelect.Close
    If rsSelect.State = 1 Then rsSelect.Close
    rsSelect.Open "Select STATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    chkGal.value = IIf((IsNull(rsSelect!Status) = True), 0, rsSelect!Status)
    File1.Path = txtDirName.Text
    File1.Refresh
    File1.Pattern = "*.mir"
    File1.Refresh
    Exit Sub
Note:
    If Me.Visible = True Then Unload Me
End Sub
Public Function fnChecktheDatabase(CheckStr As String) As Boolean
Dim rsSelect As New ADODB.Recordset
Dim RecLoc As String

RecLoc = fnMidValue(CheckStr, 99, 6)
rsSelect.Open "Select * From GalHeader2 Where T50RCL='" & RecLoc & "'", dbCompany, adOpenDynamic, adLockBatchOptimistic
If rsSelect.EOF = False Then
    fnChecktheDatabase = True
Else
    fnChecktheDatabase = False
End If
rsSelect.Close
End Function
