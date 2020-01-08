VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FSabreUpload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sabre Uploading..."
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSabre 
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
      Left            =   3780
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
      Top             =   2070
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   2070
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
      Top             =   375
      Width           =   1440
   End
End
Attribute VB_Name = "FSabreUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UploadNo As Long
Dim FNAME As String
Dim avpItems
Dim vlpLen As Long
Dim ai As Long
Dim avpItemsTemp As String
Dim vlTI As Long, vlTCount As Long

Private Sub chkSabre_Click()
'If chkSabre.Value = 0 Then
'    dbCompany.Execute "Update File set SABRESTATUS=0"
'Else
'    dbCompany.Execute "Update File set SABRESTATUS=1"
'End If
End Sub

Public Sub cmdDirUpload_Click()
On Error GoTo Note
    Dim Count As Integer
    File1.Path = txtDirName.Text
    File1.Pattern = "*.fil"
    If txtDirName.Text = "" Or File1.ListCount = 0 Then stbUpload.Panels(1).Text = "FIL Files Not Found....": Exit Sub
    'Uploading Each File
    Open (App.Path & "\_UploadingSQL_") For Random As #1
    Close #1
    FMain.cmdStop.Enabled = False
    DoEvents
    FMain.lblFName.Caption = "Uploading Sabre..."
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
    'FMain.txtSource = ""
    'FMain.txtDest = ""
    FMain.lblFName.Caption = ""
    fsObj.DeleteFile (App.Path & "\_UploadingSQL_"), True
    Me.Hide
    Exit Sub
Note:
    Me.Hide
End Sub
Private Sub cmdLoad_Click()
'On Error GoTo Note
    Dim tsObj As TextStream
    Dim LineStr As String
    Dim lnID As String
    Dim CopyStatus As Boolean
    
    CopyStatus = False
    
    Set tsObj = fsObj.OpenTextFile(txtFileName, ForReading)
    UploadNo = mSeqNumberGen("UPLOADNO")
    Do While tsObj.AtEndOfStream <> True
        DoEvents
        LineStr = tsObj.ReadLine
        'If fnChecktheDatabase(LineStr) = True Then CopyStatus = True: Exit Do
        lnID = Left(LineStr, 2)
        SelectDetails LineStr, lnID
    Loop
    tsObj.Close
    If CopyStatus = False Then
        'fsObj.MoveFile txtFileName.Text, txtRelLocation.Text & "\"
        fsObj.CopyFile txtFileName.Text, txtRelLocation.Text & "\", True
On Error Resume Next
        fsObj.DeleteFile txtFileName.Text, True
    End If
    Exit Sub
Note:
   Me.Hide
End Sub
Public Sub SelectDetails(ln As String, id As String)

Dim Query As String
Dim lnTemp As String

Static i As Integer
Static j As Integer
Static k As Integer


Static M3_flag As Boolean
Static M2_flag As Boolean
Static M1_flag As Boolean

Static M3_UploadNo As Long
Static M2_UploadNo As Long
Static M1_UploadNo As Long

    If id = "AA" And Mid(ln, 3, 1) <> "/" Then
        Query = "Insert into SabreHeader(UPLOADNO,Header1,Header2,Header3,Header4) " & _
                " Values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 1, 2)) & "','" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "', " & _
                " '" & Trim(fnMidValue(ln, 8, 4)) & "')"
        dbCompany.Execute Query
        If Trim(fnMidValue(ln, 14, 1)) <> 5 Then
            Query = "Insert into SabreM0(UPLOADNO,IU0TYP,IU0VER,IU0DKN,IU0CTJ,IU0CAP,IU0PIV,IU0RRC,IU0QUE,IU0CDK,IU0IVN,IU0ATC,IU0PNR,IU0PNL,IU0OPT,IU0F00, " & _
                    " IU0F01,IU0F02,IU0F03,IU0F04,IU0F05,IU0F06,IU0F07,IU0F08,IU0F09,IU0F0A,IU0F0B,IU0F0C,IU0F0D,IU0F0E,IU0F0F,IU0ATB,IU0DUP,IU0PCC,IU0IDC,IU0IAG, " & _
                    " IU0LIN,IU0RPR,IU0PDT,IU0TIM,IU0IS4,IU0IS1,IU0IS3,IU0TCO,IU0IDB,IU0DEP,IU0ORG,IU0ONM,IU0DST,IU0DNM,IU0NM1,IU0NM2,IU0NM3,IU0NM4,IU0NM5,IU0NM6, " & _
                    " IU0NM7,IU0NM8,IU0NM9,IU0NMA,IU0ADC,IU0PHC,IU0TYM,IU0MDT,FNAME) " & _
                    " Values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 14, 1)) & "','" & Trim(fnMidValue(ln, 15, 2)) & "','" & Trim(fnMidValue(ln, 17, 10)) & "','" & Trim(fnMidValue(ln, 27, 4)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 31, 1)) & "','" & Trim(fnMidValue(ln, 32, 1)) & "','" & Trim(fnMidValue(ln, 33, 2)) & "','" & Trim(fnMidValue(ln, 35, 1)) & "','" & Trim(fnMidValue(ln, 36, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 37, 7)) & "','" & Trim(fnMidValue(ln, 44, 10)) & "','" & Trim(fnMidValue(ln, 54, 8)) & "','" & Trim(fnMidValue(ln, 62, 8)) & "','" & Trim(fnMidValue(ln, 70, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 71, 1)) & "','" & Trim(fnMidValue(ln, 72, 1)) & "','" & Trim(fnMidValue(ln, 73, 1)) & "','" & Trim(fnMidValue(ln, 74, 1)) & "','" & Trim(fnMidValue(ln, 75, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 76, 1)) & "','" & Trim(fnMidValue(ln, 77, 1)) & "','" & Trim(fnMidValue(ln, 78, 1)) & "','" & Trim(fnMidValue(ln, 79, 1)) & "','" & Trim(fnMidValue(ln, 80, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 81, 1)) & "','" & Trim(fnMidValue(ln, 82, 1)) & "','" & Trim(fnMidValue(ln, 83, 1)) & "','" & Trim(fnMidValue(ln, 84, 1)) & "','" & Trim(fnMidValue(ln, 85, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 86, 1)) & "','" & Trim(fnMidValue(ln, 87, 1)) & "','" & Trim(fnMidValue(ln, 88, 1)) & "','" & Trim(fnMidValue(ln, 89, 5)) & "','" & Trim(fnMidValue(ln, 94, 2)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 96, 3)) & "','" & Trim(fnMidValue(ln, 99, 8)) & "','" & Trim(fnMidValue(ln, 107, 10)) & "','" & Trim(fnMidValue(ln, 117, 5)) & "','" & Trim(fnMidValue(ln, 122, 5)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 127, 5)) & "','" & Trim(fnMidValue(ln, 132, 2)) & "','" & Trim(fnMidValue(ln, 134, 3)) & "','" & Trim(fnMidValue(ln, 137, 2)) & "','" & Trim(fnMidValue(ln, 139, 3)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 142, 5)) & "','" & Trim(fnMidValue(ln, 147, 3)) & "','" & Trim(fnMidValue(ln, 150, 17)) & "','" & Trim(fnMidValue(ln, 167, 3)) & "','" & Trim(fnMidValue(ln, 170, 17)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 187, 3)) & "','" & Trim(fnMidValue(ln, 190, 3)) & "','" & Trim(fnMidValue(ln, 193, 3)) & "','" & Trim(fnMidValue(ln, 196, 3)) & "','" & Trim(fnMidValue(ln, 199, 3)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 202, 3)) & "','" & Trim(fnMidValue(ln, 205, 3)) & "','" & Trim(fnMidValue(ln, 208, 3)) & "','" & Trim(fnMidValue(ln, 211, 3)) & "','" & Trim(fnMidValue(ln, 214, 3)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 217, 2)) & "','" & Trim(fnMidValue(ln, 219, 2)) & "','" & Trim(fnMidValue(ln, 221, 5)) & "','" & Trim(fnMidValue(ln, 226, 5)) & "','" & FNAME & "')"
        Else
            Query = "Insert into SabreM0(UPLOADNO,IU0TYP,IU0VER,IU0DKN,IU0TKI,IU0CNT,IU0PR1,IU0DUP,IU0SAC,FNAME) " & _
                    " " & UploadNo & ",'" & Trim(fnMidValue(ln, 14, 1)) & "','" & Trim(fnMidValue(ln, 15, 2)) & "','" & Trim(fnMidValue(ln, 17, 10)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 27, 14)) & "','" & Trim(fnMidValue(ln, 41, 2)) & "','" & Trim(fnMidValue(ln, 43, 8)) & "','" & Trim(fnMidValue(ln, 51, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 52, 15)) & "','" & FNAME & "')"
        End If
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M1" And Mid(ln, 3, 1) <> "/" Then
        M1_flag = True
        M1_UploadNo = UploadNo
        j = 0
        Query = "Insert into SabreM1(UPLOADNO,IU1PNO,IU1PNM,IU1PRK,IU1IND,IU1AVN,IU1MLV,IU1NM3,IU1TKT,IU1SSM,IU1NM5,IU1NM7, " & _
                " IU1NM8,IU1NM9,IU1NMA) " & _
                " Values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & ",'" & Trim(fnMidValue(ln, 5, 64)) & "','" & Trim(fnMidValue(ln, 69, 30)) & "', " & _
                " '" & Trim(fnMidValue(ln, 99, 1)) & "','" & Trim(fnMidValue(ln, 100, 20)) & "','" & Trim(fnMidValue(ln, 120, 5)) & "', " & _
                " '" & Trim(fnMidValue(ln, 125, 2)) & "','" & Trim(fnMidValue(ln, 127, 1)) & "','" & Trim(fnMidValue(ln, 128, 1)) & "'," & Val(fnMidValue(ln, 129, 2)) & ", " & _
                " '" & Trim(fnMidValue(ln, 131, 2)) & "','" & Trim(fnMidValue(ln, 133, 2)) & "','" & Trim(fnMidValue(ln, 135, 2)) & "','" & Trim(fnMidValue(ln, 137, 2)) & "')"
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M2" And Mid(ln, 3, 1) <> "/" Then
        M2_flag = True
        M2_UploadNo = UploadNo
        i = 0
        Query = "Insert into SabreM2(UPLOADNO,IU2PNO,IU2PTY,IU2TCN,IU2INT,IU2SSI,IU2IND,IU2AP1,IU2AP2,IU2AP3,IU2AP4,IU2AP5,IU2AP6,IU2AP7,IU2AP8,IU2AP9,IU2APA,IU2APB,IU2APZ,IU2FSN,IU2FCC,IU2FAR,IU2T1S,IU2TX1,IU2ID1,IU2T2S, " & _
            " IU2TX2,IU2ID2,IU2T3S,IU2TX3,IU2ID3,IU2TFS,IU2TFC,IU2TFR,IU2PEN,IU2KXP,IU2SPX,IU2EPS,IU2EFC,IU2EFR,IU2PCT,IU2CSN,IU2COM,IU2NAS,IU2NET,IU2CDC,IU2CTT,IU2FTF,IU2TAT,IU2TUR,IU2TPC,IU2MIT,IU2CCB,IU2APC, " & _
            " IU2SN4,IU2SN3,IU2SN2,IU2PN4,IU2PN3,IU2PN2,IU2AVI,IU2ATH,IU2SPR,IU2EXP,IU2EXT,IU2M4C,IU2M6C,IU2VAL,IU2TNO,IU2DNO)" & _
            " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & ",'" & Trim(fnMidValue(ln, 5, 3)) & "','" & Trim(fnMidValue(ln, 8, 11)) & "','" & Trim(fnMidValue(ln, 19, 1)) & "','" & Trim(fnMidValue(ln, 20, 1)) & "','" & Trim(fnMidValue(ln, 21, 1)) & "', " & _
            " '" & Trim(fnMidValue(ln, 22, 1)) & "','" & Trim(fnMidValue(ln, 23, 1)) & "','" & Trim(fnMidValue(ln, 24, 1)) & "','" & Trim(fnMidValue(ln, 25, 1)) & "','" & Trim(fnMidValue(ln, 26, 1)) & "','" & Trim(fnMidValue(ln, 27, 1)) & "', " & _
            " '" & Trim(fnMidValue(ln, 28, 1)) & "','" & Trim(fnMidValue(ln, 29, 1)) & "','" & Trim(fnMidValue(ln, 30, 1)) & "','" & Trim(fnMidValue(ln, 31, 1)) & "','" & Trim(fnMidValue(ln, 32, 1)) & "','" & Trim(fnMidValue(ln, 33, 1)) & "', " & _
            " '" & Trim(fnMidValue(ln, 34, 1)) & "','" & Trim(fnMidValue(ln, 35, 3)) & "'," & Val(Trim(fnMidValue(ln, 38, 8))) & ",'" & Trim(fnMidValue(ln, 46, 1)) & "'," & Val(Trim(fnMidValue(ln, 47, 7))) & ",'" & Trim(fnMidValue(ln, 54, 2)) & "', " & _
            " '" & Trim(fnMidValue(ln, 56, 1)) & "'," & Val(Trim(fnMidValue(ln, 57, 7))) & ",'" & Trim(fnMidValue(ln, 64, 2)) & "','" & Trim(fnMidValue(ln, 66, 1)) & "'," & Val(Trim(fnMidValue(ln, 67, 7))) & ",'" & Trim(fnMidValue(ln, 74, 2)) & "', " & _
            " '" & Trim(fnMidValue(ln, 76, 1)) & "','" & Trim(fnMidValue(ln, 77, 3)) & "'," & Val(Trim(fnMidValue(ln, 80, 8))) & "," & Val(Trim(fnMidValue(ln, 88, 11))) & "," & Val(Trim(fnMidValue(ln, 99, 11))) & ",'" & Trim(fnMidValue(ln, 110, 6)) & "', " & _
            " '" & Trim(fnMidValue(ln, 116, 1)) & "','" & Trim(fnMidValue(ln, 117, 3)) & "'," & Val(Trim(fnMidValue(ln, 120, 8))) & ",'" & Trim(fnMidValue(ln, 128, 8)) & "','" & Trim(fnMidValue(ln, 136, 1)) & "'," & Val(Trim(fnMidValue(ln, 137, 8))) & ", " & _
            " '" & Trim(fnMidValue(ln, 145, 1)) & "'," & Val(Trim(fnMidValue(ln, 146, 8))) & ",'" & Trim(fnMidValue(ln, 154, 1)) & "'," & Val(Trim(fnMidValue(ln, 155, 1))) & ",'" & Trim(fnMidValue(ln, 156, 1)) & "','" & Trim(fnMidValue(ln, 157, 8)) & "', " & _
            " '" & Trim(fnMidValue(ln, 165, 15)) & "','" & Trim(fnMidValue(ln, 180, 4)) & "','" & Trim(fnMidValue(ln, 184, 1)) & "','" & Trim(fnMidValue(ln, 185, 1)) & "','" & Trim(fnMidValue(ln, 186, 1)) & "','" & Trim(fnMidValue(ln, 187, 5)) & "', " & _
            " '" & Trim(fnMidValue(ln, 192, 2)) & "','" & Trim(fnMidValue(ln, 194, 3)) & "','" & Trim(fnMidValue(ln, 197, 5)) & "','" & Trim(fnMidValue(ln, 202, 2)) & "','" & Trim(fnMidValue(ln, 204, 3)) & "','" & Trim(fnMidValue(ln, 207, 1)) & "', " & _
            " '" & Trim(fnMidValue(ln, 208, 9)) & "','" & Trim(fnMidValue(ln, 217, 5)) & "','" & Trim(fnMidValue(ln, 222, 4)) & "','" & Trim(fnMidValue(ln, 226, 2)) & "','" & Trim(fnMidValue(ln, 228, 2)) & "','" & Trim(fnMidValue(ln, 230, 2)) & "', " & _
            " '" & Trim(fnMidValue(ln, 232, 2)) & "','" & Trim(fnMidValue(ln, 234, 10)) & "','" & Trim(fnMidValue(ln, 244, 1)) & "')"
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M3" And Mid(ln, 3, 1) <> "/" Then
        M3_flag = True
        M3_UploadNo = UploadNo
        k = 0
        If Trim(fnMidValue(ln, 5, 1)) = 1 Then
            Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3BPI,IU3DCC,IU3DCY,IU3ACC,IU3ACY,IU3CAR,IU3FLT,IU3CLS,IU3DTM,IU3ATM,IU3ELT,IU3MLI,IU3SUP,IU3DCH,IU3NOS,IU3SCC,IU3CRT,IU3EQP,IU3ARM,IU3AVM,IU3SCT,IU3RTM,IU3GAT,IU3TMA,IU3GTA,IU3GAR,IU3RTM_1,IU3COG,IU3CRN,IU3TKT,IU3MCT,IU3SPX) " & _
                    " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & "," & Val(Trim(fnMidValue(ln, 5, 1))) & "," & Val(Trim(fnMidValue(ln, 6, 1))) & "," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "','" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 19, 3)) & "','" & Trim(fnMidValue(ln, 22, 17)) & "','" & Trim(fnMidValue(ln, 39, 3)) & "','" & Trim(fnMidValue(ln, 42, 17)) & "','" & Trim(fnMidValue(ln, 59, 2)) & "','" & Trim(fnMidValue(ln, 61, 5)) & "','" & Trim(fnMidValue(ln, 66, 2)) & "','" & Trim(fnMidValue(ln, 68, 5)) & "','" & Trim(fnMidValue(ln, 73, 5)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 78, 8)) & "','" & Trim(fnMidValue(ln, 86, 4)) & "','" & Trim(fnMidValue(ln, 90, 1)) & "','" & Trim(fnMidValue(ln, 91, 1)) & "','" & Trim(fnMidValue(ln, 92, 1)) & "','" & Trim(fnMidValue(ln, 93, 18)) & "','" & Trim(fnMidValue(ln, 111, 2)) & "','" & Trim(fnMidValue(ln, 113, 3)) & "','" & Trim(fnMidValue(ln, 116, 6)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 122, 6)) & "','" & Trim(fnMidValue(ln, 128, 2)) & "','" & Trim(fnMidValue(ln, 130, 26)) & "','" & Trim(fnMidValue(ln, 156, 4)) & "','" & Trim(fnMidValue(ln, 160, 26)) & "','" & Trim(fnMidValue(ln, 186, 5)) & "','" & Trim(fnMidValue(ln, 191, 4)) & "','" & Trim(fnMidValue(ln, 195, 5)) & "','" & Trim(fnMidValue(ln, 200, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 201, 37)) & "','" & Trim(fnMidValue(ln, 238, 1)) & "','" & Trim(fnMidValue(ln, 239, 2)) & "','" & Trim(fnMidValue(ln, 241, 12)) & "')"
        ElseIf Trim(fnMidValue(ln, 5, 1)) = 3 Then
            
            If fnMidValue(ln, 15, 3) = "HHL" Then
            
                Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3NOS,IU3SCC,IU3CCN,IU3PTY,IU3CFN,IU3CAC,IU3TDG,IU3OUT,IU3PRP,IU3ITT,IU3NRM,IU3VR2) " & _
                          " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & "," & Val(Trim(fnMidValue(ln, 5, 1))) & "," & Val(Trim(fnMidValue(ln, 6, 1))) & "," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "','','','" & Trim(fnMidValue(ln, 19, 2)) & "','" & Trim(fnMidValue(ln, 21, 15)) & "','" & Trim(fnMidValue(ln, 36, 3)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 39, 2)) & "','" & Trim(fnMidValue(ln, 41, 14)) & "','" & Trim(fnMidValue(ln, 55, 6)) & "','" & Trim(fnMidValue(ln, 61, 32)) & "','" & Trim(fnMidValue(ln, 93, 11)) & "','" & Trim(fnMidValue(ln, 104, Len(ln))) & "')"
            Else
                
               Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3CCN,IU3PTY,IU3CFN,IU3VR4) " & _
                          " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & "," & Val(Trim(fnMidValue(ln, 5, 1))) & "," & Val(Trim(fnMidValue(ln, 6, 1))) & "," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "','" & Trim(fnMidValue(ln, 19, 2)) & "','" & Trim(fnMidValue(ln, 21, 15)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 36, Len(ln))) & "')"
            End If
        
        Else
            Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3CCN,IU3PTY,IU3CFN,IU3VR4) " & _
                    " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & "," & Val(Trim(fnMidValue(ln, 5, 1))) & "," & Val(Trim(fnMidValue(ln, 6, 1))) & "," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "','" & Trim(fnMidValue(ln, 19, 2)) & "','" & Trim(fnMidValue(ln, 21, 15)) & "','" & Trim(fnMidValue(ln, 36, Len(ln))) & "')"
        End If
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M4" And Mid(ln, 3, 1) <> "/" Then
        Query = "Insert into SabreM4(UPLOADNO,IU4SEG,IU4TYP,IU4CNI,IU4ETP,IU4NVB,IU4NVA,IU4AAC,IU4AWL,IU4FBS,IU4TDS,IU4ACL,IU4AMT,IU4ETK,IU4FB2,IU4TD2,IU4CUR,IU4SP2) " & _
                " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "','" & Trim(fnMidValue(ln, 8, 1)) & "', " & _
                " '" & Trim(fnMidValue(ln, 9, 1)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "','" & Trim(fnMidValue(ln, 15, 5)) & "','" & Trim(fnMidValue(ln, 20, 2)) & "', " & _
                " '" & Trim(fnMidValue(ln, 22, 3)) & "','" & Trim(fnMidValue(ln, 25, 7)) & "','" & Trim(fnMidValue(ln, 32, 6)) & "','" & Trim(fnMidValue(ln, 38, 2)) & "', " & _
                " '" & Trim(fnMidValue(ln, 40, 8)) & "','" & Trim(fnMidValue(ln, 48, 1)) & "','" & Trim(fnMidValue(ln, 49, 12)) & "','" & Trim(fnMidValue(ln, 61, 10)) & "', " & _
                " '" & Trim(fnMidValue(ln, 71, 3)) & "','" & Trim(fnMidValue(ln, 74, 13)) & "')"
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M5" And Mid(ln, 3, 1) <> "/" Then
        lnTemp = Replace(ln, "#", "/")
        lnTemp = Replace(lnTemp, "@", "/")
        avpItems = Split(lnTemp, "/")
        ReDim Preserve avpItems(8)
'        Query = "Insert into SabreM5(UPLOADNO,IU5PTY,IU5MIN,IU5VR1,IU5VR2,IU5VR3,IU5VR4,IU5VR5,IU5VR6,IU5VR7,IU5VR8) " & _
                " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 2)) & "','" & Trim(fnMidValue(ln, 7, 1)) & "', " & _
                " '" & Trim(fnMidValue(ln, 8, 1)) & "','" & Trim(fnMidValue(ln, 9, 2)) & "','" & Trim(fnMidValue(ln, 12, 10)) & "', " & _
                " '" & Trim(fnMidValue(ln, 23, 7)) & "','" & Trim(fnMidValue(ln, 31, 8)) & "','" & Trim(fnMidValue(ln, 41, 6)) & "', " & _
                " '" & Trim(fnMidValue(ln, 48, Len(ln))) & "')"
        vlTCount = Val(avpItems(7))
        For vlTI = 1 To vlTCount
            Query = "Insert into SabreM5(UPLOADNO,IU5PTY,IU5MIN,IU5VR1,IU5VR2,IU5VR3,IU5VR4,IU5VR5,IU5VR6,IU5VR7,IU5VR8) " & _
                    " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 2)) & "','" & Trim(fnMidValue(ln, 7, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 8, 1)) & "','" & Trim(fnMidValue(ln, 9, 2)) & "','" & Trim(avpItems(1)) & "', " & _
                    " '" & Trim(avpItems(2)) & "','" & Trim(avpItems(3)) & "','" & Trim(avpItems(4)) & "', " & _
                    " '" & Trim(avpItems(5)) & "/" & Trim(avpItems(6)) & "/" & Trim(avpItems(7)) & "/" & Trim(avpItems(8)) & "')"
                
                dbCompany.Execute Query
            avpItems(1) = Val(avpItems(1)) + 1
            avpItems(2) = 0
            avpItems(3) = 0
            avpItems(4) = 0
            avpItems(7) = 0
        Next vlTI
        Exit Sub
    ElseIf id = "M6" And Mid(ln, 3, 1) <> "/" Then
        Query = "Insert into SabreM6(UPLOADNO,IU6RTY,IU6FCT,IU6CPR,IU6FC12) " & _
                " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 3)) & "','" & Trim(fnMidValue(ln, 6, 1)) & "', " & _
                " '" & Trim(fnMidValue(ln, 7, 1)) & "','" & Trim(fnMidValue(ln, 8, Len(ln))) & "')"
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M7" And Mid(ln, 3, 1) <> "/" Then
        Query = "Insert into SabreM7(UPLOADNO,IU7RKN,IU7RMK) " & _
                " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, Len(ln))) & "')"
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M9" And Mid(ln, 3, 1) <> "/" Then
        If InStr(UCase(ln), "PEN") > 0 Then
            avpItems = Split(ln, "/")
            ReDim Preserve avpItems(5)
            For ai = 1 To 5
                If UCase(Left(avpItems(ai), 2)) = "AC" And ai <> 1 Then
                    avpItemsTemp = avpItems(1)
                    avpItems(1) = avpItems(ai)
                    avpItems(ai) = avpItemsTemp
                End If
            Next ai
            For ai = 1 To 5
                If UCase(Left(avpItems(ai), 2)) = "DE" And ai <> 2 Then
                    avpItemsTemp = avpItems(2)
                    avpItems(2) = avpItems(ai)
                    avpItems(ai) = avpItemsTemp
                End If
            Next ai
            For ai = 1 To 5
                If UCase(Left(avpItems(ai), 2)) = "PO" And ai <> 3 Then
                    avpItemsTemp = avpItems(3)
                    avpItems(3) = avpItems(ai)
                    avpItems(ai) = avpItemsTemp
                End If
            Next ai
            For ai = 1 To 5
                If UCase(Left(avpItems(ai), 2)) = "CC" And ai <> 4 Then
                    avpItemsTemp = avpItems(4)
                    avpItems(4) = avpItems(ai)
                    avpItems(ai) = avpItemsTemp
                End If
            Next ai
            For ai = 1 To 5
                If UCase(Left(avpItems(ai), 2)) = "BR" And ai <> 5 Then
                    avpItemsTemp = avpItems(5)
                    avpItems(5) = avpItems(ai)
                    avpItems(ai) = avpItemsTemp
                End If
            Next ai
            'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC) " & _
            '        " values(" & UploadNo & ",'','','" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "','" _
            '        & Trim(fnMidValue(ln, 12, 4)) & "','" & Trim(fnMidValue(ln, 20, 3)) & "','" & Trim(fnMidValue(ln, 27, 4)) & "','" _
            '        & Trim(fnMidValue(ln, 35, 2)) & "')"
            Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR) " & _
                    " values(" & UploadNo & ",'','','" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "','" _
                    & Trim(Mid(avpItems(1), 4)) & "','" & Trim(Mid(avpItems(2), 4)) & "','" & Trim(Mid(avpItems(3), 4)) & "','" _
                    & Trim(Mid(avpItems(4), 4)) & "','" & Trim(Mid(avpItems(5), 4)) & "')"
        Else
            Query = "Insert into SabreM9(UPLOADNO,Details1,Details2) " & _
                    " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, Len(ln))) & "')"
        End If
        dbCompany.Execute Query
        Exit Sub
    End If
    
    If M1_flag = True And j < 5 Then
        
        If j = 0 Then
            Query = "update SabreM1 set IU1M3C = '" & Trim(ln) & "' where UPLOADNO = " & M1_UploadNo & " "
        ElseIf j = 1 Then
            Query = "update SabreM1 set IU1M5C = '" & Trim(ln) & "' where UPLOADNO = " & M1_UploadNo & ""
        ElseIf j = 2 Then
            Query = "update SabreM1 set IU1M7C = '" & Trim(ln) & "' where UPLOADNO = " & M1_UploadNo & ""
        ElseIf j = 3 Then
            Query = "update SabreM1 set IU1M8C = '" & Trim(ln) & "' where UPLOADNO = " & M1_UploadNo & ""
        ElseIf j = 4 Then
            Query = "update SabreM1 set IU1M9C = '" & Trim(ln) & "' where UPLOADNO = " & M1_UploadNo & ""
        End If
        dbCompany.Execute Query
        If j = 4 Then
            M1_flag = False
        End If
        j = j + 1
    
    End If
    
    If M2_flag = True And i < 11 Then
        
        If i = 0 Then
            Query = "update SabreM2 set IU2NM4 = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & " "
        ElseIf i = 1 Then
            Query = "update SabreM2 set IU2NM6 = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 2 Then
            Query = "update SabreM2 set IU2FOP = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 3 Then
            Query = "update SabreM2 set IU2ORG = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 4 Then
            Query = "update SabreM2 set IU2IIE = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 5 Then
            Query = "update SabreM2 set IU2END = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 6 Then
            Query = "update SabreM2 set IU2JIC = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 7 Then
            Query = "update SabreM2 set IU2NRT = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 8 Then
            Query = "update SabreM2 set IU2REA = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 9 Then
            Query = "update SabreM2 set IU2DCD = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 10 Then
            Query = "update SabreM2 set IU2SCN = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        End If
        dbCompany.Execute Query
        If i = 10 Then
            M2_flag = False
        End If
        i = i + 1
    End If
    
    If M3_flag = True And id <> "M4" And id <> "M5" And id <> "M6" And id <> "M7" And id <> "M9" Then
        If k = 0 Then
            Query = "update SabreM3 set IU3RRC = '" & Trim(ln) & "' where UPLOADNO = " & M3_UploadNo & " and IU3PC2 = 'HHL'"
        ElseIf k = 1 Then
            Query = "update SabreM3 set IU3HCP = '" & Trim(ln) & "' where UPLOADNO = " & M3_UploadNo & " and IU3PC2 = 'HHL'"
        End If
        k = k + 1
        If k = 2 Then
            k = 0
            M3_flag = False
        End If
        dbCompany.Execute Query
    Else
        M3_flag = False
    End If
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Public Function fnMidValue(str As String, Start As Integer, Ending As Integer) As String
    fnMidValue = Mid(str, Start, Ending)
End Function
Private Sub Form_Load()
On Error GoTo Note
    Dim rsSelect As New ADODB.Recordset
    Me.Icon = FMain.Icon
    If rsSelect.State = 1 Then rsSelect.Close
    rsSelect.Open "Select SABREUPLOADDIRNAME,SABREDESTDIRNAME From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    If rsSelect.EOF = False Then
        txtDirName = IIf(IsNull(rsSelect!SABREUPLOADDIRNAME), "", rsSelect!SABREUPLOADDIRNAME)
        txtRelLocation = IIf(IsNull(rsSelect!SABREDESTDIRNAME), "", rsSelect!SABREDESTDIRNAME)
    End If
    rsSelect.Close
    If rsSelect.State = 1 Then rsSelect.Close
    rsSelect.Open "Select SABRESTATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    chkSabre.value = IIf((IsNull(rsSelect!SABRESTATUS) = True), 0, rsSelect!SABRESTATUS)
    File1.Path = txtDirName.Text
    File1.Pattern = "*.fil"
    Exit Sub
Note:
    If Me.Visible = True Then Unload Me
End Sub
Public Function fnChecktheDatabase(CheckStr As String) As Boolean
Dim rsSelect As New ADODB.Recordset
Dim RecLoc As String

RecLoc = fnMidValue(CheckStr, 54, 8)
If rsSelect.State = 1 Then rsSelect.Close
rsSelect.Open "Select * From SabreM0 Where IU0PNR='" & Trim(RecLoc) & "'", dbCompany, adOpenDynamic, adLockBatchOptimistic
If rsSelect.EOF = False Then
    fnChecktheDatabase = True
Else
    fnChecktheDatabase = False
End If
rsSelect.Close
End Function

