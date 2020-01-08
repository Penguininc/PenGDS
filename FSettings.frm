VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelCommand 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   2970
      Width           =   945
   End
   Begin VB.CommandButton SaveCommand 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2490
      TabIndex        =   5
      Top             =   2970
      Width           =   945
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2925
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   5159
      _Version        =   393216
      TabHeight       =   617
      TabMaxWidth     =   2293
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Email"
      TabPicture(0)   =   "FSettings.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Network"
      TabPicture(1)   =   "FSettings.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Settings"
      TabPicture(2)   =   "FSettings.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   2265
         Left            =   -74850
         TabIndex        =   6
         Top             =   450
         Width           =   4215
         Begin VB.TextBox EffectiveIDText 
            BackColor       =   &H8000000F&
            Height          =   310
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox PENLINEIDText 
            Height          =   310
            Left            =   1590
            TabIndex        =   10
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox PermissionDeniedText 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1590
            TabIndex        =   8
            Top             =   210
            Width           =   675
         End
         Begin VB.Label Label10 
            Caption         =   "Penline ID"
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   630
            Width           =   795
         End
         Begin VB.Label PermissionDeniedLabel 
            Caption         =   "x 0.5 sec = "
            Height          =   285
            Left            =   2340
            TabIndex        =   9
            Top             =   240
            Width           =   1725
         End
         Begin VB.Label Label1 
            Caption         =   "Permission Denied"
            Height          =   285
            Left            =   150
            TabIndex        =   7
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2265
         Left            =   -74850
         TabIndex        =   3
         Top             =   450
         Width           =   4215
         Begin VB.CheckBox SpecialNetworkCheck 
            Caption         =   "Use Special Network Exists chceking"
            Height          =   225
            Left            =   180
            TabIndex        =   4
            Top             =   270
            Width           =   3735
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Email Test"
         Height          =   555
         Left            =   1350
         TabIndex        =   2
         Top             =   1260
         Width           =   1185
      End
      Begin VB.Label lblFName 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   570
         Visible         =   0   'False
         Width           =   645
      End
   End
End
Attribute VB_Name = "FSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelCommand_Click()
    Unload Me
End Sub


Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'by Abhi on 18-Nov-2015 for caseid 5798 PenGDS Email Temporary Folder should use user’s session’s temporary folders instead of user’s standard Temp folder
Dim vPrompt_String As String
Dim vVbMsgBoxStyle As VbMsgBoxStyle
    If Shift = SHIFT_MASK + CTRL_MASK + ALT_MASK Then
        If Button = vbRightButton Then
            vPrompt_String = "PenGDSTemporaryFolder: " & vbCrLf & vbCrLf & PenEMAILTemporaryFolder_String
            vVbMsgBoxStyle = vbInformation
            If Trim(PenEMAILTemporaryFolder_String) <> "" Then
                vPrompt_String = vPrompt_String & vbCrLf & vbCrLf & "Press OK to Open in new window"
                vVbMsgBoxStyle = vVbMsgBoxStyle + vbOKCancel
            End If
            If MsgBox(vPrompt_String, vVbMsgBoxStyle) = vbOK Then
                If Trim(PenEMAILTemporaryFolder_String) <> "" Then
                    Shell "explorer.exe " & PenEMAILTemporaryFolder_String & "", vbNormalFocus
                End If
            End If
        End If
    End If
'by Abhi on 18-Nov-2015 for caseid 5798 PenGDS Email Temporary Folder should use user’s session’s temporary folders instead of user’s standard Temp folder
End Sub

Private Sub SaveCommand_Click()
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
Dim DeadlockRETRY_Integer As Integer
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    
'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
DeadlockRETRY:
    dbCompany.BeginTrans
    dbCompany.Execute "Update [File] set PENLINEID='" & SkipChars(PENLINEIDText) & "'"
    dbCompany.CommitTrans
    Unload Me
Exit Sub

PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
    If PENErr_Number <> -2147168242 Then dbCompany.RollbackTrans
    'by Abhi on 24-Jul-2009 for Deadlock
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    'If PENErr_Number = -2147467259 Then 'Deadlock
    If (PENErr_Number = -2147467259 Or PENErr_Number = -2147217871) And DeadlockRETRY_Integer < 3 Then '-2147467259 Deadlock, -2147217871 Query timeout expired
        DeadlockRETRY_Integer = DeadlockRETRY_Integer + 1
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
        Debug.Print "Deadlock"
        'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
        'Resume
        'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
        Sleep 5
        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
        'GoTo DeadlockRETRY
        Resume DeadlockRETRY
        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    End If
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & "(Settings-Save)"
    Exit Sub
End Sub

Private Sub Command1_Click()
Dim vSuccess As Boolean
    'vSuccess = SendERROR("Email Testing from PenGDS[" & CHEAD & "] (" & vgsDatabase & ")", "This is a Email testing mail so please ignore.")
    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
    'vSuccess = SendERROR("Email Testing from " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "This is a Email testing mail so please ignore.")
    vSuccess = SendERROR("Email Testing from " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "This is a test email - Please ignore.", "")
    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
    If vSuccess = True Then
        MsgBox "Email sent successfully", vbInformation, App.Title
    Else
        MsgBox "Email senting failed", vbCritical, App.Title
    End If
End Sub

Private Sub Form_Load()
    
    Me.Icon = FMain.Icon
    SpecialNetworkCheck.value = Val(INIRead(INIPenAIR_String, "PenGDS", "SpecialNetwork", "0"))
    PermissionDeniedText = Val(INIRead(INIPenAIR_String, "PenGDS", "PermissionDenied", "5"))
    PENLINEIDText = Trim(getFromFileTable("PENLINEID"))
End Sub

Private Sub PENLINEIDText_Change()
    If Trim(PENLINEIDText) = "" Then
        EffectiveIDText = ""
    Else
        EffectiveIDText = "P" & Trim(PENLINEIDText)
    End If
End Sub

Private Sub PermissionDeniedText_Change()
    PermissionDeniedLabel = "x 0.5 sec = " & Val(PermissionDeniedText) * 0.5 & " sec"
    INIWrite INIPenAIR_String, "PenGDS", "PermissionDenied", Val(PermissionDeniedText)
End Sub

Private Sub SpecialNetworkCheck_Click()
    INIWrite INIPenAIR_String, "PenGDS", "SpecialNetwork", Val(SpecialNetworkCheck.value)
End Sub
