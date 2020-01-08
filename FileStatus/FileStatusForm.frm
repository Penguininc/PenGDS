VERSION 5.00
Begin VB.Form FileStatusForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileStatus"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   7065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CloseCommand 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   405
      Left            =   4425
      TabIndex        =   3
      Top             =   1170
      Width           =   1185
   End
   Begin VB.CommandButton UnlockCommand 
      Caption         =   "Unlock"
      Enabled         =   0   'False
      Height          =   405
      Left            =   2940
      TabIndex        =   2
      Top             =   1170
      Width           =   1185
   End
   Begin VB.CommandButton LockCommand 
      Caption         =   "Lock"
      Height          =   405
      Left            =   1455
      TabIndex        =   1
      Top             =   1170
      Width           =   1185
   End
   Begin VB.TextBox FilepathandNameText 
      Height          =   330
      Left            =   300
      TabIndex        =   0
      Text            =   "D:\UK_Project\GDS\Worldspan\Theflightsguru\Waiting for file available\16523711.prt"
      Top             =   540
      Width           =   6465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Path \ Name"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   300
      Width           =   1200
   End
End
Attribute VB_Name = "FileStatusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseCommand_Click()
    End
End Sub

Private Sub Form_Activate()
    LockCommand.SetFocus
End Sub

Private Sub LockCommand_Click()
    If Trim(FilepathandNameText) = "" Then Exit Sub
    If Dir(FilepathandNameText) = "" Then
        MsgBox "File not found", vbCritical, App.Title
        Exit Sub
    End If
    LockCommand.Enabled = False
    UnlockCommand.Enabled = True
    UnlockCommand.SetFocus
    Open FilepathandNameText For Binary Access Read Write Lock Read Write As #1
    Me.Caption = "FileStatus (Locked)"
End Sub

Private Sub UnlockCommand_Click()
    If Trim(FilepathandNameText) = "" Then Exit Sub
    UnlockCommand.Enabled = False
    LockCommand.Enabled = True
    LockCommand.SetFocus
    Close #1
    Me.Caption = "FileStatus (Unlocked)"
End Sub

Sub ListWins(Optional Title = "*", Optional Class = "*")
    Dim hWndThis As Long
    hWndThis = FindWindow(vbNullString, vbNullString)
    While hWndThis
        Dim sTitle As String, sClass As String
        sTitle = Space$(255)
        sTitle = Left$(sTitle, GetWindowText(hWndThis, sTitle, Len(sTitle)))
        sClass = Space$(255)
        sClass = Left$(sClass, GetClassName(hWndThis, sClass, Len(sClass)))
        If sTitle Like Title And sClass Like Class Then
            Debug.Print sTitle, sClass
        End If
        hWndThis = GetWindow(hWndThis, GW_HWNDNEXT)
    Wend
End Sub
