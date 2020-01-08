VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FRegisterForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FRegisterForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   240
      Top             =   2730
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CancelCommand 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   3
      Top             =   2730
      Width           =   915
   End
   Begin VB.CommandButton RegisterCommand 
      Caption         =   "Register"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3690
      TabIndex        =   2
      Top             =   2730
      Width           =   915
   End
   Begin VB.Frame FirstFrame 
      Caption         =   "Client Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   90
      TabIndex        =   4
      Top             =   420
      Width           =   4455
      Begin VB.TextBox ComputerText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3150
         TabIndex        =   1
         Top             =   720
         Width           =   1155
      End
      Begin VB.TextBox NameText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   0
         Top             =   720
         Width           =   1725
      End
      Begin VB.Label AtLabel 
         AutoSize        =   -1  'True
         Caption         =   "@"
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
         Left            =   2040
         TabIndex        =   9
         Top             =   720
         Width           =   210
      End
      Begin VB.Label ComputerLabel 
         AutoSize        =   -1  'True
         Caption         =   "Computer :"
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
         Left            =   1980
         TabIndex        =   8
         Top             =   450
         Width           =   960
      End
      Begin VB.Label StatusLabel 
         Alignment       =   2  'Center
         Caption         =   "Error"
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   300
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label NameLabel 
         AutoSize        =   -1  'True
         Caption         =   "Enter Client Name :"
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
         Left            =   150
         TabIndex        =   6
         Top             =   420
         Width           =   1695
      End
   End
   Begin VB.Label HeadLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " PenGDS Interface"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   30
      Width           =   4485
   End
End
Attribute VB_Name = "FRegisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelCommand_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    StatusLabel.Caption = ""
    ComputerText.Text = UCase(Winsock.LocalHostName)
    HeadLabel.Caption = HeadLabel.Caption & " [" & vgsDatabase & "]"
End Sub

Private Sub Form_Resize()
    HeadLabel.Left = 0
    HeadLabel.Top = 0
    HeadLabel.Width = ScaleWidth
    FirstFrame.Top = (HeadLabel.Top + HeadLabel.Height) + 60
    FirstFrame.Left = 120
    FirstFrame.Width = (ScaleWidth - FirstFrame.Left) - FirstFrame.Left
    RegisterCommand.Top = FirstFrame.Top + FirstFrame.Height + 120
    RegisterCommand.Left = ScaleWidth - RegisterCommand.Width - FirstFrame.Left
    CancelCommand.Top = RegisterCommand.Top
    CancelCommand.Left = RegisterCommand.Left - FirstFrame.Left - CancelCommand.Width
    Me.ScaleHeight = CancelCommand.Top + CancelCommand.Height + RegisterCommand.Left
    
    NameLabel.Left = 100
    NameText.Left = NameLabel.Left
    NameText.Width = 2480 'FirstFrame.Width - NameText.Left - NameText.Left
    AtLabel.Top = NameText.Top + 60
    AtLabel.Left = NameText.Left + NameText.Width + 10
    ComputerLabel.Top = NameLabel.Top
    ComputerText.Left = AtLabel.Left + AtLabel.Width + 20
    ComputerLabel.Left = ComputerText.Left
    ComputerText.Width = 1500
    StatusLabel.Left = NameLabel.Left
    StatusLabel.Width = FirstFrame.Width - NameText.Left - NameText.Left
End Sub

Private Sub NameText_KeyPress(KeyAscii As Integer)
    If NameText.SelStart = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub RegisterCommand_Click()
    If Trim(NameText) = "" Or Trim(ComputerText) = "" Then Exit Sub
    RegisterCommand.Enabled = False
    StatusLabel.ForeColor = &HFF0000
    StatusLabel.Caption = "Please wait..."
    DoEvents
    'If Winsock.State <> sckConnected Then
        Winsock.Close
        Winsock.Connect Host, Port
    'End If
    If Winsock.State = sckConnected Then
        Winsock.SendData "Register;" & "PenGDS Interface" & ";" & Trim(NameText.Text) & "@" & Trim(ComputerText)
    End If
        
    'StatusLabel.Caption = "Client name is already exsists." & vbCrLf & "Please try another name."
End Sub

Private Sub Winsock_Connect()
    If Winsock.State = sckConnected Then
        Winsock.SendData "Register;" & "PenGDS Interface" & ";" & Trim(NameText.Text) & "@" & Trim(ComputerText)
    End If
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim vlsData As String
    Winsock.GetData vlsData
    If vlsData = "Yes" Then
        StatusLabel.ForeColor = &HC000&
        StatusLabel.Caption = "Success"
        ClientName = Trim(NameText.Text) & "@" & Trim(ComputerText.Text)
        INIWrite INIPenAIR_String, "PenGDS", "Monitor", 1
        INIWrite INIPenAIR_String, "PenGDS", "Name", ClientName
        INIWrite INIPenAIR_String, "PenGDS", "Host", Host
        INIWrite INIPenAIR_String, "PenGDS", "Port", Port
        MsgBox "Thank you for registering", vbInformation, App.Title
        Unload Me
    Else
        StatusLabel.ForeColor = &HFF&
        StatusLabel.Caption = "Client name is already exsists." & vbCrLf & "Please try another name."
        RegisterCommand.Enabled = True
    End If
End Sub

