VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProgress 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "PenAIR"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   Icon            =   "FrmProgress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer TimerSave 
      Interval        =   500
      Left            =   1620
      Top             =   1470
   End
   Begin VB.Timer TimerUpdation 
      Interval        =   100
      Left            =   2370
      Top             =   1560
   End
   Begin VB.Frame FraMessage 
      Height          =   1185
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5085
      Begin VB.Image ImgUpdation 
         Height          =   945
         Left            =   60
         Picture         =   "FrmProgress.frx":0ECA
         Top             =   120
         Width           =   945
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1170
         TabIndex        =   1
         Top             =   540
         Width           =   3765
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   63
      ImageHeight     =   63
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":11C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":124B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":1331
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":148A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":1609
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":1788
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":19BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":1C25
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":1ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":21B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":24C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":278B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":2A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":2C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":2EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":308D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":31F5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   90
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   63
      ImageHeight     =   63
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":3321
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":3AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":43F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":4D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":56E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":60A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":6AA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":74A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":7ED8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":892D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":9367
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":9DA2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   720
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   63
      ImageHeight     =   63
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":A806
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":B76B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":C6B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":D5FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":E57E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":F4F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":10466
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":113E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgress.frx":12376
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim viSec As Currency
Dim viImgIndex

Private Sub Form_Load()
    'AnimatedGifCtl1.strGifFileName = "H:\UK_Project\PenAir\SQL Server\Client\Updating.gif"
    'AnimatedGifCtl1.LoadGif (False)
    FraMessage.Left = (ScaleWidth / 2) - (FraMessage.Width / 2)
    viSec = 10
    viImgIndex = 1
    'ImgUpdation.Left = 0
    'ImgUpdation.Top = 0
    ImgUpdation.Stretch = True
    ImgUpdation.Width = ImgUpdation.Width / 1.5
    ImgUpdation.Height = ImgUpdation.Height / 1.5
    FraMessage.Height = ImgUpdation.Top + ImgUpdation.Height + 50
    Height = (Height - (ScaleHeight - (FraMessage.Top + FraMessage.Height))) + 60
    lblMessage.Top = (FraMessage.Height / 2) - (lblMessage.Height / 2)
    lblMessage.Left = ImgUpdation.Left + ImgUpdation.Width + 30
    lblMessage.Width = (FraMessage.Width - lblMessage.Left) - 60
    Me.Refresh
End Sub

Private Sub TimerSave_Timer()
'    viSec = viSec - 0.5
''    If CmdSave.BackColor = &HFF& Then
''        CmdSave.BackColor = &H8000000F
''    Else
''        CmdSave.BackColor = &HFF&
''    End If
'    If viSec = 0 Then cmdOK.Value = True
'    If Int(viSec) = 0 Then
'        cmdOK.Caption = "&OK"
'        Exit Sub
'    End If
'    cmdOK.Caption = "&OK (" & Int(viSec) & ")"
End Sub

Private Sub TimerUpdation_Timer()
    ImgUpdation.Picture = ImageList.ListImages(viImgIndex).Picture
    ImgUpdation.Refresh
    FraMessage.Refresh
    Me.Refresh
    'ImgUpdation.Refresh
    viImgIndex = viImgIndex + 1
    If viImgIndex = ImageList.ListImages.Count Then viImgIndex = 1
End Sub
