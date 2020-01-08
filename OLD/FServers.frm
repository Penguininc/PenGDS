VERSION 5.00
Begin VB.Form FServers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pensoft Servers"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2685
   Icon            =   "FServers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   2685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      MaskColor       =   &H00C8D0D4&
      Picture         =   "FServers.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2310
      Width           =   990
   End
   Begin VB.CommandButton cmdOk 
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
      Height          =   375
      Left            =   240
      MaskColor       =   &H00C8D0D4&
      Picture         =   "FServers.frx":2230
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2310
      Width           =   990
   End
   Begin VB.ListBox lstServers 
      Height          =   2010
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   2325
   End
End
Attribute VB_Name = "FServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOk_Click()
    If Caption = "PenAIR Servers" Then
        vgsServer = lstServers.Text
    Else
        vgsDatabase = lstServers.Text
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    'Me.Icon = FMain.Icon
End Sub

Private Sub lstServers_DblClick()
    cmdOk_Click
End Sub
