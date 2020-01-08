VERSION 5.00
Begin VB.Form ShowFormNameForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screen"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "ShowFormNameForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CloseCommand 
      Cancel          =   -1  'True
      Caption         =   "Close"
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
      Height          =   345
      Left            =   3540
      TabIndex        =   2
      Top             =   540
      Width           =   1005
   End
   Begin VB.TextBox FormNameText 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   3525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Screen"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "ShowFormNameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by Abhi on 28-Jul-2009
Option Explicit

Private Sub CloseCommand_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'commented by Abhi on 23-Jul-2010 for when getting error show this form from login form
    'Me.Icon = FMain.Icon
End Sub

Public Function ShowME(ByVal pForm As Form) As Boolean
    'by Abhi on 05-Oct-2013 for caseid 3425 Shortcut for show screen name is added with application name in penair
    'FormNameText = GetFormName(pForm)
    FormNameText = "PenGDS." & GetFormName(pForm)
    'by Abhi on 05-Oct-2013 for caseid 3425 Shortcut for show screen name is added with application name in penair
    Me.Show vbModal
End Function

Private Sub FormNameText_GotFocus()
    SelectText FormNameText
End Sub
