VERSION 5.00
Object = "{9F8F4546-0259-44B3-AA22-CFE22545C52F}#1.0#0"; "lvButton.ocx"
Begin VB.Form DetailsForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Details"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox DetailsText 
      BackColor       =   &H8000000F&
      Height          =   1665
      Left            =   540
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   270
      Width           =   2625
   End
   Begin lvButton.chameleonButton CloseCommand 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Esc"
      Top             =   2250
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "DetailsForm.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin lvButton.chameleonButton RefreshCommand 
      Height          =   375
      Left            =   570
      TabIndex        =   2
      ToolTipText     =   "F5"
      Top             =   3000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Refresh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "DetailsForm.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "DetailsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
Option Explicit
Private Sub CloseCommand_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        Call RefreshCommand_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = FMain.Icon
End Sub

Private Sub Form_Resize()
On Error Resume Next
    CloseCommand.Top = ScaleHeight - CloseCommand.Height - 150
    
    DetailsText.Top = ScaleTop + 150
    DetailsText.Left = ScaleLeft + 150
    DetailsText.Height = CloseCommand.Top - (150 + 150)
    DetailsText.Width = ScaleWidth - (150 + 150)
    
    CloseCommand.Left = (ScaleWidth - DetailsText.Left) - (CloseCommand.Width)
    
    RefreshCommand.Top = CloseCommand.Top
    RefreshCommand.Left = DetailsText.Left
End Sub

Public Function ShowME(ByVal pColumnHeaderName_String As String, ByVal pDetails_String As String)
    Me.Caption = pColumnHeaderName_String
    DetailsText.Text = pDetails_String
    Me.Show , FMain
End Function
'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
Private Sub RefreshCommand_Click()
    '########################################################################################################################
    'SendERROREmailPreviousMess_String = SendERROREmailPreviousMess_String & "" _
    '    & "10/Aug/2016 05:48:33 PM(17:48:33) : Error: 0 - File is in use by another application and cannot be accessed. (Worldspan - 15221727.prt). PenGDS is automatically Resumed." _
    '    & vbCrLf _
    '    & vbCrLf _
    '    & "12/Aug/2016 10:47:21 AM(10:47:21) : Error: -2147467259 - Data provider or other service returned an E_FAIL status." & vbCrLf _
    '    & "MSGTAG=KFT, Table=AmdLineKFT, Field=PSGRID, Field does not exist in Collection!, " & vbCrLf _
    '    & "(Amadeus - AIR00002_20160811125358.Txt) Moved to folder 'Error Files'. PenGDS is automatically Resumed." _
    '    & vbCrLf _
    '    & vbCrLf
    '########################################################################################################################
    
    DetailsText.Text = SendERROREmailPreviousMess_String
End Sub
