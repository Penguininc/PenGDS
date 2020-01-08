VERSION 5.00
Begin VB.Form FServers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pensoft Servers"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "FServers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerEnd 
      Interval        =   1
      Left            =   30
      Top             =   2250
   End
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
      Left            =   2100
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
      Left            =   900
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
Dim fRow_Long As Long, fRows_Long As Long
Dim fSearchString_String As String
'by Abhi on 25-Sep-2014 for caseid 4552 modal form will not showing in taskbar
Private Const WS_EX_APPWINDOW               As Long = &H40000
Private Const GWL_EXSTYLE                   As Long = (-20)
Private Const SW_HIDE                       As Long = 0
Private Const SW_SHOW                       As Long = 5

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private m_bActivated As Boolean
'by Abhi on 25-Sep-2014 for caseid 4552 modal form will not showing in taskbar

Private Sub cmdCancel_Click()
    INIWrite INIPenAIR_String, "PenGDS", "GDS", "OFF"
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

Private Sub Form_Activate()
    'by Abhi on 25-Sep-2014 for caseid 4552 modal form will not showing in taskbar
    If Not m_bActivated Then
        m_bActivated = True
        Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
        Call ShowWindow(hwnd, SW_HIDE)
        Call ShowWindow(hwnd, SW_SHOW)
    End If
    'by Abhi on 25-Sep-2014 for caseid 4552 modal form will not showing in taskbar
End Sub

Private Sub Form_Load()
    'Me.Icon = FMain.Icon
    lstServers.Width = ScaleWidth - (lstServers.Left + lstServers.Left)
    'by Abhi on 25-Sep-2014 for caseid 4552 modal form will not showing in taskbar
    m_bActivated = False
    'by Abhi on 25-Sep-2014 for caseid 4552 modal form will not showing in taskbar
End Sub

Private Sub lstServers_DblClick()
    cmdOk_Click
End Sub

Private Sub lstServers_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        lstServers.Selected(0) = True
        fSearchString_String = ""
        Exit Sub
    End If
    fSearchString_String = fSearchString_String & Chr(KeyAscii)
    fRows_Long = lstServers.ListCount
    For fRow_Long = 0 To fRows_Long
        If InStr(1, lstServers.List(fRow_Long), fSearchString_String, vbTextCompare) > 0 Then
            lstServers.Selected(fRow_Long) = True
            KeyAscii = 0
            Exit For
        End If
        DoEvents
    Next
End Sub

Private Sub TimerEnd_Timer()
On Error Resume Next
    If (UCase(Dir(App.Path & "\EndSQL.txt")) = UCase("EndSQL.txt") Or UCase(Dir(App.Path & "\EndGDS.txt")) = UCase("EndGDS.txt")) And fsObj.FileExists(App.Path & "\_UploadingSQL_") = False Then
        'by Abhi on 22-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module
        'Open (App.Path & "\ENDGDS.TXT") For Random As #1
        'by Abhi on 25-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module is Append
        'Open (App.Path & "\ENDGDS.TXT") For Output Shared As #1
        Open (App.Path & "\ENDGDS.TXT") For Append Shared As #1
            'by Abhi on 22-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module
            Print #1, "Updating... PenGDS-FServers-tmrEnd_Timer()"
        Close #1
        End
    End If
End Sub
