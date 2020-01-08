VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4350
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox PWD 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2400
      Width           =   1275
   End
   Begin VB.CommandButton cmdOk 
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
      Left            =   1200
      MaskColor       =   &H00C8D0D4&
      Picture         =   "frmLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   990
   End
   Begin VB.TextBox txtusername 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1305
      MaxLength       =   25
      TabIndex        =   0
      Top             =   375
      Width           =   2115
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
      Left            =   2400
      MaskColor       =   &H00C8D0D4&
      Picture         =   "frmLogin.frx":27E6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   990
   End
   Begin VB.TextBox txtpass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1305
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderStyle     =   3  'Dot
      FillColor       =   &H000000C0&
      Height          =   1935
      Index           =   0
      Left            =   120
      Top             =   75
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   1935
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3480
      Picture         =   "frmLogin.frx":3B4C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   450
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   900
      Width           =   690
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Private Sub cmdCancel_Click()
    mLogin = False
    Unload Me
End Sub
Private Sub cmdOk_Click()
Dim rsLogin As New ADODB.Recordset
If Trim(txtusername) = "" Then MsgBox "Invalid User Name", vbInformation, "PenGDS": txtusername.SetFocus: txtusername.SelLength = 25: Exit Sub
If Trim(txtpass) = "" Then MsgBox "Invalid Password", vbInformation, "PenGDS": txtpass.SetFocus: Exit Sub

SQL = "Select * From Users Where UNAME='" & Replace(Trim(txtusername), "'", "''") & "'"

rsLogin.Open SQL, dbCompany, adOpenDynamic, adLockBatchOptimistic
If rsLogin.EOF = True Then
    MsgBox "Invalid User Name ", vbInformation, "PenGDS": txtusername.SetFocus:  Exit Sub
End If
rsLogin.Close

SQL = "Select * From Users Where UPWD='" & Replace(Trim(txtpass), "'", "''") & "'"
rsLogin.Open SQL, dbCompany, adOpenDynamic, adLockBatchOptimistic
If rsLogin.EOF = True Then
    MsgBox "Invalid Password ", vbInformation, "PenGDS": txtpass.SetFocus: Exit Sub
End If
rsLogin.Close

If UCase(Trim(txtusername)) = UCase("Admin") Then
    SQL = "Select * From Users Where UNAME='" & Replace(Trim(txtusername), "'", "''") & "' And " & _
            "UPWD='" & Replace(Trim(txtpass), "'", "''") & "'"
    rsLogin.Open SQL, dbCompany, adOpenDynamic, adLockBatchOptimistic
    If rsLogin.EOF = False Then
        mLogin = True
    Else
        MsgBox "Invalid Login", vbInformation, "PenGDS": txtpass.SetFocus: Exit Sub
    End If
    rsLogin.Close
Else
    MsgBox "Access Denied!", vbInformation, "PenGDS"
End If
'If mUserId > 0 Then Call AssignUserRights
Unload Me
End Sub
Private Sub Form_Load()
Loginmakeround Me, hwnd
Shape1(0).Top = 60: Shape1(0).Visible = False
Shape1(1).Top = 75: Shape1(1).Visible = False
'mLogin = False
End Sub

Private Sub txtpass_GotFocus()
    cmdOk.default = True
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOk.SetFocus
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtpass.SetFocus
End Sub
Public Sub Loginmakeround(FormName As Object, handle As Long)
X = (FormName.Width / Screen.TwipsPerPixelX)
Y = (FormName.Height / Screen.TwipsPerPixelY)
N = 25
SetWindowRgn handle, CreateRoundRectRgn(0, 0, X, Y, N, N), True
'FormName.Icon = FMain.Icon
End Sub
