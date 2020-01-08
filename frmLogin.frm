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
   Begin VB.Label SecuredLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "128-bit PGP 1.0 Secured"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1305
      TabIndex        =   7
      Top             =   1200
      Width           =   2115
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
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
Dim rsLogin As New ADODB.Recordset
'by Abhi on 20-Oct-2009 for Penguin Encryption
Dim vUPWD_String As String
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
Dim DeadlockRETRY_Integer As Integer
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS

If Trim(txtusername) = "" Then MsgBox "Invalid User Name", vbCritical, App.Title: txtusername.SetFocus: txtusername.SelLength = 25: Exit Sub
If Trim(txtpass) = "" Then MsgBox "Invalid Password", vbCritical, App.Title: txtpass.SetFocus: Exit Sub
'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
DeadlockRETRY:
'by Abhi on 20-Oct-2009 for Penguin Encryption
vUPWD_String = Trim(txtpass)
If Encrypted_Boolean Then
    vUPWD_String = PENEncryptDecrypt(Encrypt, vgsDatabase, vUPWD_String)
End If

'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
'SQL = "Select * From Users Where UNAME='" & Replace(Trim(txtusername), "'", "''") & "'"
SQL = "Select UNAME From Users WITH (NOLOCK) Where UNAME='" & Replace(Trim(txtusername), "'", "''") & "'"
'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
'rsLogin.Open SQL, dbCompany, adOpenDynamic, adLockBatchOptimistic
rsLogin.Open SQL, dbCompany, adOpenForwardOnly, adLockReadOnly
If rsLogin.EOF = True Then
    MsgBox "Invalid User Name ", vbCritical, App.Title: txtusername.SetFocus:   Exit Sub
End If
rsLogin.Close

'by Abhi on 20-Oct-2009 for Penguin Encryption
'SQL = "Select * From Users Where UPWD='" & Replace(Trim(txtpass), "'", "''") & "'"
'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
'SQL = "Select * From Users Where UPWD='" & Replace(Trim(vUPWD_String), "'", "''") & "'"
SQL = "Select UPWD From Users WITH (NOLOCK) Where UPWD='" & Replace(Trim(vUPWD_String), "'", "''") & "'"
'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
'rsLogin.Open SQL, dbCompany, adOpenDynamic, adLockBatchOptimistic
rsLogin.Open SQL, dbCompany, adOpenForwardOnly, adLockReadOnly
If rsLogin.EOF = True Then
    MsgBox "Invalid Password ", vbCritical, App.Title: txtpass.SetFocus: Exit Sub
End If
rsLogin.Close

If UCase(Trim(txtusername)) = UCase("Admin") Then
    'by Abhi on 20-Oct-2009 for Penguin Encryption
    'SQL = "Select * From Users Where UNAME='" & Replace(Trim(txtusername), "'", "''") & "' And " & _
            "UPWD='" & Replace(Trim(txtpass), "'", "''") & "'"
    'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'SQL = "Select * From Users Where UNAME='" & Replace(Trim(txtusername), "'", "''") & "' And " & _
            "UPWD='" & Replace(Trim(vUPWD_String), "'", "''") & "'"
    SQL = "Select UNAME From Users WITH (NOLOCK) Where UNAME='" & Replace(Trim(txtusername), "'", "''") & "' And " & _
            "UPWD='" & Replace(Trim(vUPWD_String), "'", "''") & "'"
    'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsLogin.Open SQL, dbCompany, adOpenDynamic, adLockBatchOptimistic
    rsLogin.Open SQL, dbCompany, adOpenForwardOnly, adLockReadOnly
    If rsLogin.EOF = False Then
        mLogin = True
    Else
        MsgBox "Invalid Login", vbCritical, App.Title: txtpass.SetFocus: Exit Sub
    End If
    rsLogin.Close
Else
    MsgBox "Access Denied!", vbCritical, App.Title
End If
'If mUserId > 0 Then Call AssignUserRights
Unload Me
Exit Sub
PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    'by Abhi on 24-Jul-2009 for Deadlock
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    'If PENErr_Number = -2147467259 Then 'Deadlock
    If (PENErr_Number = -2147467259 Or PENErr_Number = -2147217871) And DeadlockRETRY_Integer < 3 Then '-2147467259 Deadlock, -2147217871 Query timeout expired
        DeadlockRETRY_Integer = DeadlockRETRY_Integer + 1
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
        Debug.Print "Deadlock"
        'commented by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
        'Resume
        'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
        Sleep 5
        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
        'GoTo DeadlockRETRY
        Resume DeadlockRETRY
        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    End If
    Screen.MousePointer = vbNormal
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title
    Exit Sub
End Sub
Private Sub Form_Load()
'Loginmakeround Me, hwnd
'Shape1(0).Top = 60: Shape1(0).Visible = False
'Shape1(1).Top = 75: Shape1(1).Visible = False
'mLogin = False
    If Encrypted_Boolean = False Then
        SecuredLabel.Visible = False
    End If
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
x = (FormName.Width / Screen.TwipsPerPixelX)
y = (FormName.Height / Screen.TwipsPerPixelY)
N = 25
SetWindowRgn handle, CreateRoundRectRgn(0, 0, x, y, N, N), True
'FormName.Icon = FMain.Icon
End Sub
