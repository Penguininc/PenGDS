VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{9F8F4546-0259-44B3-AA22-CFE22545C52F}#1.0#0"; "lvbutton.ocx"
Object = "{16920A9E-0846-11D5-8B89-BD3D07939431}#1.0#0"; "TRAYAREA.OCX"
Begin VB.Form FMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PenGDS Interface"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.chameleonButton cmdStart 
      Height          =   405
      Left            =   3030
      TabIndex        =   0
      Top             =   3030
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Start"
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
      FCOL            =   49152
      FCOLO           =   65280
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FMain.frx":0ECA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkEnableAll 
      Caption         =   "Enable All"
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   3000
      Width           =   1035
   End
   Begin VB.CheckBox chkAutoStart 
      Caption         =   "Auto Start"
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   3270
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2925
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   5159
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
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
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FMain.frx":0EE6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "prgUpload"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tmrEnd"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tmrUpload"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TrayArea1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Galileo"
      TabPicture(1)   =   "FMain.frx":0F02
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(1)=   "chkGalilieo"
      Tab(1).Control(2)=   "txtGalilieoSource"
      Tab(1).Control(3)=   "txtGalilieoDest"
      Tab(1).Control(4)=   "lblSource(0)"
      Tab(1).Control(5)=   "lblDestination(1)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Sabre"
      TabPicture(2)   =   "FMain.frx":0F1E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture3"
      Tab(2).Control(1)=   "txtDest"
      Tab(2).Control(2)=   "txtSource"
      Tab(2).Control(3)=   "chkSabre"
      Tab(2).Control(4)=   "lblDestination(0)"
      Tab(2).Control(5)=   "lblSource(1)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Worldspan"
      TabPicture(3)   =   "FMain.frx":0F3A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture4"
      Tab(3).Control(1)=   "txtWorldspanDest"
      Tab(3).Control(2)=   "txtWorldspanSource"
      Tab(3).Control(3)=   "chkWorldspan"
      Tab(3).Control(4)=   "lblDestination(2)"
      Tab(3).Control(5)=   "lblSource(2)"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Amadeus"
      TabPicture(4)   =   "FMain.frx":0F56
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblSource(3)"
      Tab(4).Control(1)=   "lblDestination(3)"
      Tab(4).Control(2)=   "chkAmadeus"
      Tab(4).Control(3)=   "txtAmadeusSource"
      Tab(4).Control(4)=   "txtAmadeusDest"
      Tab(4).Control(5)=   "Picture5"
      Tab(4).ControlCount=   6
      Begin SystemTrayControl.TrayArea TrayArea1 
         Left            =   420
         Top             =   1200
         _ExtentX        =   900
         _ExtentY        =   900
      End
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   -70935
         Picture         =   "FMain.frx":0F72
         ScaleHeight     =   690
         ScaleWidth      =   2520
         TabIndex        =   34
         Top             =   550
         Width           =   2520
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   -70935
         Picture         =   "FMain.frx":1A36
         ScaleHeight     =   690
         ScaleWidth      =   2520
         TabIndex        =   33
         Top             =   550
         Width           =   2520
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   -70905
         Picture         =   "FMain.frx":2B53
         ScaleHeight     =   690
         ScaleWidth      =   2520
         TabIndex        =   32
         Top             =   550
         Width           =   2520
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   -70935
         Picture         =   "FMain.frx":3210
         ScaleHeight     =   690
         ScaleWidth      =   2520
         TabIndex        =   31
         Top             =   550
         Width           =   2520
      End
      Begin VB.TextBox txtAmadeusDest 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73650
         TabIndex        =   28
         Top             =   2340
         Width           =   5265
      End
      Begin VB.TextBox txtAmadeusSource 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73650
         TabIndex        =   27
         Top             =   1920
         Width           =   5265
      End
      Begin VB.CheckBox chkAmadeus 
         Alignment       =   1  'Right Justify
         Caption         =   "Enable"
         Height          =   315
         Left            =   -74820
         TabIndex        =   26
         Top             =   720
         Width           =   1365
      End
      Begin VB.TextBox txtWorldspanDest 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73650
         TabIndex        =   23
         Top             =   2340
         Width           =   5265
      End
      Begin VB.TextBox txtWorldspanSource 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73650
         TabIndex        =   22
         Top             =   1920
         Width           =   5265
      End
      Begin VB.CheckBox chkWorldspan 
         Alignment       =   1  'Right Justify
         Caption         =   "Enable"
         Height          =   315
         Left            =   -74820
         TabIndex        =   21
         Top             =   720
         Width           =   1365
      End
      Begin VB.TextBox txtDest 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73650
         TabIndex        =   18
         Top             =   2340
         Width           =   5265
      End
      Begin VB.TextBox txtSource 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73650
         TabIndex        =   17
         Top             =   1920
         Width           =   5265
      End
      Begin VB.CheckBox chkSabre 
         Alignment       =   1  'Right Justify
         Caption         =   "Enable"
         Height          =   315
         Left            =   -74820
         TabIndex        =   16
         Top             =   720
         Width           =   1365
      End
      Begin VB.Timer tmrUpload 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   330
         Top             =   540
      End
      Begin VB.Timer tmrEnd 
         Interval        =   1
         Left            =   810
         Top             =   540
      End
      Begin VB.CheckBox chkGalilieo 
         Alignment       =   1  'Right Justify
         Caption         =   "Enable"
         Height          =   315
         Left            =   -74820
         TabIndex        =   6
         Top             =   720
         Width           =   1365
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   2910
         Picture         =   "FMain.frx":3E30
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   13
         Top             =   570
         Width           =   960
      End
      Begin VB.TextBox txtGalilieoSource 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73650
         TabIndex        =   9
         Top             =   1920
         Width           =   5265
      End
      Begin VB.TextBox txtGalilieoDest 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73650
         TabIndex        =   11
         Top             =   2340
         Width           =   5265
      End
      Begin MSComctlLib.ProgressBar prgUpload 
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   2310
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblDestination 
         AutoSize        =   -1  'True
         Caption         =   "Destination"
         Height          =   195
         Index           =   3
         Left            =   -74790
         TabIndex        =   30
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   3
         Left            =   -74790
         TabIndex        =   29
         Top             =   1980
         Width           =   510
      End
      Begin VB.Label lblDestination 
         AutoSize        =   -1  'True
         Caption         =   "Destination"
         Height          =   195
         Index           =   2
         Left            =   -74790
         TabIndex        =   25
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   2
         Left            =   -74790
         TabIndex        =   24
         Top             =   1980
         Width           =   510
      End
      Begin VB.Label lblDestination 
         AutoSize        =   -1  'True
         Caption         =   "Destination"
         Height          =   195
         Index           =   0
         Left            =   -74790
         TabIndex        =   20
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   1
         Left            =   -74790
         TabIndex        =   19
         Top             =   1980
         Width           =   510
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
         TabIndex        =   15
         Top             =   1950
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "PEN GDS INTERFACE"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   165
         TabIndex        =   7
         Top             =   1680
         Width           =   6495
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   0
         Left            =   -74790
         TabIndex        =   12
         Top             =   1980
         Width           =   510
      End
      Begin VB.Label lblDestination 
         AutoSize        =   -1  'True
         Caption         =   "Destination"
         Height          =   195
         Index           =   1
         Left            =   -74790
         TabIndex        =   10
         Top             =   2400
         Width           =   795
      End
   End
   Begin MSComctlLib.StatusBar stbUpload 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   3540
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6006
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6006
         EndProperty
      EndProperty
   End
   Begin lvButton.chameleonButton cmdStop 
      Height          =   405
      Left            =   4320
      TabIndex        =   1
      Top             =   3030
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Stop"
      ENAB            =   0   'False
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
      FCOL            =   192
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FMain.frx":547A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin lvButton.chameleonButton cmdExit 
      Height          =   405
      Left            =   5610
      TabIndex        =   2
      Top             =   3030
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Exit"
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
      FCOL            =   12582912
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FMain.frx":5496
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1410
      Picture         =   "FMain.frx":54B2
      Top             =   3090
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Upload_Status As Integer
Dim Upload_StatusMIR As Integer
Dim Upload_StatusAIR As Integer
Dim Upload_StatusWORLD As Integer

Private Sub chkAmadeus_Click()
    If chkAmadeus.value = vbChecked Then
        txtAmadeusSource.Locked = True
        txtAmadeusSource.BackColor = &H8000000F
        txtAmadeusDest.Locked = True
        txtAmadeusDest.BackColor = &H8000000F
        SSTab1.TabPicture(4) = Image3.Picture
    Else
        txtAmadeusSource.Locked = False
        txtAmadeusSource.BackColor = &H80000005
        txtAmadeusDest.Locked = False
        txtAmadeusDest.BackColor = &H80000005
        SSTab1.TabPicture(4) = LoadPicture()
    End If
    dbCompany.Execute "Update [File] set AMDSTATUS=" & chkAmadeus.value
    SetEnableAll
End Sub

Private Sub chkAutoStart_Click()
    dbCompany.Execute "Update [File] set AUTOSTART=" & chkAutoStart.value
End Sub

Private Sub chkEnableAll_Click()
    If chkEnableAll.value <> vbGrayed Then
        chkGalilieo.value = chkEnableAll.value
        chkSabre.value = chkEnableAll.value
        chkWorldspan.value = chkEnableAll.value
        chkAmadeus.value = chkEnableAll.value
    End If
End Sub

Private Sub chkGalilieo_Click()
    If chkGalilieo.value = vbChecked Then
        txtGalilieoSource.Locked = True
        txtGalilieoSource.BackColor = &H8000000F
        txtGalilieoDest.Locked = True
        txtGalilieoDest.BackColor = &H8000000F
        SSTab1.TabPicture(1) = Image3.Picture
    Else
        txtGalilieoSource.Locked = False
        txtGalilieoSource.BackColor = &H80000005
        txtGalilieoDest.Locked = False
        txtGalilieoDest.BackColor = &H80000005
        SSTab1.TabPicture(1) = LoadPicture()
    End If
    dbCompany.Execute "Update [File] set STATUS=" & chkGalilieo.value
    SetEnableAll
End Sub

Private Sub chkSabre_Click()
    If chkSabre.value = vbChecked Then
        txtSource.Locked = True
        txtSource.BackColor = &H8000000F
        txtDest.Locked = True
        txtDest.BackColor = &H8000000F
        SSTab1.TabPicture(2) = Image3.Picture
    Else
        txtSource.Locked = False
        txtSource.BackColor = &H80000005
        txtDest.Locked = False
        txtDest.BackColor = &H80000005
        SSTab1.TabPicture(2) = LoadPicture()
    End If
    dbCompany.Execute "Update [File] set SABRESTATUS=" & chkSabre.value
    SetEnableAll
End Sub

Private Sub chkWorldspan_Click()
    If chkWorldspan.value = vbChecked Then
        txtWorldspanSource.Locked = True
        txtWorldspanSource.BackColor = &H8000000F
        txtWorldspanDest.Locked = True
        txtWorldspanDest.BackColor = &H8000000F
        SSTab1.TabPicture(3) = Image3.Picture
    Else
        txtWorldspanSource.Locked = False
        txtWorldspanSource.BackColor = &H80000005
        txtWorldspanDest.Locked = False
        txtWorldspanDest.BackColor = &H80000005
        SSTab1.TabPicture(3) = LoadPicture()
    End If
    
    dbCompany.Execute "Update [File] set WSPSTATUS=" & chkWorldspan.value
    SetEnableAll
End Sub

Private Sub cmdExit_Click()
    If MsgBox("Do you want to exit?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        End
    End If
End Sub
Public Sub cmdStart_Click()
    cmdStart.Enabled = False
    
    chkEnableAll.Enabled = False
    chkGalilieo.Enabled = False
    chkSabre.Enabled = False
    chkWorldspan.Enabled = False
    chkAmadeus.Enabled = False
    txtGalilieoSource.Locked = True
    txtGalilieoSource.BackColor = &H8000000F
    txtGalilieoDest.Locked = True
    txtGalilieoDest.BackColor = &H8000000F
    txtSource.Locked = True
    txtSource.BackColor = &H8000000F
    txtDest.Locked = True
    txtDest.BackColor = &H8000000F
    txtWorldspanSource.Locked = True
    txtWorldspanSource.BackColor = &H8000000F
    txtWorldspanDest.Locked = True
    txtWorldspanDest.BackColor = &H8000000F
    txtAmadeusSource.Locked = True
    txtAmadeusSource.BackColor = &H8000000F
    txtAmadeusDest.Locked = True
    txtAmadeusDest.BackColor = &H8000000F
    
    tmrUpload.Enabled = True
    cmdStop.Enabled = True
    cmdExit.Enabled = False
    dbCompany.Execute "Update [File] set UPLOADDIRNAME='" & txtGalilieoSource & "',DESTDIRNAME='" & txtGalilieoDest & "', " _
                        & "SABREUPLOADDIRNAME='" & txtSource & "',SABREDESTDIRNAME='" & txtDest & "', " _
                        & "WSPUPLOADDIRNAME='" & txtWorldspanSource & "',WSPDESTDIRNAME='" & txtWorldspanDest & "', " _
                        & "AMDUPLOADDIRNAME='" & txtAmadeusSource & "',AMDDESTDIRNAME='" & txtAmadeusDest & "'"
    
    stbUpload.Panels(1).Text = "Waiting...."
'    If rsSelect.State = 1 Then rsSelect.Close
'    rsSelect.Open "Select UPLOADDIRNAME,DESTDIRNAME,STATUS From File", CnCompany, adOpenDynamic, adLockBatchOptimistic
'    Upload_Status = IIf(IsNull(rsSelect!Status), 0, rsSelect!Status)
'    If Upload_Status = 1 Then
'        With FUpload
'            If rsSelect.EOF = False Then
'
'                .txtDirName = IIf(IsNull(rsSelect!UPLOADDIRNAME), "", rsSelect!UPLOADDIRNAME)
'                .txtRelLocation = IIf(IsNull(rsSelect!DESTDIRNAME), "", rsSelect!DESTDIRNAME)
'
'                txtSource = IIf(IsNull(rsSelect!UPLOADDIRNAME), "", rsSelect!UPLOADDIRNAME)
'                txtDest = IIf(IsNull(rsSelect!DESTDIRNAME), "", rsSelect!DESTDIRNAME)
'
'            End If
'            If .txtDirName = "" Or .txtRelLocation = "" Then
'                Exit Sub
'            End If
'            cnt = .File1.ListCount
'            .File1.Refresh
'            DoEvents
'            If .File1.ListCount > cnt And Upload_Status = 1 Then
'               If FUpload.Visible = False Then
'                 Unload FSabreUpload
'                 Load FUpload
'                 .Show
'                 .Refresh
'               End If
'               .cmdDirUpload_Click
'            ElseIf .File1.ListCount = cnt And .File1.ListCount <> 0 And mir_flag = False Then
'               mir_flag = True 'if files are same allow upload once
'               If FUpload.Visible = False Then
'                 If FSabreUpload.Visible = True Then Unload FSabreUpload
'                 Load FUpload
'                 .Show
'                 .Refresh
'               End If
'               .cmdDirUpload_Click
'            End If
'        End With
'    End If
End Sub

Private Sub cmdStop_Click()
    cmdStop.Enabled = False
    tmrUpload.Enabled = False
    'txtSource = ""
    'txtDest = ""
    cmdStart.Enabled = True
    cmdExit.Enabled = True
    chkGalilieo_Click
    chkSabre_Click
    chkWorldspan_Click
    chkAmadeus_Click
    stbUpload.Panels(1).Text = "Process Stopped."
    
    chkEnableAll.Enabled = True
    chkGalilieo.Enabled = True
    chkSabre.Enabled = True
    chkWorldspan.Enabled = True
    chkAmadeus.Enabled = True
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
    Set TrayArea1.Icon = FMain.Icon
    TrayArea1.Visible = True
    SSTab1.Tab = 0
    'chkEnableAll.Value = vbGrayed
    'chkSabre.Value = vbChecked
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = 1
End If
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Me.Hide
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
    Case 1
        If chkGalilieo.Enabled = True Then chkGalilieo.SetFocus
    Case 2
        If chkSabre.Enabled = True Then chkSabre.SetFocus
    Case 3
        If chkWorldspan.Enabled = True Then chkWorldspan.SetFocus
    Case 4
        If chkAmadeus.Enabled = True Then chkAmadeus.SetFocus
    End Select
End Sub


Private Sub tmrEnd_Timer()
On Error Resume Next
    If UCase(Dir(App.Path & "\EndSQL.txt")) = UCase("EndSQL.txt") And fsObj.FileExists(App.Path & "\_UploadingSQL_") = False Then
        End
    End If
End Sub
Private Sub tmrUpload_Timer()
On Error GoTo errPara
If rsSelect.State = 1 Then rsSelect.Close
rsSelect.Open "Select WSPUPLOADDIRNAME,WSPDESTDIRNAME,WSPSTATUS,AMDUPLOADDIRNAME,AMDDESTDIRNAME,AMDSTATUS,SABREUPLOADDIRNAME,SABREDESTDIRNAME,SABRESTATUS,UPLOADDIRNAME,DESTDIRNAME,STATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic

If chkGalilieo.value = vbChecked Then
    Upload_StatusMIR = IIf(IsNull(rsSelect!Status), 0, rsSelect!Status)
    stbUpload.Panels(1).Text = "Waiting...."
    With FUpload
        If rsSelect.EOF = False Then
            
            .txtDirName = IIf(IsNull(rsSelect!UPLOADDIRNAME), "", rsSelect!UPLOADDIRNAME)
            .txtRelLocation = IIf(IsNull(rsSelect!DESTDIRNAME), "", rsSelect!DESTDIRNAME)
            
            txtGalilieoSource = IIf(IsNull(rsSelect!UPLOADDIRNAME), "", rsSelect!UPLOADDIRNAME)
            txtGalilieoDest = IIf(IsNull(rsSelect!DESTDIRNAME), "", rsSelect!DESTDIRNAME)
            
        End If
        If .txtDirName = "" Or .txtRelLocation = "" Then
            Exit Sub
        End If
        cnt = .File1.ListCount
        .File1.Refresh
        DoEvents
        If .File1.ListCount > cnt And Upload_StatusMIR = 1 Then
           If FUpload.Visible = False Then
             If FSabreUpload.Visible = True Then Unload FSabreUpload
             Load FUpload
             .Refresh
           End If
           .cmdDirUpload_Click
        ElseIf .File1.ListCount = cnt And .File1.ListCount <> 0 And fil_flagMIR = False Then
           fil_flagMIR = True 'if files are same allow upload once
           If FUpload.Visible = False Then
             Unload FSabreUpload
             Load FUpload
             .Refresh
           End If
           .cmdDirUpload_Click
        End If
    End With
End If
If chkSabre.value = vbChecked Then
    'Uploading Sabre Files
    Upload_Status = IIf(IsNull(rsSelect!SABRESTATUS), 0, rsSelect!SABRESTATUS)
    stbUpload.Panels(1).Text = "Waiting...."
    With FSabreUpload
        If rsSelect.EOF = False Then
            
            .txtDirName = IIf(IsNull(rsSelect!SABREUPLOADDIRNAME), "", rsSelect!SABREUPLOADDIRNAME)
            .txtRelLocation = IIf(IsNull(rsSelect!SABREDESTDIRNAME), "", rsSelect!SABREDESTDIRNAME)
            
            txtSource = IIf(IsNull(rsSelect!SABREUPLOADDIRNAME), "", rsSelect!SABREUPLOADDIRNAME)
            txtDest = IIf(IsNull(rsSelect!SABREDESTDIRNAME), "", rsSelect!SABREDESTDIRNAME)
            
        End If
        If .txtDirName = "" Or .txtRelLocation = "" Then
            Exit Sub
        End If
        cnt = .File1.ListCount
        .File1.Refresh
        DoEvents
        If .File1.ListCount > cnt And Upload_Status = 1 Then
           If FSabreUpload.Visible = False Then
             If FUpload.Visible = True Then Unload FUpload
             Load FSabreUpload
             .Refresh
           End If
           .cmdDirUpload_Click
        ElseIf .File1.ListCount = cnt And .File1.ListCount <> 0 And fil_flag = False Then
           fil_flag = True 'if files are same allow upload once
           If FSabreUpload.Visible = False Then
             Unload FUpload
             Load FSabreUpload
             .Refresh
           End If
           .cmdDirUpload_Click
        End If
    End With
End If

If chkWorldspan.value = vbChecked Then
    'Uploading Worldspan
    Upload_StatusWORLD = IIf(IsNull(rsSelect!WSPSTATUS), 0, rsSelect!WSPSTATUS)
    stbUpload.Panels(1).Text = "Waiting...."
    With FWorldSpan
        If rsSelect.EOF = False Then
            
            .txtDirName = IIf(IsNull(rsSelect!WSPUPLOADDIRNAME), "", rsSelect!WSPUPLOADDIRNAME)
            .txtRelLocation = IIf(IsNull(rsSelect!WSPDESTDIRNAME), "", rsSelect!WSPDESTDIRNAME)
            
            txtWorldspanSource = IIf(IsNull(rsSelect!WSPUPLOADDIRNAME), "", rsSelect!WSPUPLOADDIRNAME)
            txtWorldspanDest = IIf(IsNull(rsSelect!WSPDESTDIRNAME), "", rsSelect!WSPDESTDIRNAME)
        End If
                If .txtDirName = "" Or .txtRelLocation = "" Then
            Exit Sub
        End If
        cnt = .File1.ListCount
        .File1.Refresh
        DoEvents
        If .File1.ListCount > cnt And Upload_StatusWORLD = 1 Then
           If FWorldSpan.Visible = False Then
             If FUpload.Visible = True Then Unload FUpload
             If FSabreUpload.Visible = True Then Unload FSabreUpload
             If FAmadeusUpload.Visible = True Then Unload FAmadeusUpload
             Load FWorldSpan
             .Refresh
           End If
           .cmdDirUpload_Click
        ElseIf .File1.ListCount = cnt And .File1.ListCount <> 0 And WSP_flag = False Then
           WSP_flag = True 'if files are same allow upload once
           If FWorldSpan.Visible = False Then
             Unload FUpload
             Unload FSabreUpload
             Unload FAmadeusUpload
             Load FWorldSpan
             .Refresh
           End If
           .cmdDirUpload_Click
        End If

    End With
End If

If chkAmadeus.value = vbChecked Then
    'Uploading Amadeus Files
    Upload_StatusAIR = IIf(IsNull(rsSelect!AMDSTATUS), 0, rsSelect!AMDSTATUS)
    stbUpload.Panels(1).Text = "Waiting...."
    With FAmadeusUpload
        If rsSelect.EOF = False Then
            
            .txtDirName = IIf(IsNull(rsSelect!AMDUPLOADDIRNAME), "", rsSelect!AMDUPLOADDIRNAME)
            .txtRelLocation = IIf(IsNull(rsSelect!AMDDESTDIRNAME), "", rsSelect!AMDDESTDIRNAME)
            
            txtAmadeusSource = IIf(IsNull(rsSelect!AMDUPLOADDIRNAME), "", rsSelect!AMDUPLOADDIRNAME)
            txtAmadeusDest = IIf(IsNull(rsSelect!AMDDESTDIRNAME), "", rsSelect!AMDDESTDIRNAME)
        End If
                If .txtDirName = "" Or .txtRelLocation = "" Then
            Exit Sub
        End If
        cnt = .File1.ListCount
        .File1.Refresh
        DoEvents
        If .File1.ListCount > cnt And Upload_StatusAIR = 1 Then
           If FAmadeusUpload.Visible = False Then
             If FUpload.Visible = True Then Unload FUpload
             If FSabreUpload.Visible = True Then Unload FSabreUpload
             Load FAmadeusUpload
             .Refresh
           End If
           .cmdDirUpload_Click
        ElseIf .File1.ListCount = cnt And .File1.ListCount <> 0 And Air_flag = False Then
           Air_flag = True 'if files are same allow upload once
           If FAmadeusUpload.Visible = False Then
             Unload FUpload
             Unload FSabreUpload
             Load FAmadeusUpload
             .Refresh
           End If
           .cmdDirUpload_Click
        End If

    End With
End If
Exit Sub
errPara:

Dim myErr As ErrObject
Set myErr = Err
If myErr.Number = 76 Then
    cmdStop.Enabled = True
    cmdStop_Click
    stbUpload.Panels(1).Text = "Path Not Found"
Else
    Err.Raise myErr.Number, myErr.Source, myErr.Description
    Resume
End If

End Sub

Function SetEnableAll()
    If chkGalilieo.value = vbChecked And chkSabre.value = vbChecked And chkWorldspan.value = vbChecked And chkAmadeus.value = vbChecked Then
        chkEnableAll.value = vbChecked
    ElseIf chkGalilieo.value = vbUnchecked And chkSabre.value = vbUnchecked And chkWorldspan.value = vbUnchecked And chkAmadeus.value = vbUnchecked Then
        chkEnableAll.value = vbUnchecked
    Else
        chkEnableAll.value = vbGrayed
    End If
End Function

Private Sub TrayArea1_DblClick()
    Me.WindowState = vbNormal
    Me.Show
End Sub

