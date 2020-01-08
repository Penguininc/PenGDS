VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{9F8F4546-0259-44B3-AA22-CFE22545C52F}#1.0#0"; "lvButton.ocx"
Object = "{16920A9E-0846-11D5-8B89-BD3D07939431}#1.0#0"; "TrayArea.ocx"
Begin VB.Form FMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PenGDS Server"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   780
      TabIndex        =   11
      Top             =   2970
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   450
      TabIndex        =   9
      Top             =   3510
      Visible         =   0   'False
      Width           =   795
   End
   Begin lvButton.chameleonButton cmdApply 
      Height          =   525
      Left            =   1920
      TabIndex        =   0
      Top             =   4020
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "Apply"
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
      TabIndex        =   3
      Top             =   4050
      Width           =   1035
   End
   Begin VB.CheckBox chkAutoStart 
      Caption         =   "Auto Start"
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   4320
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4005
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   7064
      _Version        =   393216
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
      Tab(0).Control(2)=   "ImageError"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ImageStarted"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ImageUpdating"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "VersionLabel"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SysInfo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TrayArea1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Picture1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "tmrUpload"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "tmrEnd"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Picture2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "GDS"
      TabPicture(1)   =   "FMain.frx":0F02
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Settings"
      TabPicture(2)   =   "FMain.frx":105C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   150
         Picture         =   "FMain.frx":1078
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   960
      End
      Begin MSComctlLib.ListView ListView 
         Height          =   3195
         Left            =   -74790
         TabIndex        =   10
         Top             =   570
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5636
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Client"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Path"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "WindowName"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Timer tmrEnd 
         Interval        =   1
         Left            =   540
         Top             =   450
      End
      Begin VB.Timer tmrUpload 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   60
         Top             =   450
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   2280
         Picture         =   "FMain.frx":26C2
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   7
         Top             =   1095
         Width           =   960
      End
      Begin SystemTrayControl.TrayArea TrayArea1 
         Left            =   150
         Top             =   765
         _ExtentX        =   900
         _ExtentY        =   900
      End
      Begin SysInfoLib.SysInfo SysInfo 
         Left            =   210
         Top             =   1290
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Label VersionLabel 
         Alignment       =   2  'Center
         Caption         =   "PEN GDS SERVER"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   165
         TabIndex        =   14
         Top             =   3090
         Width           =   5265
      End
      Begin VB.Image ImageUpdating 
         Height          =   240
         Left            =   180
         Picture         =   "FMain.frx":3D0C
         Top             =   3210
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImageStarted 
         Height          =   240
         Left            =   450
         Picture         =   "FMain.frx":3E56
         Top             =   3240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImageError 
         Height          =   240
         Left            =   300
         Picture         =   "FMain.frx":3FA0
         Top             =   3150
         Visible         =   0   'False
         Width           =   240
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
         TabIndex        =   8
         Top             =   915
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "PEN GDS SERVER"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   165
         TabIndex        =   5
         Top             =   2745
         Width           =   5265
      End
   End
   Begin MSComctlLib.StatusBar stbUpload 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   4575
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4895
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4895
         EndProperty
      EndProperty
   End
   Begin lvButton.chameleonButton cmdExit 
      Height          =   525
      Left            =   4710
      TabIndex        =   1
      Top             =   4020
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   926
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
      MICON           =   "FMain.frx":40EA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin lvButton.chameleonButton cmdStop 
      Height          =   525
      Left            =   3780
      TabIndex        =   13
      Top             =   4020
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "Force Stop"
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
      FCOL            =   192
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FMain.frx":4106
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin lvButton.chameleonButton cmdStart 
      Height          =   525
      Left            =   2850
      TabIndex        =   15
      Top             =   4020
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "Force Start"
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
      MICON           =   "FMain.frx":4122
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin lvButton.chameleonButton ShowAllchameleonButton 
      Height          =   255
      Left            =   1110
      TabIndex        =   16
      Top             =   4020
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Show All"
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
      MICON           =   "FMain.frx":413E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin lvButton.chameleonButton HideAllchameleonButton 
      Height          =   255
      Left            =   1110
      TabIndex        =   17
      Top             =   4290
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Hide All"
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
      FCOLO           =   8421504
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FMain.frx":415A
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
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ShowStopAccess As Boolean

'Private Sub chkAutoStart_Click()
Public Sub chkAutoStart_Click()
    INIWrite App.Path & "\PenGDS Server.ini", "General", "AutoStart", chkAutoStart.Value
End Sub

Private Sub chkEnableAll_Click()
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
    If chkEnableAll.Value <> vbGrayed Then
        For i = 1 To ListView.ListItems.Count
            ListView.ListItems(i).Checked = chkEnableAll.Value
            If ListView.ListItems(i).Checked = True Then
                'ListView.ListItems(i).Bold = True
                ListView.ListItems(i).ForeColor = &HC000&
            Else
                'ListView.ListItems(i).Bold = False
                ListView.ListItems(i).ForeColor = &HC0&
            End If
        Next
    End If
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
Exit Sub

PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " (chkEnableAll_Click())"
    Exit Sub
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
End Sub

Public Sub cmdApply_Click()
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
    cmdExit.Enabled = False
    cmdApply.Enabled = False
    cmdStart.Enabled = False
    cmdStop.Enabled = False
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    ShowAllchameleonButton.Enabled = False
    HideAllchameleonButton.Enabled = False
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    DoEvents
    For i = 1 To ListView.ListItems.Count
        With ListView.ListItems(i)
            stbUpload.Panels(2).Text = .Text
            DoEvents
            If .Checked = True Then
                If .ListSubItems(1).Text = "Stopped" Then
                    If Not (UCase(Dir(.ListSubItems(2).Text & "\EndSQL.txt")) = UCase("EndSQL.txt") Or UCase(Dir(.ListSubItems(2).Text & "\EndGDS.txt")) = UCase("EndGDS.txt")) Then
                        Shell """" & .ListSubItems(2).Text & "\PenGDS Interface.exe" & """" & "vbMinimized", vbNormalFocus
                    End If
                End If
            Else
                If .ListSubItems(1).Text = "Started" Then
                    'by Abhi on 18-May-2016 for caseid 6391 PenGDS Server only start client's pengds when we tick Auto Start
                    If chkAutoStart.Value = vbChecked Then
                    'by Abhi on 18-May-2016 for caseid 6391 PenGDS Server only start client's pengds when we tick Auto Start
                        Open (.ListSubItems(2).Text & "\ENDGDS.TXT") For Random As #1
                        Close #1
                        INIPenAIR_String = .ListSubItems(2).Text & "\PenAIR.ini"
                        If UCase(Dir(INIPenAIR_String)) = "" Then
                            INIPenAIR_String = .ListSubItems(2).Text & "\PenSoft.ini"
                        End If
                        While INIRead(INIPenAIR_String, "PenGDS", "GDS", "OFF") = "ON"
                            DoEvents
                        Wend
                        fsObj.DeleteFile (.ListSubItems(2).Text & "\ENDGDS.TXT"), True
                    'by Abhi on 18-May-2016 for caseid 6391 PenGDS Server only start client's pengds when we tick Auto Start
                    End If
                    'by Abhi on 18-May-2016 for caseid 6391 PenGDS Server only start client's pengds when we tick Auto Start
                End If
            End If
        End With
    Next
    For i = 1 To ListView.ListItems.Count
        With ListView.ListItems(i)
            stbUpload.Panels(2).Text = .Text
            DoEvents
            INIWrite App.Path & "\PenGDS Server.ini", "Clients", UCase(.Text), .Checked
        End With
    Next
    stbUpload.Panels(2).Text = ""
    DoEvents
    'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
    If Startup_Boolean = False Then
    'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
        cmdApply.Enabled = True
        cmdExit.Enabled = True
        cmdStart.Enabled = True
        cmdStop.Enabled = True
        'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
        ShowAllchameleonButton.Enabled = True
        HideAllchameleonButton.Enabled = True
        'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
        DoEvents
    'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
    End If
    'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file

'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
Exit Sub

PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " (cmdApply_Click())"
    If Startup_Boolean = False Then
        cmdApply.Enabled = True
        cmdExit.Enabled = True
        cmdStart.Enabled = True
        cmdStop.Enabled = True
        'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
        ShowAllchameleonButton.Enabled = True
        HideAllchameleonButton.Enabled = True
        'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
        DoEvents
    End If
    Exit Sub
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
End Sub

Private Sub cmdExit_Click()
    If MsgBox("Do you want to exit?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        'INIWrite App.Path & "\PenSoft.ini", "PenGDS", "GDS", "OFF"
        End
    End If
End Sub

Private Sub cmdStart_Click()
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
    cmdExit.Enabled = False
    cmdApply.Enabled = False
    cmdStart.Enabled = False
    cmdStop.Enabled = False
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    ShowAllchameleonButton.Enabled = False
    HideAllchameleonButton.Enabled = False
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    DoEvents
    For i = 1 To ListView.ListItems.Count
        With ListView.ListItems(i)
            stbUpload.Panels(2).Text = .Text
            DoEvents
            If .Checked = True Then
                If Not (UCase(Dir(.ListSubItems(2).Text & "\EndSQL.txt")) = UCase("EndSQL.txt") Or UCase(Dir(.ListSubItems(2).Text & "\EndGDS.txt")) = UCase("EndGDS.txt")) Then
                    Shell """" & .ListSubItems(2).Text & "\PenGDS Interface.exe" & """" & "vbMinimized", vbNormalFocus
                End If
                'If .ListSubItems(1).Text = "Started" Then
                    
                    'Open (.ListSubItems(2).Text & "\ENDGDS.TXT") For Random As #1
                    'Close #1
                    'While INIRead(.ListSubItems(2).Text & "\PenSoft.ini", "PenGDS", "GDS", "OFF") = "ON"
                    '    DoEvents
                    'Wend
                    'Sleep 5000
                    'fsObj.DeleteFile (.ListSubItems(2).Text & "\ENDGDS.TXT"), True
                'End If
            End If
        End With
    Next
    For i = 1 To ListView.ListItems.Count
        With ListView.ListItems(i)
            stbUpload.Panels(2).Text = .Text
            DoEvents
            INIWrite App.Path & "\PenGDS Server.ini", "Clients", UCase(.Text), .Checked
        End With
    Next
    stbUpload.Panels(2).Text = ""
    DoEvents
    cmdApply.Enabled = True
    cmdExit.Enabled = True
    cmdStart.Enabled = True
    cmdStop.Enabled = True
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    ShowAllchameleonButton.Enabled = True
    HideAllchameleonButton.Enabled = True
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    DoEvents
    tmrUpload.Enabled = True
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
Exit Sub

PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " (cmdStart_Click())"
    cmdApply.Enabled = True
    cmdExit.Enabled = True
    cmdStart.Enabled = True
    cmdStop.Enabled = True
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    ShowAllchameleonButton.Enabled = True
    HideAllchameleonButton.Enabled = True
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    DoEvents
    Exit Sub
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
End Sub

Private Sub cmdStop_Click()
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
    tmrUpload.Enabled = False
    cmdExit.Enabled = False
    cmdApply.Enabled = False
    cmdStart.Enabled = False
    cmdStop.Enabled = False
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    ShowAllchameleonButton.Enabled = False
    HideAllchameleonButton.Enabled = False
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    DoEvents
    For i = 1 To ListView.ListItems.Count
        With ListView.ListItems(i)
            stbUpload.Panels(2).Text = .Text
            DoEvents
            If .Checked = True Then
                If .ListSubItems(1).Text = "Started" Then
                    If Not (UCase(Dir(.ListSubItems(2).Text & "\EndSQL.txt")) = UCase("EndSQL.txt") Or UCase(Dir(.ListSubItems(2).Text & "\EndGDS.txt")) = UCase("EndGDS.txt")) Then
                        Open (.ListSubItems(2).Text & "\ENDGDS.TXT") For Random As #1
                        Close #1
                        'While INIRead(.ListSubItems(2).Text & "\PenSoft.ini", "PenGDS", "GDS", "OFF") = "ON"
                        '    DoEvents
                        'Wend
                        Sleep 5000
                        fsObj.DeleteFile (.ListSubItems(2).Text & "\ENDGDS.TXT"), True
                    End If
                End If
            End If
        End With
    Next
    For i = 1 To ListView.ListItems.Count
        With ListView.ListItems(i)
            stbUpload.Panels(2).Text = .Text
            DoEvents
            INIWrite App.Path & "\PenGDS Server.ini", "Clients", UCase(.Text), .Checked
        End With
    Next
    stbUpload.Panels(2).Text = ""
    DoEvents
    cmdApply.Enabled = True
    cmdExit.Enabled = True
    cmdStart.Enabled = True
    cmdStop.Enabled = True
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    ShowAllchameleonButton.Enabled = True
    HideAllchameleonButton.Enabled = True
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    DoEvents
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
Exit Sub

PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " (cmdStop_Click())"
    cmdApply.Enabled = True
    cmdExit.Enabled = True
    cmdStart.Enabled = True
    cmdStop.Enabled = True
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    ShowAllchameleonButton.Enabled = True
    HideAllchameleonButton.Enabled = True
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    DoEvents
    Exit Sub
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
End Sub

Private Sub Form_Load()
    Set TrayArea1.Icon = FMain.Icon
    'TrayArea1.Visible = True
    'SSTab1.Tab = 0
    ShowStopAccess = True
    Picture1.Picture = Picture2.Picture
    VersionLabel = App.Major & "." & App.Minor & ".0." & App.Revision
    'chkEnableAll.Value = vbGrayed
    'chkSabre.Value = vbChecked
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = 1
Else
    'INIWrite App.Path & "\PenSoft.ini", "PenGDS", "GDS", "OFF"
End If
End Sub


Private Sub Form_Resize()
    'If Me.WindowState = vbMinimized Then
    '    Me.Hide
    'End If
End Sub

Private Sub HideAllchameleonButton_Click()
'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
    cmdExit.Enabled = False
    cmdApply.Enabled = False
    cmdStart.Enabled = False
    cmdStop.Enabled = False
    ShowAllchameleonButton.Enabled = False
    HideAllchameleonButton.Enabled = False
    DoEvents
    For i = 1 To ListView.ListItems.Count
        With ListView.ListItems(i)
            stbUpload.Panels(2).Text = .Text
            DoEvents
            If .ListSubItems(1).Text = "Started" Then
                CONAME_String = .ListSubItems(3).Text
                If Trim(CONAME_String) <> "" Then
                    Call BringToBackorHide(CONAME_String, False)
                End If
            End If
        End With
    Next
    stbUpload.Panels(2).Text = ""
    DoEvents
    cmdApply.Enabled = True
    cmdExit.Enabled = True
    cmdStart.Enabled = True
    cmdStop.Enabled = True
    ShowAllchameleonButton.Enabled = True
    HideAllchameleonButton.Enabled = True
    DoEvents
Exit Sub

PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " (HideAllchameleonButton())"
    cmdApply.Enabled = True
    cmdExit.Enabled = True
    cmdStart.Enabled = True
    cmdStop.Enabled = True
    ShowAllchameleonButton.Enabled = True
    HideAllchameleonButton.Enabled = True
    DoEvents
    Exit Sub
'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
End Sub

Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    '--set the sortkey to the column header's index - 1
    ListView.SortKey = ColumnHeader.Index - 1
    ListView.Sorted = True
    
    '--toggle the sort order between ascending & descending
    ListView.SortOrder = 1 Xor ListView.SortOrder
End Sub

Private Sub ListView_DblClick()
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
    If ListView.ListItems.Count > 0 Then
        CONAME_String = ListView.SelectedItem.ListSubItems(3).Text
        'MsgBox CONAME_String
        If Trim(CONAME_String) <> "" Then
            Call BringToFront(CONAME_String)
        End If
    End If
    'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
End Sub

Private Sub ListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        'Item.Bold = True
        'Item.ForeColor = &HC000&
        Item.Selected = True
        'Item.
    Else
        'Item.Bold = False
        'Item.ForeColor = vbBlack
        Item.Selected = False
    End If
End Sub


Private Sub ShowAllchameleonButton_Click()
'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
    cmdExit.Enabled = False
    cmdApply.Enabled = False
    cmdStart.Enabled = False
    cmdStop.Enabled = False
    ShowAllchameleonButton.Enabled = False
    HideAllchameleonButton.Enabled = False
    DoEvents
    For i = 1 To ListView.ListItems.Count
        With ListView.ListItems(i)
            stbUpload.Panels(2).Text = .Text
            DoEvents
            If .ListSubItems(1).Text = "Started" Then
                CONAME_String = .ListSubItems(3).Text
                If Trim(CONAME_String) <> "" Then
                    Call BringToFront(CONAME_String, False)
                End If
            End If
        End With
    Next
    stbUpload.Panels(2).Text = ""
    DoEvents
    cmdApply.Enabled = True
    cmdExit.Enabled = True
    cmdStart.Enabled = True
    cmdStop.Enabled = True
    ShowAllchameleonButton.Enabled = True
    HideAllchameleonButton.Enabled = True
    DoEvents
Exit Sub

PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " (ShowAllchameleonButton())"
    cmdApply.Enabled = True
    cmdExit.Enabled = True
    cmdStart.Enabled = True
    cmdStop.Enabled = True
    ShowAllchameleonButton.Enabled = True
    HideAllchameleonButton.Enabled = True
    DoEvents
    Exit Sub
'by Abhi on 17-Jun-2015 for caseid 5325 Option for open PenGDS and show the main window instead of minimize to system tray
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
        Case 1
            'ListView.SetFocus
            ListView.ListItems(1).Selected = False
    End Select
End Sub

Private Sub SysInfo_PowerSuspend()
    tmrUpload.Enabled = False
    End
End Sub

Private Sub tmrEnd_Timer()
On Error Resume Next
    If (UCase(Dir(App.Path & "\EndSQL.txt")) = UCase("EndSQL.txt") Or UCase(Dir(App.Path & "\EndGDS.txt")) = UCase("EndGDS.txt")) And fsObj.FileExists(App.Path & "\_UploadingSQL_") = False Then
        'INIWrite App.Path & "\PenSoft.ini", "PenGDS", "GDS", "OFF"
        End
    End If
End Sub
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
'Private Sub tmrUpload_Timer()
Public Sub tmrUpload_Timer()
On Error GoTo PENErr
Dim PENErr_Number As String, PENErr_Description As String
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
    For iTimer = 1 To ListView.ListItems.Count
        'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
        stbUpload.Panels(1).Text = "Started and Running... " & TimeFormat12HRS(Now)
        'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
        With ListView.ListItems(iTimer)
            INIPenAIR_String = .ListSubItems(2).Text & "\PenAIR.ini"
            If UCase(Dir(INIPenAIR_String)) = "" Then
                INIPenAIR_String = .ListSubItems(2).Text & "\PenSoft.ini"
            End If
            If INIRead(INIPenAIR_String, "PenGDS", "GDS", "OFF") = "ON" Then
                '.Checked = True
                '.Bold = True
                .ForeColor = &HC000&
                .ListSubItems(1).Text = "Started"
                '.ListSubItems(1).Bold = True
                .ListSubItems(1).ForeColor = &HC000&
                SSTab1.TabPicture(1) = ImageStarted.Picture
                If UCase(Dir(.ListSubItems(2).Text & "\EndSQL.txt")) = UCase("EndSQL.txt") Or UCase(Dir(.ListSubItems(2).Text & "\EndGDS.txt")) = UCase("EndGDS.txt") Then
                    .ListSubItems(1).Text = "Updating..."
                End If
            Else
                If .ListSubItems(1).Text = "Started" Then
                    'Email = "TOEMAIL=abhi@penguininc.com,biji@penguininc.com;SUB=GDS Stopped;BODY=Please Check;SMTPUSER=support@penguininc.com;FROMEMAIL=support@penguininc.com;SMTPPWD=penguin1234;SMTPSERVER=mail.penguininc.com;SMTPPORT=21;"
                    'Email = """" & Email & """"
                    'Shell App.Path & "\SendEMail.exe " & Email, vbNormalFocus
                    'Shell """" & App.Path & "\SendEMail.exe" & """ " & Email, vbNormalFocus
                End If
                If .Checked = True And .ListSubItems(1).Text = "Stopped" Then
                    If INIRead(App.Path & "\PenGDS Server.ini", "Clients", UCase(.Text), False) = True Then
                        If Not (UCase(Dir(.ListSubItems(2).Text & "\EndSQL.txt")) = UCase("EndSQL.txt") Or UCase(Dir(.ListSubItems(2).Text & "\EndGDS.txt")) = UCase("EndGDS.txt")) Then
                            'by Abhi on 18-May-2016 for caseid 6391 PenGDS Server only start client's pengds when we tick Auto Start
                            If chkAutoStart.Value = vbChecked Then
                            'by Abhi on 18-May-2016 for caseid 6391 PenGDS Server only start client's pengds when we tick Auto Start
                                Shell """" & .ListSubItems(2).Text & "\PenGDS Interface.exe" & """" & "vbMinimized", vbNormalFocus
                            'by Abhi on 18-May-2016 for caseid 6391 PenGDS Server only start client's pengds when we tick Auto Start
                            End If
                            'by Abhi on 18-May-2016 for caseid 6391 PenGDS Server only start client's pengds when we tick Auto Start
                        Else
                            .ForeColor = &H80FF&
                            .ListSubItems(1).Text = "Updating..."
                            '.ListSubItems(1).Bold = True
                            .ListSubItems(1).ForeColor = &H80FF&
                            SSTab1.TabPicture(1) = ImageUpdating.Picture
                        End If
                    End If
                Else
                    '.Checked = False
                    '.Bold = False
                    .ForeColor = &HC0&
                    .ListSubItems(1).Text = "Stopped"
                    '.ListSubItems(1).Bold = True
                    .ListSubItems(1).ForeColor = &HC0&
                End If
            End If
        End With
    Next
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
Exit Sub

PENErr:
    PENErr_Number = Err.Number
    PENErr_Description = Err.Description
    MsgBox "Error: " & PENErr_Number & vbCrLf & vbCrLf & PENErr_Description, vbCritical, App.Title & " (tmrUpload_Timer())"
    tmrUpload.Enabled = False
    stbUpload.Panels(1).Text = "Process Stopped."
    Exit Sub
'by Abhi on 29-Dec-2014 for caseid 4855 PenGDS Server is not auto started when we update the client and remove the endgds.txt file
End Sub

Function SetEnableAll()
    If chkGalilieo.Value = vbChecked And chkSabre.Value = vbChecked And chkWorldspan.Value = vbChecked And chkAmadeus.Value = vbChecked Then
        chkEnableAll.Value = vbChecked
    ElseIf chkGalilieo.Value = vbUnchecked And chkSabre.Value = vbUnchecked And chkWorldspan.Value = vbUnchecked And chkAmadeus.Value = vbUnchecked Then
        chkEnableAll.Value = vbUnchecked
    Else
        chkEnableAll.Value = vbGrayed
    End If
End Function

Private Sub TrayArea1_DblClick()
    Me.WindowState = vbNormal
    Me.Show
End Sub

