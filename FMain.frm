VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{9F8F4546-0259-44B3-AA22-CFE22545C52F}#1.0#0"; "lvButton.ocx"
Object = "{16920A9E-0846-11D5-8B89-BD3D07939431}#1.0#0"; "TrayArea.ocx"
Begin VB.Form FMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PenGDS Interface"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1200
      TabIndex        =   42
      Top             =   3270
      Visible         =   0   'False
      Width           =   795
   End
   Begin lvButton.chameleonButton cmdStart 
      Height          =   405
      Left            =   4110
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
      MICON           =   "FMain.frx":08CA
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
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   5159
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
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
      TabPicture(0)   =   "FMain.frx":08E6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "VersionLabel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "HelpFunctionalchameleonButton"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "prgUpload"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Picture1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tmrEnd"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tmrUpload"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TrayArea1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Winsock"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Galileo"
      TabPicture(1)   =   "FMain.frx":0902
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDestination(1)"
      Tab(1).Control(1)=   "lblSource(0)"
      Tab(1).Control(2)=   "lblSource(7)"
      Tab(1).Control(3)=   "SearchGalileoCommand"
      Tab(1).Control(4)=   "txtGalilieoDest"
      Tab(1).Control(5)=   "txtGalilieoSource"
      Tab(1).Control(6)=   "chkGalilieo"
      Tab(1).Control(7)=   "Picture2"
      Tab(1).Control(8)=   "txtGalileoExt"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Sabre"
      TabPicture(2)   =   "FMain.frx":091E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblSource(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblDestination(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblSource(6)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "SearchSabreCommand"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chkSabre"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtSource"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtDest"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Picture3"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtSabreExt"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "chkSabreIncludeItineraryOnly"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Worldspan"
      TabPicture(3)   =   "FMain.frx":093A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtWorldspanExt"
      Tab(3).Control(1)=   "Picture4"
      Tab(3).Control(2)=   "txtWorldspanDest"
      Tab(3).Control(3)=   "txtWorldspanSource"
      Tab(3).Control(4)=   "chkWorldspan"
      Tab(3).Control(5)=   "SearchWorldspanCommand"
      Tab(3).Control(6)=   "lblSource(5)"
      Tab(3).Control(7)=   "lblDestination(2)"
      Tab(3).Control(8)=   "lblSource(2)"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Amadeus"
      TabPicture(4)   =   "FMain.frx":0956
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblSource(3)"
      Tab(4).Control(1)=   "lblDestination(3)"
      Tab(4).Control(2)=   "lblSource(4)"
      Tab(4).Control(3)=   "lblSource(10)"
      Tab(4).Control(4)=   "OutOfBookingsDateAmadeusText"
      Tab(4).Control(5)=   "SearchAmadeusCommand"
      Tab(4).Control(6)=   "chkAmadeus"
      Tab(4).Control(7)=   "txtAmadeusSource"
      Tab(4).Control(8)=   "txtAmadeusDest"
      Tab(4).Control(9)=   "Picture5"
      Tab(4).Control(10)=   "txtAmadeusExt"
      Tab(4).Control(11)=   "OutOfBookingsDateAmadeusCommand"
      Tab(4).ControlCount=   12
      TabCaption(5)   =   "Legacy"
      TabPicture(5)   =   "FMain.frx":0972
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblSource(8)"
      Tab(5).Control(1)=   "lblDestination(4)"
      Tab(5).Control(2)=   "lblSource(9)"
      Tab(5).Control(3)=   "SearchLegacyCommand"
      Tab(5).Control(4)=   "LegacyExtText"
      Tab(5).Control(5)=   "Picture6"
      Tab(5).Control(6)=   "LegacyDestinationPathText"
      Tab(5).Control(7)=   "LegacySourcePathText"
      Tab(5).Control(8)=   "LegacyEnableCheck"
      Tab(5).ControlCount=   9
      Begin VB.CommandButton OutOfBookingsDateAmadeusCommand 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   270
         Left            =   -72240
         Picture         =   "FMain.frx":098E
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1560
         Width           =   260
      End
      Begin VB.CheckBox LegacyEnableCheck 
         Alignment       =   1  'Right Justify
         Caption         =   "Enable"
         Height          =   315
         Left            =   -74820
         TabIndex        =   52
         Top             =   720
         Width           =   1365
      End
      Begin VB.TextBox LegacySourcePathText 
         Height          =   330
         Left            =   -73650
         TabIndex        =   53
         Top             =   1920
         Width           =   5115
      End
      Begin VB.TextBox LegacyDestinationPathText 
         Height          =   330
         Left            =   -73650
         TabIndex        =   54
         Top             =   2340
         Width           =   6225
      End
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   -69960
         ScaleHeight     =   690
         ScaleWidth      =   2520
         TabIndex        =   56
         Top             =   550
         Width           =   2520
      End
      Begin VB.TextBox LegacyExtText 
         Height          =   330
         Left            =   -68415
         TabIndex        =   55
         ToolTipText     =   "File Extension Eg: *.xls"
         Top             =   1920
         Width           =   975
      End
      Begin MSWinsockLib.Winsock Winsock 
         Left            =   240
         Top             =   1800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CheckBox chkSabreIncludeItineraryOnly 
         Caption         =   "Include Itinerary Only (DIT ENTRIES)"
         Height          =   405
         Left            =   -73650
         TabIndex        =   18
         Top             =   1440
         Width           =   2025
      End
      Begin VB.TextBox txtAmadeusExt 
         Height          =   330
         Left            =   -68415
         TabIndex        =   35
         ToolTipText     =   "File Extension Eg: *.air;*.txt"
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtWorldspanExt 
         Height          =   330
         Left            =   -68415
         TabIndex        =   27
         ToolTipText     =   "File Extension Eg: *.prt"
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtSabreExt 
         Height          =   330
         Left            =   -68415
         TabIndex        =   21
         ToolTipText     =   "File Extension Eg: *.fil;*.pnr"
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtGalileoExt 
         Height          =   330
         Left            =   -68415
         TabIndex        =   9
         ToolTipText     =   "File Extension Eg: *.mir"
         Top             =   1920
         Width           =   975
      End
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
         Left            =   -69960
         Picture         =   "FMain.frx":0AC0
         ScaleHeight     =   690
         ScaleWidth      =   2520
         TabIndex        =   41
         Top             =   550
         Width           =   2520
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   -69960
         Picture         =   "FMain.frx":1584
         ScaleHeight     =   690
         ScaleWidth      =   2520
         TabIndex        =   40
         Top             =   550
         Width           =   2520
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   -69960
         Picture         =   "FMain.frx":2327
         ScaleHeight     =   690
         ScaleWidth      =   2520
         TabIndex        =   39
         Top             =   550
         Width           =   2520
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   -69960
         Picture         =   "FMain.frx":2829
         ScaleHeight     =   690
         ScaleWidth      =   2520
         TabIndex        =   38
         Top             =   550
         Width           =   2520
      End
      Begin VB.TextBox txtAmadeusDest 
         Height          =   330
         Left            =   -73650
         TabIndex        =   34
         Top             =   2340
         Width           =   6225
      End
      Begin VB.TextBox txtAmadeusSource 
         Height          =   330
         Left            =   -73650
         TabIndex        =   33
         Top             =   1920
         Width           =   5115
      End
      Begin VB.CheckBox chkAmadeus 
         Alignment       =   1  'Right Justify
         Caption         =   "Enable"
         Height          =   315
         Left            =   -74820
         TabIndex        =   30
         Top             =   720
         Width           =   1365
      End
      Begin VB.TextBox txtWorldspanDest 
         Height          =   330
         Left            =   -73650
         TabIndex        =   26
         Top             =   2340
         Width           =   6225
      End
      Begin VB.TextBox txtWorldspanSource 
         Height          =   330
         Left            =   -73650
         TabIndex        =   25
         Top             =   1920
         Width           =   5115
      End
      Begin VB.CheckBox chkWorldspan 
         Alignment       =   1  'Right Justify
         Caption         =   "Enable"
         Height          =   315
         Left            =   -74820
         TabIndex        =   24
         Top             =   720
         Width           =   1365
      End
      Begin VB.TextBox txtDest 
         Height          =   330
         Left            =   -73650
         TabIndex        =   20
         Top             =   2340
         Width           =   6225
      End
      Begin VB.TextBox txtSource 
         Height          =   330
         Left            =   -73650
         TabIndex        =   19
         Top             =   1920
         Width           =   5115
      End
      Begin VB.CheckBox chkSabre 
         Alignment       =   1  'Right Justify
         Caption         =   "Enable"
         Height          =   315
         Left            =   -74820
         TabIndex        =   17
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
         Picture         =   "FMain.frx":3719
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   14
         Top             =   570
         Width           =   960
      End
      Begin VB.TextBox txtGalilieoSource 
         Height          =   330
         Left            =   -73650
         TabIndex        =   7
         Top             =   1920
         Width           =   5115
      End
      Begin VB.TextBox txtGalilieoDest 
         Height          =   330
         Left            =   -73650
         TabIndex        =   8
         Top             =   2340
         Width           =   6225
      End
      Begin MSComctlLib.ProgressBar prgUpload 
         Height          =   375
         Left            =   180
         TabIndex        =   15
         Top             =   2310
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin lvButton.chameleonButton SearchGalileoCommand 
         Height          =   285
         Left            =   -68415
         TabIndex        =   43
         ToolTipText     =   "PenSEARCH"
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "FMain.frx":4D63
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin lvButton.chameleonButton SearchSabreCommand 
         Height          =   285
         Left            =   -68415
         TabIndex        =   44
         ToolTipText     =   "PenSEARCH"
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "FMain.frx":4D7F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin lvButton.chameleonButton SearchWorldspanCommand 
         Height          =   285
         Left            =   -68415
         TabIndex        =   45
         ToolTipText     =   "PenSEARCH"
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "FMain.frx":4D9B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin lvButton.chameleonButton SearchAmadeusCommand 
         Height          =   285
         Left            =   -68415
         TabIndex        =   46
         ToolTipText     =   "PenSEARCH"
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "FMain.frx":4DB7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin lvButton.chameleonButton SearchLegacyCommand 
         Height          =   285
         Left            =   -68415
         TabIndex        =   57
         ToolTipText     =   "PenSEARCH"
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "FMain.frx":4DD3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin lvButton.chameleonButton HelpFunctionalchameleonButton 
         Height          =   300
         Left            =   7020
         TabIndex        =   61
         ToolTipText     =   "Functional Help"
         Top             =   720
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         BTYPE           =   13
         TX              =   ""
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FMain.frx":4DEF
         PICN            =   "FMain.frx":4E0B
         PICH            =   "FMain.frx":53A5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox OutOfBookingsDateAmadeusText 
         Height          =   310
         Left            =   -73650
         TabIndex        =   31
         ToolTipText     =   "Out Of Bookings Date"
         Top             =   1530
         Width           =   1695
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "Out of Bkgs Dt"
         Height          =   195
         Index           =   10
         Left            =   -74790
         TabIndex        =   62
         ToolTipText     =   "Out Of Bookings Date"
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   9
         Left            =   -74790
         TabIndex        =   60
         Top             =   1980
         Width           =   510
      End
      Begin VB.Label lblDestination 
         AutoSize        =   -1  'True
         Caption         =   "Destination"
         Height          =   195
         Index           =   4
         Left            =   -74790
         TabIndex        =   59
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label lblSource 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ver 1"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   8
         Left            =   -67815
         TabIndex        =   58
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label VersionLabel 
         Alignment       =   2  'Center
         Caption         =   "VERSION"
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
         TabIndex        =   51
         Top             =   2025
         Width           =   6495
      End
      Begin VB.Label lblSource 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Spec:2010-4"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   7
         Left            =   -68355
         TabIndex        =   50
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lblSource 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Revision 17, Version 25"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   6
         Left            =   -69120
         TabIndex        =   49
         Top             =   1320
         Width           =   1680
      End
      Begin VB.Label lblSource 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Level 20"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   5
         Left            =   -68055
         TabIndex        =   48
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblSource 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 207"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   4
         Left            =   -68280
         TabIndex        =   47
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label lblDestination 
         AutoSize        =   -1  'True
         Caption         =   "Destination"
         Height          =   195
         Index           =   3
         Left            =   -74790
         TabIndex        =   37
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   3
         Left            =   -74790
         TabIndex        =   36
         Top             =   1980
         Width           =   510
      End
      Begin VB.Label lblDestination 
         AutoSize        =   -1  'True
         Caption         =   "Destination"
         Height          =   195
         Index           =   2
         Left            =   -74790
         TabIndex        =   29
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   2
         Left            =   -74790
         TabIndex        =   28
         Top             =   1980
         Width           =   510
      End
      Begin VB.Label lblDestination 
         AutoSize        =   -1  'True
         Caption         =   "Destination"
         Height          =   195
         Index           =   0
         Left            =   -74790
         TabIndex        =   23
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   1
         Left            =   -74790
         TabIndex        =   22
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
         TabIndex        =   16
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
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   165
         TabIndex        =   10
         Top             =   1680
         Width           =   6495
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   0
         Left            =   -74790
         TabIndex        =   13
         Top             =   1980
         Width           =   510
      End
      Begin VB.Label lblDestination 
         AutoSize        =   -1  'True
         Caption         =   "Destination"
         Height          =   195
         Index           =   1
         Left            =   -74790
         TabIndex        =   12
         Top             =   2400
         Width           =   795
      End
   End
   Begin MSComctlLib.StatusBar stbUpload 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   3540
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5822
            MinWidth        =   5822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "GDS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5503
            Key             =   "FILENAME"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "SOURCEPATH"
         EndProperty
      EndProperty
   End
   Begin lvButton.chameleonButton cmdStop 
      Height          =   405
      Left            =   5400
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
      MICON           =   "FMain.frx":593F
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
      Left            =   6690
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
      MICON           =   "FMain.frx":595B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   2730
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image ImageError 
      Height          =   240
      Left            =   1980
      Picture         =   "FMain.frx":5977
      Top             =   2970
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageOK 
      Height          =   240
      Left            =   1560
      Picture         =   "FMain.frx":5AC1
      Top             =   2970
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1170
      Picture         =   "FMain.frx":5C0B
      Top             =   2970
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
Dim Upload_StatusLegacy As Integer

Dim ShowStopAccess As Boolean
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
'Dim NoofPermissionDenied As Long
'Dim PermissionDenied As Long
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
'Dim fForcedStop_Boolean As Boolean
Public fForcedStop_Boolean As Boolean
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available

Private Sub chkAmadeus_Click()
    If chkAmadeus.value = vbChecked Then
        txtAmadeusSource.Locked = True
        txtAmadeusSource.BackColor = &H8000000F
        txtAmadeusDest.Locked = True
        txtAmadeusDest.BackColor = &H8000000F
        txtAmadeusExt.Locked = True
        txtAmadeusExt.BackColor = &H8000000F
        'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
        OutOfBookingsDateAmadeusText.Locked = True
        OutOfBookingsDateAmadeusText.BackColor = &H8000000F
        OutOfBookingsDateAmadeusCommand.Enabled = False
        'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'SSTab1.TabPicture(4) = Image3.Picture
        If SSTab1.TabPicture(4) <> ImageError.Picture Then
            SSTab1.TabPicture(4) = Image3.Picture
        End If
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    Else
        txtAmadeusSource.Locked = False
        txtAmadeusSource.BackColor = &H80000005
        txtAmadeusDest.Locked = False
        txtAmadeusDest.BackColor = &H80000005
        txtAmadeusExt.Locked = False
        txtAmadeusExt.BackColor = &H80000005
        'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
        OutOfBookingsDateAmadeusText.Locked = False
        OutOfBookingsDateAmadeusText.BackColor = &H80000005
        OutOfBookingsDateAmadeusCommand.Enabled = True
        'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
        SSTab1.TabPicture(4) = LoadPicture()
    End If
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'dbCompany.Execute "Update [File] set AMDSTATUS=" & chkAmadeus.value
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
    SetEnableAll
End Sub

Private Sub chkAutoStart_Click()
    dbCompany.Execute "Update [File] set AUTOSTART=" & chkAutoStart.value
    'by Abhi on 16-Feb-2016 for caseid 6044 PenGDS and GDSAuto Monitoring in PenAIR if enabled show the last date and time
    If chkAutoStart.value = vbChecked And (chkEnableAll.value = vbChecked Or chkEnableAll.value = vbGrayed) Then
        INIWrite INIPenAIR_String, "PenGDS", "PenGDSACTIVE", Now & " STOP"
    Else
        INIWrite INIPenAIR_String, "PenGDS", "PenGDSACTIVE", ""
    End If
    'by Abhi on 16-Feb-2016 for caseid 6044 PenGDS and GDSAuto Monitoring in PenAIR if enabled show the last date and time
End Sub

Private Sub chkEnableAll_Click()
    If chkEnableAll.value <> vbGrayed Then
        chkGalilieo.value = chkEnableAll.value
        chkSabre.value = chkEnableAll.value
        chkWorldspan.value = chkEnableAll.value
        chkAmadeus.value = chkEnableAll.value
        'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
        LegacyEnableCheck.value = chkEnableAll.value
    End If
End Sub

Private Sub chkGalilieo_Click()
    If chkGalilieo.value = vbChecked Then
        txtGalilieoSource.Locked = True
        txtGalilieoSource.BackColor = &H8000000F
        txtGalilieoDest.Locked = True
        txtGalilieoDest.BackColor = &H8000000F
        txtGalileoExt.Locked = True
        txtGalileoExt.BackColor = &H8000000F
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'SSTab1.TabPicture(1) = Image3.Picture
        If SSTab1.TabPicture(1) <> ImageError.Picture Then
            SSTab1.TabPicture(1) = Image3.Picture
        End If
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    Else
        txtGalilieoSource.Locked = False
        txtGalilieoSource.BackColor = &H80000005
        txtGalilieoDest.Locked = False
        txtGalilieoDest.BackColor = &H80000005
        txtGalileoExt.Locked = False
        txtGalileoExt.BackColor = &H80000005
        SSTab1.TabPicture(1) = LoadPicture()
    End If
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'dbCompany.Execute "Update [File] set STATUS=" & chkGalilieo.value
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
    SetEnableAll
End Sub

Private Sub chkSabre_Click()
    If chkSabre.value = vbChecked Then
        txtSource.Locked = True
        txtSource.BackColor = &H8000000F
        txtDest.Locked = True
        txtDest.BackColor = &H8000000F
        txtSabreExt.Locked = True
        txtSabreExt.BackColor = &H8000000F
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'SSTab1.TabPicture(2) = Image3.Picture
        If SSTab1.TabPicture(2) <> ImageError.Picture Then
            SSTab1.TabPicture(2) = Image3.Picture
        End If
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    Else
        txtSource.Locked = False
        txtSource.BackColor = &H80000005
        txtDest.Locked = False
        txtDest.BackColor = &H80000005
        txtSabreExt.Locked = False
        txtSabreExt.BackColor = &H80000005
        SSTab1.TabPicture(2) = LoadPicture()
    End If
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'dbCompany.Execute "Update [File] set SABRESTATUS=" & chkSabre.value
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
    SetEnableAll
End Sub

Private Sub chkSabreIncludeItineraryOnly_Click()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub chkWorldspan_Click()
    If chkWorldspan.value = vbChecked Then
        txtWorldspanSource.Locked = True
        txtWorldspanSource.BackColor = &H8000000F
        txtWorldspanDest.Locked = True
        txtWorldspanDest.BackColor = &H8000000F
        txtWorldspanExt.Locked = True
        txtWorldspanExt.BackColor = &H8000000F
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'SSTab1.TabPicture(3) = Image3.Picture
        If SSTab1.TabPicture(3) <> ImageError.Picture Then
            SSTab1.TabPicture(3) = Image3.Picture
        End If
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    Else
        txtWorldspanSource.Locked = False
        txtWorldspanSource.BackColor = &H80000005
        txtWorldspanDest.Locked = False
        txtWorldspanDest.BackColor = &H80000005
        txtWorldspanExt.Locked = False
        txtWorldspanExt.BackColor = &H80000005
        SSTab1.TabPicture(3) = LoadPicture()
    End If
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'dbCompany.Execute "Update [File] set WSPSTATUS=" & chkWorldspan.value
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
    SetEnableAll
End Sub

Private Sub cmdExit_Click()
    If Trim(cmdStart.Tag) = "Stop" Then Exit Sub
    If MsgBox("Do you want to exit?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        If SendERROREmailPreviousMess_String <> "" Then
            SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & "SendERROREmailBeforeExit" & " - " & "SendERROREmailBeforeExit" & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & "). " & App.ProductName & " is Exit.", "SendERROREmailBeforeExit"
        End If
        INIWrite INIPenAIR_String, "PenGDS", "GDS", "OFF"
        End
    End If
End Sub
Public Sub cmdStart_Click()
On Error Resume Next
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
Dim vDefaultSMTP_String As String
Dim Query As String
Dim Q As New QueryString
Dim mDQ As Boolean
'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
    
    If Trim(cmdStart.Tag) = "Stop" Then Exit Sub
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    If Trim(OutOfBookingsDateAmadeusText.Text) <> "" Then
        OutOfBookingsDateAmadeusText.Text = DateFormat(OutOfBookingsDateAmadeusText.Text)
        If mFnCheckValidDate(OutOfBookingsDateAmadeusText.Text, lblSource(10).Caption) = False Then
            mSetFocus OutOfBookingsDateAmadeusText
            Exit Sub
        End If
    End If
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    
    NoofPermissionDenied = 0
    cmdStart.Enabled = False
    
    chkEnableAll.Enabled = False
    chkGalilieo.Enabled = False
    chkSabre.Enabled = False
    chkWorldspan.Enabled = False
    chkAmadeus.Enabled = False
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    LegacyEnableCheck.Enabled = False
    chkSabreIncludeItineraryOnly.Enabled = False
    txtGalilieoSource.Locked = True
    txtGalilieoSource.BackColor = &H8000000F
    txtGalilieoDest.Locked = True
    txtGalilieoDest.BackColor = &H8000000F
    txtGalileoExt.Locked = True
    txtGalileoExt.BackColor = &H8000000F
    txtSource.Locked = True
    txtSource.BackColor = &H8000000F
    txtDest.Locked = True
    txtDest.BackColor = &H8000000F
    txtSabreExt.Locked = True
    txtSabreExt.BackColor = &H8000000F
    txtWorldspanSource.Locked = True
    txtWorldspanSource.BackColor = &H8000000F
    txtWorldspanDest.Locked = True
    txtWorldspanDest.BackColor = &H8000000F
    txtWorldspanExt.Locked = True
    txtWorldspanExt.BackColor = &H8000000F
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    OutOfBookingsDateAmadeusText.Locked = True
    OutOfBookingsDateAmadeusText.BackColor = &H8000000F
    OutOfBookingsDateAmadeusCommand.Enabled = False
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    txtAmadeusSource.Locked = True
    txtAmadeusSource.BackColor = &H8000000F
    txtAmadeusDest.Locked = True
    txtAmadeusDest.BackColor = &H8000000F
    txtAmadeusExt.Locked = True
    txtAmadeusExt.BackColor = &H8000000F
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    LegacySourcePathText.Locked = True
    LegacySourcePathText.BackColor = &H8000000F
    LegacyDestinationPathText.Locked = True
    LegacyDestinationPathText.BackColor = &H8000000F
    LegacyExtText.Locked = True
    LegacyExtText.BackColor = &H8000000F
    
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    fForcedStop_Boolean = False
    'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    tmrUpload.Enabled = True
    cmdStop.Enabled = True
    cmdExit.Enabled = False
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    If AnyDataChanged_Boolean = True Then
        'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
        'dbCompany.Execute "Update [File] set UPLOADDIRNAME='" & txtGalilieoSource & "',DESTDIRNAME='" & txtGalilieoDest & "', " _
                            & "SABREUPLOADDIRNAME='" & txtSource & "',SABREDESTDIRNAME='" & txtDest & "', " _
                            & "WSPUPLOADDIRNAME='" & txtWorldspanSource & "',WSPDESTDIRNAME='" & txtWorldspanDest & "', " _
                            & "AMDUPLOADDIRNAME='" & txtAmadeusSource & "',AMDDESTDIRNAME='" & txtAmadeusDest & "'"
        dbCompany.Execute "Update [File] set UPLOADDIRNAME='" & txtGalilieoSource & "',DESTDIRNAME='" & txtGalilieoDest & "', " _
                            & "SABREUPLOADDIRNAME='" & txtSource & "',SABREDESTDIRNAME='" & txtDest & "', " _
                            & "WSPUPLOADDIRNAME='" & txtWorldspanSource & "',WSPDESTDIRNAME='" & txtWorldspanDest & "', " _
                            & "AMDUPLOADDIRNAME='" & txtAmadeusSource & "',AMDDESTDIRNAME='" & txtAmadeusDest & "', " _
                            & "STATUS=" & chkGalilieo.value & ",SABRESTATUS=" & chkSabre.value & ",WSPSTATUS=" & chkWorldspan.value & ",AMDSTATUS=" & chkAmadeus.value & ", " _
                            & "LegacyEnable=" & LegacyEnableCheck.value & ",LegacySourcePath='" & LegacySourcePathText & "',LegacyDestinationPath='" & LegacyDestinationPathText & "', " _
                            & "OutOfBookingsDateAmadeus='" & DateFormatBlankto1900(OutOfBookingsDateAmadeusText.Text) & "'"
        INIWrite INIPenAIR_String, "PenGDS", "GalileoExt", FMain.txtGalileoExt
        INIWrite INIPenAIR_String, "PenGDS", "SabreExt", FMain.txtSabreExt
        INIWrite INIPenAIR_String, "PenGDS", "SabreIncludeItineraryOnly", FMain.chkSabreIncludeItineraryOnly.value
        INIWrite INIPenAIR_String, "PenGDS", "WorldspanExt", FMain.txtWorldspanExt
        INIWrite INIPenAIR_String, "PenGDS", "AmadeusExt", FMain.txtAmadeusExt
        'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
        INIWrite INIPenAIR_String, "PenGDS", "LegacyExt", FMain.LegacyExtText
        
        'PenSearch
        INIWrite App.Path & "\PenSEARCH.ini", "Galileo", "Source", txtGalilieoSource
        INIWrite App.Path & "\PenSEARCH.ini", "Galileo", "Target", txtGalilieoDest
        INIWrite App.Path & "\PenSEARCH.ini", "Galileo", "Ext", txtGalileoExt
        INIWrite App.Path & "\PenSEARCH.ini", "Sabre", "Source", txtSource
        INIWrite App.Path & "\PenSEARCH.ini", "Sabre", "Target", txtDest
        INIWrite App.Path & "\PenSEARCH.ini", "Sabre", "Ext", txtSabreExt
        INIWrite App.Path & "\PenSEARCH.ini", "Worldspan", "Source", txtWorldspanSource
        INIWrite App.Path & "\PenSEARCH.ini", "Worldspan", "Target", txtWorldspanDest
        INIWrite App.Path & "\PenSEARCH.ini", "Worldspan", "Ext", txtWorldspanExt
        INIWrite App.Path & "\PenSEARCH.ini", "Amadeus", "Source", txtAmadeusSource
        INIWrite App.Path & "\PenSEARCH.ini", "Amadeus", "Target", txtAmadeusDest
        INIWrite App.Path & "\PenSEARCH.ini", "Amadeus", "Ext", txtAmadeusExt
        'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
        INIWrite App.Path & "\PenSEARCH.ini", "Legacy", "Source", LegacySourcePathText
        INIWrite App.Path & "\PenSEARCH.ini", "Legacy", "Target", LegacyDestinationPathText
        INIWrite App.Path & "\PenSEARCH.ini", "Legacy", "Ext", LegacyExtText
        
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        'by Abhi on 25-Mar-2013 for caseid 2786 PenGDS Sabre PNR missing - SabreM1 row not found. we need to skip this file
        If Trim(FMain.txtSource) <> "" Then
            If PathExists(FMain.txtSource) = True Then
                'If Dir(FMain.txtSource & "\Error M1 Missing\", vbDirectory) = "" Then
                If PathExists(FMain.txtSource & "\Error M1 Missing\") = False Then
                    MkDir FMain.txtSource & "\Error M1 Missing\"
                End If
            End If
        End If
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        
        'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
        If Trim(FMain.OutOfBookingsDateAmadeusText.Text) <> "" Then
            If Trim(FMain.txtAmadeusSource) <> "" Then
                If PathExists(FMain.txtAmadeusSource) = True Then
                    If PathExists(FMain.txtAmadeusSource & "\Outofbookings\") = False Then
                        MkDir FMain.txtAmadeusSource & "\Outofbookings\"
                    End If
                End If
            End If
        End If
        'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
        
        'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        AnyDataChanged_Boolean = False
    End If
    
    'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
    'SMTPUSER=noreply@penguininc.com;SMTPPWD=s@pth@m@s!;SMTPSERVER=mail.penguininc.com;SMTPPORT=587;SMTPSSL=1;
    vDefaultSMTP_String = HelpServiceGetServiceURL("DefaultSMTP", False)
    vDefaultSMTP_String = Replace(vDefaultSMTP_String, "http://", "")
    'vDefaultSMTP_String = "SMTPUSER=noreply@penguininc.com;SMTPPWD=s@pth@m@s!;SMTPSERVER=mail.penguininc.com;SMTPPORT=587;SMTPSSL=1;"
    Query = vDefaultSMTP_String
    If Len(Query) > 0 Then
        Q.Query = Query
        If Len(Q("Open")) > 0 Then
            mDQ = True
        End If
    End If
    If Trim(Query) <> "" Then
        If Q("SMTPUSER") <> "" Then
            INIWrite INIPenAIR_String, "DefaultSMTP", "SMTPUSER", Q("SMTPUSER")
        End If
        If Q("SMTPPWD") <> "" Then
            INIWrite INIPenAIR_String, "DefaultSMTP", "SMTPPWD", Q("SMTPPWD")
        End If
        If Q("SMTPSERVER") <> "" Then
            INIWrite INIPenAIR_String, "DefaultSMTP", "SMTPSERVER", Q("SMTPSERVER")
        End If
        If Q("SMTPPORT") <> "" Then
            INIWrite INIPenAIR_String, "DefaultSMTP", "SMTPPORT", Q("SMTPPORT")
        End If
        If Q("SMTPSSL") <> "" Then
            INIWrite INIPenAIR_String, "DefaultSMTP", "SMTPSSL", Q("SMTPSSL")
        End If
    End If
    'by Abhi on 02-Oct-2017 for caseid 7924 noreply@penguininc.com DefaultSMTP as Service enable for email alerts in PenAIR and PenGDS
    
    'by Abhi on 16-Mar-2012 for caseid 2151 change waiting to Active - Waiting for GDS files... for pengds
    stbUpload.Panels(1).Text = "Active - Waiting for GDS files..."
    SendStatus "Started"
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

Public Sub cmdStop_Click()
On Error Resume Next
    If ShowStopAccess = True Then
        If UCase(INIRead(INIPenAIR_String, "PenGDS", "StopAccess", "No")) = UCase("Yes") Then
            frmLogin.Show 1
            If mLogin = False Then Exit Sub
        End If
    End If
    ShowStopAccess = True
    cmdStop.Enabled = False
    'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    fForcedStop_Boolean = True
    'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
    tmrUpload.Enabled = False
    'txtSource = ""
    'txtDest = ""
    cmdStart.Enabled = True
    cmdExit.Enabled = True
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'chkGalilieo_Click
    'chkSabre_Click
    'chkWorldspan_Click
    'chkAmadeus_Click
    ''by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    'LegacyEnableCheck_Click
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    
    stbUpload.Panels(1).Text = "Process Stopped."
    
    chkEnableAll.Enabled = True
    chkGalilieo.Enabled = True
    chkSabre.Enabled = True
    chkWorldspan.Enabled = True
    chkAmadeus.Enabled = True
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    LegacyEnableCheck.Enabled = True
    
    'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    Call chkGalilieo_Click
    Call chkSabre_Click
    Call chkWorldspan_Click
    Call chkAmadeus_Click
    Call LegacyEnableCheck_Click
    'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    
    chkSabreIncludeItineraryOnly.Enabled = True
    fsObj.DeleteFile (App.Path & "\_UploadingSQL_"), True
    SendStatus "Stopped"
    'by Abhi on 12-Mar-2012 for caseid 1551 PenAIR Monitor new status added
    'by Abhi on 16-Feb-2016 for caseid 6044 PenGDS and GDSAuto Monitoring in PenAIR if enabled show the last date and time
    'INIWrite INIPenAIR_String, "PenGDS", "PenGDSACTIVE", ""
    If chkAutoStart.value = vbChecked And (chkEnableAll.value = vbChecked Or chkEnableAll.value = vbGrayed) Then
        INIWrite INIPenAIR_String, "PenGDS", "PenGDSACTIVE", Now & " STOP"
    Else
        INIWrite INIPenAIR_String, "PenGDS", "PenGDSACTIVE", ""
    End If
    'by Abhi on 16-Feb-2016 for caseid 6044 PenGDS and GDSAuto Monitoring in PenAIR if enabled show the last date and time
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    AnyDataChanged_Boolean = False
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
End Sub

Private Sub Form_Load()
    Set TrayArea1.Icon = FMain.Icon
    TrayArea1.Visible = True
    SSTab1.Tab = 0
    DoEvents
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    Picture1.Left = (SSTab1.Width / 2) - (Picture1.Width / 2)
    Label1.Width = SSTab1.Width - (Label1.Left + Label1.Left)
    VersionLabel.Width = SSTab1.Width - (VersionLabel.Left + VersionLabel.Left)
    prgUpload.Width = SSTab1.Width - (prgUpload.Left + prgUpload.Left)
    ShowStopAccess = True
    VersionLabel = App.Major & "." & App.Minor & ".0." & App.Revision
    'chkEnableAll.Value = vbGrayed
    'chkSabre.Value = vbChecked
    'If Winsock.State <> sckConnected And Monitor = True Then
        Winsock.Close
        Winsock.Connect Host, Port
    'End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = SHIFT_MASK + CTRL_MASK + ALT_MASK Then
        If Button = vbRightButton Then
            Call ShowFormName(Me)
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Me.Hide
    Cancel = 1
Else
    INIWrite INIPenAIR_String, "PenGDS", "GDS", "OFF"
End If
End Sub


Private Sub Form_Resize()
'    If Me.WindowState = vbMinimized Then
'        Me.Hide
'    End If
    HelpFunctionalchameleonButton.Top = ScaleTop + SSTab1.TabHeight + 30
    HelpFunctionalchameleonButton.Left = ScaleWidth - HelpFunctionalchameleonButton.Width - 60
    HelpFunctionalchameleonButton.ZOrder 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    INIWrite INIPenAIR_String, "PenGDS", "GDS", "OFF"
End Sub

Private Sub HelpFunctionalchameleonButton_Click()
    Call HelpService(Me, HelpServiceURLType_Functional)
End Sub

Private Sub HelpFunctionalchameleonButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = vbShiftMask + vbCtrlMask + vbAltMask Then
        If Button = vbRightButton Then
            Call HelpService(Me, HelpServiceURLType_Technical)
        End If
    End If
End Sub

Private Sub LegacyDestinationPathText_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub LegacyDestinationPathText_GotFocus()
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    SelectText LegacyDestinationPathText
End Sub

Private Sub LegacyEnableCheck_Click()
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    If LegacyEnableCheck.value = vbChecked Then
        LegacySourcePathText.Locked = True
        LegacySourcePathText.BackColor = &H8000000F
        LegacyDestinationPathText.Locked = True
        LegacyDestinationPathText.BackColor = &H8000000F
        LegacyExtText.Locked = True
        LegacyExtText.BackColor = &H8000000F
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'SSTab1.TabPicture(5) = Image3.Picture
        'by Abhi on 06-Sep-2017 for caseid 7766 Penline and Company currency change in penair is not reflecting in pengds
        If SSTab1.TabPicture(5) <> ImageError.Picture Then
            SSTab1.TabPicture(5) = Image3.Picture
        End If
        'by Abhi on 06-Sep-2017 for caseid 7766 Penline and Company currency change in penair is not reflecting in pengds
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    Else
        LegacySourcePathText.Locked = False
        LegacySourcePathText.BackColor = &H80000005
        LegacyDestinationPathText.Locked = False
        LegacyDestinationPathText.BackColor = &H80000005
        LegacyExtText.Locked = False
        LegacyExtText.BackColor = &H80000005
        SSTab1.TabPicture(5) = LoadPicture()
    End If
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'dbCompany.Execute "Update [File] set LegacyEnable=" & LegacyEnableCheck.value
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
    SetEnableAll
End Sub

Private Sub LegacyExtText_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub LegacyExtText_GotFocus()
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    SelectText LegacyExtText
End Sub

Private Sub LegacySourcePathText_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub LegacySourcePathText_GotFocus()
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    SelectText LegacySourcePathText
End Sub

Private Sub OutOfBookingsDateAmadeusCommand_Click()
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    If OutOfBookingsDateAmadeusText.Locked = True Then Exit Sub
    Call ELcls.SetMonthView(Me.Top + 1000, Me.Left + 3200, Me, OutOfBookingsDateAmadeusText.Name, OutOfBookingsDateAmadeusText.Text)
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
End Sub

Private Sub OutOfBookingsDateAmadeusText_Change()
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    AnyDataChanged_Boolean = True
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
End Sub

Private Sub OutOfBookingsDateAmadeusText_GotFocus()
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    SelectText OutOfBookingsDateAmadeusText
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = SHIFT_MASK + CTRL_MASK + ALT_MASK Then
        If Button = vbRightButton Then
            FSettings.Show
        End If
    End If
End Sub

Private Sub SearchAmadeusCommand_Click()
    Call CallPenSEARCH("Amadeus")
End Sub

Private Sub SearchGalileoCommand_Click()
    Call CallPenSEARCH("Galileo")
End Sub

Private Sub SearchLegacyCommand_Click()
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    Call CallPenSEARCH("Legacy")
End Sub

Private Sub SearchSabreCommand_Click()
    Call CallPenSEARCH("Sabre")
End Sub

Private Sub SearchWorldspanCommand_Click()
    Call CallPenSEARCH("Worldspan")
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
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    Case 5
        If LegacyEnableCheck.Enabled = True Then LegacyEnableCheck.SetFocus
    End Select
End Sub


Private Sub SSTab1_DblClick()
    'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    If SSTab1.Tab = 0 Then
        '########################################################################################################################
        'SendERROREmailPreviousMess_String = "" _
        '    & "12/Aug/2016 10:47:21 AM(10:47:21) : Error: -2147467259 - Data provider or other service returned an E_FAIL status." & vbCrLf _
        '    & "MSGTAG=KFT, Table=AmdLineKFT, Field=PSGRID, Field does not exist in Collection!, " & vbCrLf _
        '    & "(Amadeus - AIR00002_20160811125358.Txt) Moved to folder 'Error Files'. PenGDS is automatically Resumed." '_
        '    '& vbCrLf _
        '    '& vbCrLf _
        '    '& "12/Aug/2016 10:47:21 AM(10:47:21) : Error: -2147467259 - Data provider or other service returned an E_FAIL status." & vbCrLf _
        '    '& "MSGTAG=KFT, Table=AmdLineKFT, Field=PSGRID, Field does not exist in Collection!, " & vbCrLf _
        '    '& "(Amadeus - AIR00002_20160811125358.Txt) Moved to folder 'Error Files'. PenGDS is automatically Resumed." _
        '    '& vbCrLf _
        '    '& vbCrLf _
        '    '& "12/Aug/2016 10:47:21 AM(10:47:21) : Error: -2147467259 - Data provider or other service returned an E_FAIL status." & vbCrLf _
        '    '& "MSGTAG=KFT, Table=AmdLineKFT, Field=PSGRID, Field does not exist in Collection!, " & vbCrLf _
        '    '& "(Amadeus - AIR00002_20160811125358.Txt) Moved to folder 'Error Files'. PenGDS is automatically Resumed." _
        '    '& vbCrLf _
        '    '& vbCrLf _
        '    '& "12/Aug/2016 10:47:21 AM(10:47:21) : Error: -2147467259 - Data provider or other service returned an E_FAIL status." & vbCrLf _
        '    '& "MSGTAG=KFT, Table=AmdLineKFT, Field=PSGRID, Field does not exist in Collection!, " & vbCrLf _
        '    '& "(Amadeus - AIR00002_20160811125358.Txt) Moved to folder 'Error Files'. PenGDS is automatically Resumed." _
        '    '& vbCrLf _
        '    '& vbCrLf _
        '    '& "12/Aug/2016 10:47:21 AM(10:47:21) : Error: -2147467259 - Data provider or other service returned an E_FAIL status." & vbCrLf _
        '    '& "MSGTAG=KFT, Table=AmdLineKFT, Field=PSGRID, Field does not exist in Collection!, " & vbCrLf _
        '    '& "(Amadeus - AIR00002_20160811125358.Txt) Moved to folder 'Error Files'. PenGDS is automatically Resumed."
        'DetailsForm.ShowME Me.Caption & " [Errors]", SendERROREmailPreviousMess_String
        'SendERROREmailPreviousMess_String = ""
        '########################################################################################################################
        
        If SSTab1.TabPicture(0) = ImageError.Picture Then
            If Trim(SendERROREmailPreviousMess_String) <> "" Then
                DetailsForm.ShowME Me.Caption & " [Errors]", SendERROREmailPreviousMess_String
            End If
        End If
    End If
    'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
End Sub

Private Sub SysInfo_PowerSuspend()
    INIWrite INIPenAIR_String, "PenGDS", "GDS", "OFF"
    End
End Sub

Private Sub tmrEnd_Timer()
On Error Resume Next
    If (UCase(Dir(App.Path & "\EndSQL.txt")) = UCase("EndSQL.txt") Or UCase(Dir(App.Path & "\EndGDS.txt")) = UCase("EndGDS.txt")) And fsObj.FileExists(App.Path & "\_UploadingSQL_") = False Then
        'by Abhi on 22-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module
        'Open (App.Path & "\ENDGDS.TXT") For Random As #1
        'by Abhi on 25-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module is Append
        'Open (App.Path & "\ENDGDS.TXT") For Output Shared As #1
        Open (App.Path & "\ENDGDS.TXT") For Append Shared As #1
            'by Abhi on 22-Sep-2010 for caseid 1505 Saving text content in ENDGDS.TXT to check from which module
            Print #1, "Updating... PenGDS-FMain-tmrEnd_Timer()"
        Close #1
        INIWrite INIPenAIR_String, "PenGDS", "GDS", "OFF"
        SendStatus "Updating"
        'by Abhi on 25-Sep-2010 for caseid 1505 will send a mail tp supoort
        'by Abhi on 04-Feb-2011 for caseid 1505 will send a mail to support for travelnet
        'by Abhi on 24-Nov-2011 for caseid 1505 ENDGDS.TXT to check from which module disabled
        'If UCase(Trim(vgsDatabase)) = "PENTRAVELUP" Or UCase(Trim(vgsDatabase)) = "PENTRAVELNET" Then
        '    Call SendENDEmail
        'End If
        End
    End If
    If Val(Winsock.Tag) <> Winsock.State And Winsock.State <> 0 Then
        Debug.Print Winsock.State & "=" & Winsock.Tag
        Winsock.Tag = Winsock.State
        DoEvents
        Debug.Print Winsock.State
    End If
    If (Winsock.State <> sckConnected And Winsock.State <> sckConnecting And Winsock.State <> sckInProgress) And Monitor = True Then
        Winsock.Close
        DoEvents
        JustConnected = True
        Winsock.Connect Host, Port
    End If
End Sub
Private Sub tmrUpload_Timer()
On Error GoTo PENErr
Dim ErrNumber As String
Dim ErrDescription As String
Dim UploadingGDS As String
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
Dim DeadlockRETRY_Integer As Integer
'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
'by Abhi on 17-Nov-2011 for caseid 1551 PenGDS and GDSAuto Monitoring in PenAIR
INIWrite INIPenAIR_String, "PenGDS", "PenGDSACTIVE", Now
'by Abhi on 02-Nov-2010 for caseid 1533 PenGDS Cannot start more transactions on this session
tmrUpload.Enabled = False
'by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
DeadlockRETRY:
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'commented by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
'dbCompany.BeginTrans
'commented by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
'Call getPENLINEID
'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
'If rsSelect.State = 1 Then rsSelect.Close
'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
'rsSelect.Open "Select WSPUPLOADDIRNAME,WSPDESTDIRNAME,WSPSTATUS,AMDUPLOADDIRNAME,AMDDESTDIRNAME,AMDSTATUS,SABREUPLOADDIRNAME,SABREDESTDIRNAME,SABRESTATUS,UPLOADDIRNAME,DESTDIRNAME,STATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
'rsSelect.Open "Select WSPUPLOADDIRNAME,WSPDESTDIRNAME,WSPSTATUS,AMDUPLOADDIRNAME,AMDDESTDIRNAME,AMDSTATUS,SABREUPLOADDIRNAME,SABREDESTDIRNAME,SABRESTATUS,UPLOADDIRNAME,DESTDIRNAME,STATUS From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table

'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
stbUpload.Panels(2).Text = ""
'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
stbUpload.Panels(3).Text = ""
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
'Debug.Print 0 / 0
'Debug.Print "a" + 1
'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email

If chkGalilieo.value = vbChecked Then
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'Upload_StatusMIR = IIf(IsNull(rsSelect!Status), 0, rsSelect!Status)
    Upload_StatusMIR = chkGalilieo.value
    'by Abhi on 16-Mar-2012 for caseid 2151 change waiting to Active - Waiting for GDS files... for pengds
    'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
    'stbUpload.Panels(1).Text = "Active - Waiting for GDS files..."
    stbUpload.Panels(1).Text = "Active - Waiting for GDS files... " & TimeFormat12HRS(Now)
    stbUpload.Panels(2).Text = "Galileo"
    'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
    With FUpload
        'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'If rsSelect.EOF = False Then
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
            'txtGalilieoSource = IIf(IsNull(rsSelect!UPLOADDIRNAME), "", rsSelect!UPLOADDIRNAME)
            'txtGalilieoDest = IIf(IsNull(rsSelect!DESTDIRNAME), "", rsSelect!DESTDIRNAME)
            'by Abhi on 15-May-2013 for caseid 3121 Pengds should not stuck on error it should move files to error folder for later checking
            stbUpload.Panels("SOURCEPATH").Text = txtGalilieoSource
            If txtGalilieoSource = "" Or txtGalilieoDest = "" Then
                GoTo ParaSabre
            End If
            UploadingGDS = "Galilieo"
            If PathExists(txtGalilieoSource) = False Or PathExists(txtGalilieoDest) = False Then
                SSTab1.TabPicture(1) = ImageError.Picture
                SendStatus "Error(Local Path)"
                GoTo ParaSabre
            End If
            SSTab1.TabPicture(1) = Image3.Picture
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
            '.txtDirName = IIf(IsNull(rsSelect!UPLOADDIRNAME), "", rsSelect!UPLOADDIRNAME)
            '.txtRelLocation = IIf(IsNull(rsSelect!DESTDIRNAME), "", rsSelect!DESTDIRNAME)
            .txtDirName = txtGalilieoSource
            .txtRelLocation = txtGalilieoDest
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'End If
        .File1.Path = .txtDirName.Text
        .File1.Refresh
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

ParaSabre:
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
If fForcedStop_Boolean = True Then
    Exit Sub
End If
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
stbUpload.Panels(2).Text = ""
'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
stbUpload.Panels(3).Text = ""
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
If chkSabre.value = vbChecked Then
    'Uploading Sabre Files
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'Upload_Status = IIf(IsNull(rsSelect!SABRESTATUS), 0, rsSelect!SABRESTATUS)
    Upload_Status = chkSabre.value
    'by Abhi on 16-Mar-2012 for caseid 2151 change waiting to Active - Waiting for GDS files... for pengds
    'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
    'stbUpload.Panels(1).Text = "Active - Waiting for GDS files..."
    stbUpload.Panels(1).Text = "Active - Waiting for GDS files... " & TimeFormat12HRS(Now)
    stbUpload.Panels(2).Text = "Sabre"
    'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
    With FSabreUpload
        'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'If rsSelect.EOF = False Then
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
            'txtSource = IIf(IsNull(rsSelect!SABREUPLOADDIRNAME), "", rsSelect!SABREUPLOADDIRNAME)
            'txtDest = IIf(IsNull(rsSelect!SABREDESTDIRNAME), "", rsSelect!SABREDESTDIRNAME)
            txtSource = txtSource
            txtDest = txtDest
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
            'by Abhi on 15-May-2013 for caseid 3121 Pengds should not stuck on error it should move files to error folder for later checking
            stbUpload.Panels("SOURCEPATH").Text = txtSource
            If txtSource = "" Or txtDest = "" Then
                GoTo ParaWorldspan
            End If
            UploadingGDS = "Sabre"
            If PathExists(txtSource) = False Or PathExists(txtDest) = False Then
                SSTab1.TabPicture(2) = ImageError.Picture
                SendStatus "Error(Local Path)"
                GoTo ParaWorldspan
            End If
            SSTab1.TabPicture(2) = Image3.Picture
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
            '.txtDirName = IIf(IsNull(rsSelect!SABREUPLOADDIRNAME), "", rsSelect!SABREUPLOADDIRNAME)
            '.txtRelLocation = IIf(IsNull(rsSelect!SABREDESTDIRNAME), "", rsSelect!SABREDESTDIRNAME)
            .txtDirName = txtSource
            .txtRelLocation = txtDest
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'End If
        .File1.Path = .txtDirName.Text
        .File1.Refresh
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

ParaWorldspan:
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
If fForcedStop_Boolean = True Then
    Exit Sub
End If
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
stbUpload.Panels(2).Text = ""
'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
stbUpload.Panels(3).Text = ""
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
If chkWorldspan.value = vbChecked Then
    'Uploading Worldspan
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'Upload_StatusWORLD = IIf(IsNull(rsSelect!WSPSTATUS), 0, rsSelect!WSPSTATUS)
    Upload_StatusWORLD = chkWorldspan.value
    'by Abhi on 16-Mar-2012 for caseid 2151 change waiting to Active - Waiting for GDS files... for pengds
    'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
    'stbUpload.Panels(1).Text = "Active - Waiting for GDS files..."
    stbUpload.Panels(1).Text = "Active - Waiting for GDS files... " & TimeFormat12HRS(Now)
    stbUpload.Panels(2).Text = "Worldspan"
    'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
    With FWorldSpan
        'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'If rsSelect.EOF = False Then
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
            'txtWorldspanSource = IIf(IsNull(rsSelect!WSPUPLOADDIRNAME), "", rsSelect!WSPUPLOADDIRNAME)
            'txtWorldspanDest = IIf(IsNull(rsSelect!WSPDESTDIRNAME), "", rsSelect!WSPDESTDIRNAME)
            txtWorldspanSource = txtWorldspanSource
            txtWorldspanDest = txtWorldspanDest
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
            'by Abhi on 15-May-2013 for caseid 3121 Pengds should not stuck on error it should move files to error folder for later checking
            stbUpload.Panels("SOURCEPATH").Text = txtWorldspanSource
            If txtWorldspanSource = "" Or txtWorldspanDest = "" Then
                GoTo ParaAmadeus
            End If
            UploadingGDS = "Worldspan"
            If PathExists(txtWorldspanSource) = False Or PathExists(txtWorldspanDest) = False Then
                SSTab1.TabPicture(3) = ImageError.Picture
                SendStatus "Error(Local Path)"
                GoTo ParaAmadeus
            End If
            SSTab1.TabPicture(3) = Image3.Picture
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
            '.txtDirName = IIf(IsNull(rsSelect!WSPUPLOADDIRNAME), "", rsSelect!WSPUPLOADDIRNAME)
            '.txtRelLocation = IIf(IsNull(rsSelect!WSPDESTDIRNAME), "", rsSelect!WSPDESTDIRNAME)
            .txtDirName = txtWorldspanSource
            .txtRelLocation = txtWorldspanDest
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'End If
        .File1.Path = .txtDirName.Text
        .File1.Refresh
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

ParaAmadeus:
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
If fForcedStop_Boolean = True Then
    Exit Sub
End If
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
stbUpload.Panels(2).Text = ""
'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
stbUpload.Panels(3).Text = ""
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
If chkAmadeus.value = vbChecked Then
    'Uploading Amadeus Files
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'Upload_StatusAIR = IIf(IsNull(rsSelect!AMDSTATUS), 0, rsSelect!AMDSTATUS)
    Upload_StatusAIR = chkAmadeus
    'by Abhi on 16-Mar-2012 for caseid 2151 change waiting to Active - Waiting for GDS files... for pengds
    'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
    'stbUpload.Panels(1).Text = "Active - Waiting for GDS files..."
    stbUpload.Panels(1).Text = "Active - Waiting for GDS files... " & TimeFormat12HRS(Now)
    stbUpload.Panels(2).Text = "Amadeus"
    'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
    With FAmadeusUpload
        'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'If rsSelect.EOF = False Then
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
            'txtAmadeusSource = IIf(IsNull(rsSelect!AMDUPLOADDIRNAME), "", rsSelect!AMDUPLOADDIRNAME)
            'txtAmadeusDest = IIf(IsNull(rsSelect!AMDDESTDIRNAME), "", rsSelect!AMDDESTDIRNAME)
            txtAmadeusSource = txtAmadeusSource
            txtAmadeusDest = txtAmadeusDest
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
            'by Abhi on 15-May-2013 for caseid 3121 Pengds should not stuck on error it should move files to error folder for later checking
            stbUpload.Panels("SOURCEPATH").Text = txtAmadeusSource
            If txtAmadeusSource = "" Or txtAmadeusDest = "" Then
                'by Abhi on 02-Nov-2010 for caseid 1533 PenGDS Cannot start more transactions on this session
                'Exit Sub
                'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
                'GoTo ParaFinal
                GoTo ParaLegacy
            End If
            UploadingGDS = "Amadeus"
            If PathExists(txtAmadeusSource) = False Or PathExists(txtAmadeusDest) = False Then
                SSTab1.TabPicture(4) = ImageError.Picture
                SendStatus "Error(Local Path)"
                'by Abhi on 02-Nov-2010 for caseid 1533 PenGDS Cannot start more transactions on this session
                'Exit Sub
                'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
                'GoTo ParaFinal
                GoTo ParaLegacy
            End If
            SSTab1.TabPicture(4) = Image3.Picture
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
            '.txtDirName = IIf(IsNull(rsSelect!AMDUPLOADDIRNAME), "", rsSelect!AMDUPLOADDIRNAME)
            '.txtRelLocation = IIf(IsNull(rsSelect!AMDDESTDIRNAME), "", rsSelect!AMDDESTDIRNAME)
            .txtDirName = txtAmadeusSource
            .txtRelLocation = txtAmadeusDest
            'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
        'End If
        .File1.Path = .txtDirName.Text
        .File1.Refresh
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

'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
ParaLegacy:
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
If fForcedStop_Boolean = True Then
    Exit Sub
End If
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
stbUpload.Panels(2).Text = ""
'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
stbUpload.Panels(3).Text = ""
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
If LegacyEnableCheck.value = vbChecked Then
    'Uploading Legacy Files
''    Upload_StatusLegacy = LegacyEnableCheck.value
    'by Abhi on 16-Mar-2012 for caseid 2151 change waiting to Active - Waiting for GDS files... for pengds
    'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
    'stbUpload.Panels(1).Text = "Active - Waiting for GDS files..."
    stbUpload.Panels(1).Text = "Active - Waiting for GDS files... " & TimeFormat12HRS(Now)
    stbUpload.Panels(2).Text = "Legacy"
    'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
    With FLegacy
        'by Abhi on 15-May-2013 for caseid 3121 Pengds should not stuck on error it should move files to error folder for later checking
        stbUpload.Panels("SOURCEPATH").Text = LegacySourcePathText
        If Trim(LegacySourcePathText) = "" Or Trim(LegacyDestinationPathText) = "" Then
            'by Abhi on 02-Nov-2010 for caseid 1533 PenGDS Cannot start more transactions on this session
            'Exit Sub
            GoTo ParaFinal
        End If
        UploadingGDS = "Legacy"
        If PathExists(LegacySourcePathText) = False Or PathExists(LegacyDestinationPathText) = False Then
            SSTab1.TabPicture(5) = ImageError.Picture
            SendStatus "Error(Local Path)"
            'by Abhi on 02-Nov-2010 for caseid 1533 PenGDS Cannot start more transactions on this session
            'Exit Sub
            GoTo ParaFinal
        End If
        SSTab1.TabPicture(5) = Image3.Picture
        
        .txtDirName = LegacySourcePathText
        .txtRelLocation = LegacyDestinationPathText
        .File1.Path = LegacySourcePathText
        .File1.Refresh
        cnt = .File1.ListCount
        .File1.Refresh
        DoEvents
        If .File1.ListCount > 0 Then
''           If FLegacy.Visible = False Then
''             If FUpload.Visible = True Then Unload FUpload
''             If FSabreUpload.Visible = True Then Unload FSabreUpload
             Load FLegacy
             .Refresh
''           End If
           .cmdDirUpload_Click
''        ElseIf .File1.ListCount = cnt And .File1.ListCount <> 0 And Air_flag = False Then
''           Air_flag = True 'if files are same allow upload once
''           If FAmadeusUpload.Visible = False Then
''             Unload FUpload
''             Unload FSabreUpload
''             Load FAmadeusUpload
''             .Refresh
''           End If
''           .cmdDirUpload_Click
        End If

    End With
End If

'by Abhi on 02-Nov-2010 for caseid 1533 PenGDS Cannot start more transactions on this session
ParaFinal:
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
If fForcedStop_Boolean = True Then
    Exit Sub
End If
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
stbUpload.Panels(2).Text = ""
'by Abhi on 29-Jul-2014 for caseid 4365 Enabled time in PenGDS status bar to know activity
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
stbUpload.Panels(3).Text = ""
'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'commented by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
'dbCompany.CommitTrans
'commented by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
'by Abhi on 02-Nov-2010 for caseid 1533 PenGDS Cannot start more transactions on this session
'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
'tmrUpload.Enabled = True
If fForcedStop_Boolean = False Then
    tmrUpload.Enabled = True
End If
'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
DoEvents
Exit Sub
PENErr:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    PermissionDenied = INIRead(INIPenAIR_String, "PenGDS", "PermissionDenied", "5") '5
    If Val(ErrNumber) = 70 And NoofPermissionDenied < PermissionDenied Then
        NoofPermissionDenied = NoofPermissionDenied + 1
        Sleep 500
        Resume
    'by Abhi on 24-Jul-2010 for caseid 1436 Deadlock in PenGDS
    'commented by Abhi on 27-Oct-2010 for caseid 1527 DeadlockRETRY
    'ElseIf ErrNumber = -2147467259 Then 'Deadlock
    '    Debug.Print "Deadlock"
    '    Resume
    Else
        If Val(ErrNumber) = 0 Then
            'SendERROR "Warning in PenGDS[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ", PenGDS is automatically Resumed."
            'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
            'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ", " & App.ProductName & " is automatically Resumed."
            'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
            'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ", " & App.ProductName & " is automatically Resumed.", ErrNumber
            SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ", " & ErrDetails_String & App.ProductName & " is automatically Resumed.", ErrNumber
            ErrDetails_String = ""
            'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
            'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
            Resume
        Else
            'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
            'by Abhi on 26-Oct-2010 for caseid 1527 DeadlockRETRY
            'dbCompany.RollbackTrans
            'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
            If PENErr_BeginTrans = True Then
            'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
                If ErrNumber <> -2147168242 Then dbCompany.RollbackTrans
            'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
                PENErr_BeginTrans = False
            End If
            'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
            'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
            'If ErrNumber = -2147467259 Then 'Deadlock
            If (ErrNumber = -2147467259 Or ErrNumber = -2147217871) And DeadlockRETRY_Integer < 3 Then '-2147467259 Deadlock, -2147217871 Query timeout expired
                DeadlockRETRY_Integer = DeadlockRETRY_Integer + 1
            'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
                Debug.Print "Deadlock"
                Sleep 5
                'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                'GoTo DeadlockRETRY
                Resume DeadlockRETRY
                'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
            End If
            cmdStop_Click
            NoofPermissionDenied = 0
            'by Abhi on 25-Mar-2013 for caseid 2786 PenGDS Sabre PNR missing - SabreM1 row not found. we need to skip this file
            If Val(ErrNumber) = -2147220991 Then
                'SendERROR "Warning in PenGDS[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & ") Moved to folder 'Error M1 Missing'. PenGDS is automatically Resumed."
                'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & ") Moved to folder 'Error M1 Missing'. " & App.ProductName & " is automatically Resumed."
                'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & ") Moved to folder 'Error M1 Missing'. " & App.ProductName & " is automatically Resumed.", ErrNumber
                SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ErrDetails_String & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & ") Moved to folder 'Error M1 Missing'. " & App.ProductName & " is automatically Resumed.", ErrNumber
                ErrDetails_String = ""
                'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                FMain.stbUpload.Panels(2).Text = ""
                FMain.stbUpload.Panels(3).Text = ""
            Else
                'by Abhi on 15-May-2013 for caseid 3121 Pengds should not stuck on error it should move files to error folder for later checking
                If Trim(stbUpload.Panels(2).Text) = "" And Trim(stbUpload.Panels(3).Text) = "" Or Val(ErrNumber) = -2147217871 Then
                    'SendERROR "Warning in PenGDS[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & "). PenGDS is automatically Resumed."
                    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                    'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed."
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                    'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed.", ErrNumber
                    SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ErrDetails_String & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed.", ErrNumber
                    ErrDetails_String = ""
                    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                    FMain.stbUpload.Panels(2).Text = ""
                    FMain.stbUpload.Panels(3).Text = ""
                Else
                    'by Abhi on 15-May-2013 for caseid 3121 Pengds should not stuck on error it should move files to error folder for later checking
                    'Me.WindowState = vbNormal
                    'Me.Show
                    'SendStatus "Error"
                    'SendERROR "ERROR in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & ")"
                    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
                    If Trim(FMain.stbUpload.Panels("FILENAME").Text) <> "" Then
                    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
                        If Dir(stbUpload.Panels("SOURCEPATH").Text & "\Error Files\", vbDirectory) = "" Then
                            MkDir stbUpload.Panels("SOURCEPATH").Text & "\Error Files\"
                        End If
                        'by Abhi on 07-Jul-2014 for caseid 4258 Change in logic for file name checking when moving to target or error files
                        FMain.stbUpload.Panels("FILENAME").Text = IfExistsinTargetRename(FMain.stbUpload.Panels("SOURCEPATH").Text, FMain.stbUpload.Panels("FILENAME").Text, FMain.stbUpload.Panels("SOURCEPATH").Text & "\Error Files\")
                        'by Abhi on 07-Jul-2014 for caseid 4258 Change in logic for file name checking when moving to target or error files
                        fsObj.CopyFile stbUpload.Panels("SOURCEPATH").Text & "\" & stbUpload.Panels("FILENAME").Text, stbUpload.Panels("SOURCEPATH").Text & "\Error Files\", True
                        'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
                        'On Error Resume Next
                        'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
                        fsObj.DeleteFile stbUpload.Panels("SOURCEPATH").Text & "\" & stbUpload.Panels("FILENAME").Text, True
                        'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
                        'On Error GoTo 0
                        'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
                        'SendERROR "Warning in PenGDS[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & ") Moved to folder 'Error Files'. PenGDS is automatically Resumed."
                        'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                        'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & ") Moved to folder 'Error Files'. " & App.ProductName & " is automatically Resumed."
                        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                        'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & ") Moved to folder 'Error Files'. " & App.ProductName & " is automatically Resumed.", ErrNumber
                        SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ErrDetails_String & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & ") Moved to folder 'Error Files'. " & App.ProductName & " is automatically Resumed.", ErrNumber
                        ErrDetails_String = ""
                        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                        'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
                    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
                    Else
                        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                        'SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed.", ErrNumber
                        SendERROR "Warning in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & ErrDetails_String & " (" & stbUpload.Panels(2).Text & " - " & stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed.", ErrNumber
                        ErrDetails_String = ""
                        'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
                    End If
                    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
                    FMain.stbUpload.Panels(2).Text = ""
                    FMain.stbUpload.Panels(3).Text = ""
                End If
            End If
            'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
            'ret = MsgBox("Error: " & ErrNumber & vbCrLf & ErrDescription, vbAbortRetryIgnore, App.Title)
            'If ret = vbAbort Then
            '    ShowStopAccess = False
            '    cmdStop_Click
            'ElseIf ret = vbRetry Then
            '    Resume
            'ElseIf ret = vbIgnore Then
            '    Resume Next
            'End If
            'by Abhi on 25-Mar-2013 for caseid 2786 PenGDS Sabre PNR missing - SabreM1 row not found. we need to skip this file
            'by Abhi on 15-May-2013 for caseid 3121 Pengds should not stuck on error it should move files to error folder for later checking
            'If Val(ErrNumber) = -2147220991 Then
                ret = vbOK
            'Else
            '    ret = MsgBox("Error: " & ErrNumber & vbCrLf & ErrDescription, vbOKCancel + vbCritical, App.Title)
            'End If
            If ret = vbOK Then
                cmdStart_Click
            End If
        End If
    End If
End Sub

Function SetEnableAll()
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    'If chkGalilieo.value = vbChecked And chkSabre.value = vbChecked And chkWorldspan.value = vbChecked And chkAmadeus.value = vbChecked Then
    If chkGalilieo.value = vbChecked And chkSabre.value = vbChecked And chkWorldspan.value = vbChecked And chkAmadeus.value = vbChecked And LegacyEnableCheck.value = vbChecked Then
        chkEnableAll.value = vbChecked
    'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
    'ElseIf chkGalilieo.value = vbUnchecked And chkSabre.value = vbUnchecked And chkWorldspan.value = vbUnchecked And chkAmadeus.value = vbUnchecked Then
    ElseIf chkGalilieo.value = vbUnchecked And chkSabre.value = vbUnchecked And chkWorldspan.value = vbUnchecked And chkAmadeus.value = vbUnchecked And LegacyEnableCheck.value = vbUnchecked Then
        chkEnableAll.value = vbUnchecked
    Else
        chkEnableAll.value = vbGrayed
    End If
End Function

Private Sub TrayArea1_DblClick()
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub txtAmadeusDest_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtAmadeusDest_GotFocus()
    SelectText txtAmadeusDest
End Sub

Private Sub txtAmadeusExt_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtAmadeusExt_GotFocus()
    SelectText txtAmadeusExt
End Sub

Private Sub txtAmadeusSource_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtAmadeusSource_GotFocus()
    SelectText txtAmadeusSource
End Sub

Private Sub txtDest_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtDest_GotFocus()
    SelectText txtDest
End Sub

Private Sub txtGalileoExt_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtGalileoExt_GotFocus()
    SelectText txtGalileoExt
End Sub

Private Sub txtGalilieoDest_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtGalilieoDest_GotFocus()
    SelectText txtGalilieoDest
End Sub

Private Sub txtGalilieoSource_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtGalilieoSource_GotFocus()
    SelectText txtGalilieoSource
End Sub

Private Sub txtSabreExt_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtSabreExt_GotFocus()
    SelectText txtSabreExt
End Sub

Private Sub txtSource_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtSource_GotFocus()
    SelectText txtSource
End Sub

Private Sub txtWorldspanDest_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtWorldspanDest_GotFocus()
    SelectText txtWorldspanDest
End Sub

Private Sub txtWorldspanExt_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtWorldspanExt_GotFocus()
    SelectText txtWorldspanExt
End Sub

Private Sub txtWorldspanSource_Change()
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    AnyDataChanged_Boolean = True
End Sub

Private Sub txtWorldspanSource_GotFocus()
    SelectText txtWorldspanSource
End Sub

Private Sub Winsock_Close()
    JustConnected = False
    Winsock.Close
End Sub

Public Function SendStatus(ByVal STATUS_StringVar As String)
    If Winsock.State = sckConnected And Monitor = True Then
        SendComplete = False
        Winsock.SendData "Status;" & ClientName & ";" & "PenGDS Interface" & ";" & STATUS_StringVar
    End If
    Waiting4SendComplete
End Function

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim vlsData As String
    Winsock.GetData vlsData
    If vlsData = "CheckStatus" Then
        If cmdStart.Enabled = True Then
            SendStatus "Stopped"
        ElseIf cmdStart.Enabled = False Then
            SendStatus "Started"
        End If
    ElseIf vlsData = "Restart" Then
        If cmdStop.Enabled = True Then
            ShowStopAccess = False
            cmdStop_Click
            Sleep 3000
            cmdStart_Click
        End If
    ElseIf vlsData = "Stop" Then
        cmdStart.Tag = "Stop"
        If cmdStop.Enabled = True Then
            ShowStopAccess = False
            cmdStop_Click
        End If
    ElseIf vlsData = "Start" Then
        cmdStart.Tag = ""
        If cmdStart.Enabled = True Then
            cmdStart_Click
        End If
    End If
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Winsock_Close
    Winsock.Close
End Sub

Private Sub Winsock_SendComplete()
    If JustConnected = True Then
        JustConnected = False
        If cmdStart.Enabled = True Then
            SendStatus "Stopped"
        ElseIf cmdStart.Enabled = False Then
            SendStatus "Started"
        End If
    End If
    SendComplete = True
End Sub

Private Sub Winsock_Connect()
    'Caption = "Connected"
    SendStatus "Online"
End Sub

Public Function Waiting4SendComplete()
    If Monitor = True And (Winsock.State = sckConnected Or Winsock.State = sckInProgress) Then
        While Not SendComplete
            DoEvents
        Wend
    End If
End Function

Private Sub Winsock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    SendComplete = False
End Sub
