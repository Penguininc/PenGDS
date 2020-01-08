VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FSabreUpload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sabre Uploading..."
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSabre 
      Caption         =   "Include"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Left            =   390
      TabIndex        =   10
      Top             =   1200
      Width           =   1365
   End
   Begin VB.TextBox txtRelLocation 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1680
      TabIndex        =   8
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Files Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtDirName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1680
      TabIndex        =   5
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdDirUpload 
      Caption         =   "DirUpload"
      Height          =   375
      Left            =   3780
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   5070
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar Pgr 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Upload"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2070
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   2070
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   " Rel. Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   735
      Width           =   1440
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      Caption         =   " Dir. Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   375
      Width           =   1440
   End
End
Attribute VB_Name = "FSabreUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UploadNo As Long
Dim FNAME As String
Dim avpItems
Dim vlpLen As Long
Dim ai As Long
Dim avpItemsTemp As String
Dim vlTI As Long, vlTCount As Long
Dim Preid_String As String
Dim PreidNO_Long As Long
Dim aPEN
Dim aPENFARE
'by Abhi on 27-Aug-2010 for caseid 1473 Penline PENO
Dim aPENO
'by Abhi on 27-Aug-2010 for caseid 708 BLK should affect on above tickets for PER PP
Dim fNOBLKPP_String As String
'by Abhi on 31-Aug-2010 for caseid 1433 Penline PENAUTOOFF
Dim aPENAUTOOFF
'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
Dim aiTo_Long As Long
'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking
Dim aPENAGROSS
'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking
'by Abhi on 15-Oct-2018 for caseid 9304 Passenger Type picking issue -Sabre file
Dim faPENFARE4_String As String
'by Abhi on 15-Oct-2018 for caseid 9304 Passenger Type picking issue -Sabre file
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
Dim aPENCS
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

Private Sub chkSabre_Click()
'If chkSabre.Value = 0 Then
'    dbCompany.Execute "Update File set SABRESTATUS=0"
'Else
'    dbCompany.Execute "Update File set SABRESTATUS=1"
'End If
End Sub

Public Sub cmdDirUpload_Click()
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'On Error GoTo Note
'by Abhi on 07-Jul-2014 for caseid 4262 Sabre File not found unknown error
'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
'On Error GoTo PENErr
'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
Dim ErrNumber As String
Dim ErrDescription As String
'by Abhi on 07-Jul-2014 for caseid 4262 Sabre File not found unknown error
    Dim Count As Integer
    
    FMain.SendStatus FMain.SSTab1.TabCaption(2)
    File1.Path = txtDirName.Text
    File1.Pattern = FMain.txtSabreExt
    If txtDirName.Text = "" Or File1.ListCount = 0 Then stbUpload.Panels(1).Text = "FIL Files Not Found....": Exit Sub
    'Uploading Each File
    Open (App.Path & "\_UploadingSQL_") For Random As #1
    Close #1
    FMain.cmdStop.Enabled = False
    DoEvents
    FMain.lblFName.Caption = "Uploading Sabre..."
    FMain.stbUpload.Panels(1).Text = "Reading..."
    FMain.stbUpload.Panels(2).Text = "Sabre"
    FMain.prgUpload.Max = IIf((File1.ListCount = 1), File1.ListCount, File1.ListCount - 1)
    Pgr.value = 5
    'by Abhi on 18-May-2011 for caseid 1757 Added Events for GDSAuto
    Call EventLogDelete
    'by Abhi on 06-Sep-2017 for caseid 7766 Penline and Company currency change in penair is not reflecting in pengds
    Call getPENLINEID
    COMCID_String = getFromFileTable("COMCID")
    'by Abhi on 06-Sep-2017 for caseid 7766 Penline and Company currency change in penair is not reflecting in pengds
    For Count = 0 To File1.ListCount - 1
        'by Abhi on 17-Nov-2011 for caseid 1551 PenGDS and GDSAuto Monitoring in PenAIR
        INIWrite INIPenAIR_String, "PenGDS", "PenGDSACTIVE", Now
        DoEvents
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        ErrDetails_String = ""
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
        LUFPNR_String = ""
        'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
        FMain.stbUpload.Panels(3).Text = File1.List(Count)
        txtFileName.Text = File1.Path & "\" & File1.List(Count)
        'by Abhi on 09-Nov-2012 for caseid 2368 Sabre pnr missing checking sequence changed
        'by Abhi on 01-Nov-2012 for caseid 2368 Sabre pnr missing
        'commented by Abhi on 09-Nov-2012 for caseid 2368 Sabre pnr missing checking sequence changed
        'Call EventLog("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & "      " & " - " & FMain.stbUpload.Panels(3).Text & ", Wait4EOFSabre...")
        'by Abhi on 03-Nov-2012 for caseid 2368 Sabre pnr missing File not found
        'Call Wait4EOFSabre(File1.Path & "\" & File1.List(Count))
        'by Abhi on 09-Nov-2012 for caseid 2368 Sabre pnr missing checking sequence changed
        'Call Wait4EOFSabre(File1.Path & "\" & txtFileName.Text)
        'commented by Abhi on 09-Nov-2012 for caseid 2368 Sabre pnr missing checking sequence changed
        'Call Wait4EOFSabre(txtFileName.Text)
        'by Abhi on 09-Nov-2012 for caseid 2368 Sabre pnr missing checking sequence changed
        'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
        'Sleep 3000
        'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
        'commented by Abhi on 09-Nov-2012 for caseid 2368 Sabre pnr missing checking sequence changed
        'Call EventLog("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & "      " & " - " & FMain.stbUpload.Panels(3).Text & ", Wait4EOFSabre... Found")
        txtFileName.Text = IfExistsinTargetRename(File1.Path, File1.List(Count), txtRelLocation.Text)
        'by Abhi on 07-Oct-2010 for caseid 1516 PenGDS Amadeus slow reading
        'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
        'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
        'If Trim(txtFileName.Text) = "" Then Exit Sub
        If Trim(txtFileName.Text) <> "" Then
        'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
            'by Abhi on 26-Jul-2014 for caseid 4347 PenGDS stuck on process Waiting for file available
            FMain.stbUpload.Panels(1).Text = "Reading..."
            FNAME = txtFileName.Text
            'FMain.SendStatus FNAME
            txtFileName.Text = File1.Path & "\" & txtFileName.Text
            cmdLoad_Click
        'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
        End If
        'by Abhi on 16-Aug-2014 for caseid 4440 PenGDS Warning(IfExistsinTargetRename) Error: 53 - File not found
        If Pgr.value < 90 Then
            Pgr.value = Pgr.value + 10
        End If
        FMain.prgUpload.value = Count
        DoEvents
    Next
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    If FMain.fForcedStop_Boolean = True Then
        Exit Sub
    End If
    'by Abhi on 22-Jul-2016 for caseid 6609 PenGDS "Waiting for file available" message can be replaced with "File is in use, waiting..."
    Pgr.value = 100
    File1.Refresh
    FMain.prgUpload.value = 0
    FMain.stbUpload.Panels(1).Text = "Process Completed."
    FMain.cmdStop.Enabled = True
    FMain.stbUpload.Panels(2).Text = ""
    FMain.stbUpload.Panels(3).Text = ""
    'by Abhi on 16-Mar-2012 for caseid 2151 change waiting to Active - Waiting for GDS files... for pengds
    FMain.stbUpload.Panels(1).Text = "Active - Waiting for GDS files..."
    'FMain.txtSource = ""
    'FMain.txtDest = ""
    FMain.lblFName.Caption = ""
    'by Abhi on 09-Nov-2012 for caseid 2368 Sabre pnr missing checking sequence changed
    If UCase(Dir(App.Path & "\_UploadingSQL_")) = UCase("_UploadingSQL_") Then
        fsObj.DeleteFile (App.Path & "\_UploadingSQL_"), True
    End If
    FMain.SendStatus "Started"
    Me.Hide
    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
    'Exit Sub
    'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'Note:
'    Me.Hide
'by Abhi on 07-Jul-2014 for caseid 4262 Sabre File not found unknown error
'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
'Exit Sub
'PENErr:
'    ErrNumber = Err.Number
'    ErrDescription = Err.Description
'
'    FMain.cmdStop_Click
'    NoofPermissionDenied = 0
'    'SendERROR "Warning(Sabre) in PenGDS[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & "). PenGDS is automatically Resumed."
'    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
'    'SendERROR "Warning(Sabre) in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed."
'    SendERROR "Warning(Sabre) in " & App.ProductName & "[" & CHEAD & "] (" & vgsDatabase & ")", "Error: " & ErrNumber & " - " & ErrDescription & " (" & FMain.stbUpload.Panels(2).Text & " - " & FMain.stbUpload.Panels(3).Text & "). " & App.ProductName & " is automatically Resumed.", ErrNumber
'    'by Abhi on 12-Dec-2014 for caseid 4820 Delay of 5 minutes for pengds error notification email
'    FMain.cmdStart_Click
''by Abhi on 07-Jul-2014 for caseid 4262 Sabre File not found unknown error
'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
End Sub
Private Sub cmdLoad_Click()
'On Error GoTo Note
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
On Error GoTo PENErr
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    Dim tsObj As TextStream
    Dim LineStr As String
    Dim lnID As String
    Dim CopyStatus As Boolean
    Dim Lines, crrItem As String
    'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    Dim vSQL_String As String
    'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    Dim vDateLastModified_String As String
    'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
    'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
    Dim vFileDateCreated_String As String
    'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
    
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    GDSDeadlockRETRY_Integer = 0
    'by Abhi on 22-Nov-2014 for caseid 4736 Query timeout expired in PenGDS
    
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
DeadlockRETRY:
'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    CopyStatus = False
    'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
    LUFPNR_String = ""
    'by Abhi on 27-Aug-2010 for caseid 708 BLK should affect on above tickets for PER PP
    fNOBLKPP_String = ""
    'by Abhi on 08-Mar-2013 for caseid 2958 PenGDS Sabre pnr file not refeshing
    Preid_String = ""
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    NoofPermissionDenied = 0
    'by Abhi on 15-Oct-2013 for caseid 3455 No transaction is active if the file is stuck in PenFTP side
    'by Abhi on 21-Apr-2010 for caseid 1320 NOFOLDER for in penline
    
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    GIT_CUSTOMERUSERCODE_String = ""
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
    GIT_PENWAIT_String = "N"
    'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
    
    If isNOFOLDERExists(txtFileName) = False Then
        Set tsObj = fsObj.OpenTextFile(txtFileName, ForReading)
        'by Abhi on 04-Jan-2010 for caseid 1582 first free numbers sequence number genaration
        'UploadNo = mSeqNumberGen("UPLOADNO")
        'by Abhi on 13-Feb-2012 for caseid 1582 first free numbers as stored procedure in pengds
        'UploadNo = PENFirstFreeNumber("UPLOADNO")
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        dbCompany.BeginTrans
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        PENErr_BeginTrans = True
        'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        UploadNo = FirstFreeNumber("UPLOADNO")
        PreidNO_Long = 0
        'by Abhi on 11-Nov-2012 for caseid 2368 Sabre pnr missing file content log
        'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
        'Call EventLogSabre("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & "      " & " - " & FMain.stbUpload.Panels(3).Text, True)
        Call EventLogSabre("#PenGDS# " & PadR(FMain.stbUpload.Panels(2).Text, 9) & " - " & PadR("", 6) & " - " & PadR(FMain.stbUpload.Panels(3).Text, 27), True)
        'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
        Do While tsObj.AtEndOfStream <> True
            DoEvents
            LineStr = tsObj.ReadLine
            'by Abhi on 04-Mar-2013 for caseid 2958 PenGDS Sabre pnr file not refeshing
            Call EventLogSabre(LineStr)
            'If fnChecktheDatabase(LineStr) = True Then CopyStatus = True: Exit Do
            Lines = Split(LineStr, Chr(13))
            'by Abhi on 08-Mar-2013 for caseid 2958 PenGDS Sabre pnr file not refeshing
            Call EventLogSabre("UBound(Lines)=" & UBound(Lines))
            If FMain.chkSabreIncludeItineraryOnly.value = vbUnchecked Then
                If UBound(Lines) <> -1 Then
                    If Left(Lines(0), 2) = "AA" And Mid(Lines(0), 3, 1) <> "/" Then
                        If Val(Mid(Lines(0), 14, 1)) = 3 Then 'Itinerary Only (DIT ENTRIES) 'IU0TYP
                            Exit Do
                        End If
                    End If
                End If
            End If
            If Preid_String = "AA" Then
                ReDim Preserve Lines(0)
            End If
            For j = 0 To UBound(Lines)
                crrItem = Lines(j)
                'commented by Abhi on 25-Feb-2013 for caseid 2958 PenGDS Sabre pnr file not refeshing
                'Call EventLogSabre(crrItem)
                lnID = Left(crrItem, 2)
                'by Abhi on 01-Nov-2012 for caseid 2368 Sabre pnr missing
                'Call EventLogSabre("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & LUFPNR_String & " - " & FMain.stbUpload.Panels(3).Text & ", LID: " & lnID)
                SelectDetails crrItem, lnID
                DoEvents
            Next
            
            'lnID = Left(LineStr, 2)
            'SelectDetails LineStr, lnID
            DoEvents
        Loop
        tsObj.Close
        'by Abhi on 04-Mar-2013 for caseid 2958 PenGDS Sabre pnr file not refeshing
        Set tsObj = Nothing
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
        FMain.stbUpload.Panels(1).Text = "Writing GDS In Tray..."
        DoEvents
        'by Abhi on 19-Oct-2013 for caseid 3465 Sabre file loading issue pick TRANSACTION TYPE=A,B,C
        'vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC]) " _
            & "SELECT     'FIL' AS GIT_ID, 'SABRE' AS GIT_GDS, 'SABRE' AS GIT_INTERFACE, dbo.SabreM0.UPLOADNO AS GIT_UPLOADNO, SUBSTRING(dbo.SabreM0.IU0MDT, 1, 2) " _
            & "                      + '/' + SUBSTRING(dbo.SabreM0.IU0MDT, 3, 3) + '/' + '" & Year(Date) & "' AS GIT_PNRDATE, dbo.SabreM0.IU0PNR AS GIT_PNR, dbo.SabreM1.IU1PNM AS GIT_LASTNAME, " _
            & "                      '' AS GIT_FIRSTNAME, dbo.SabreM0.IU0IAG AS GIT_BOOKINGAGENT, dbo.SabreM5.IU5VR4 AS TKTNBRGIT_TICKETNUMBER, dbo.SabreM0.FNAME AS GIT_FILENAME, " _
            & "                       dbo.SabreM0.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.SabreM0.IU0PCC AS GIT_PCC " _
            & "FROM         dbo.SabreM0 INNER JOIN " _
            & "                      dbo.SabreM1 ON dbo.SabreM0.UPLOADNO = dbo.SabreM1.UPLOADNO LEFT OUTER JOIN " _
            & "                      dbo.SabreM5 ON CAST(CASE WHEN isnumeric(LEFT(SabreM5.IU5MIN, 2)) = 1 THEN LEFT(SabreM5.IU5MIN, 2) ELSE SabreM5.IU5PTY END AS int) " _
            & "                      = dbo.SabreM1.IU1PNO AND dbo.SabreM5.UPLOADNO = dbo.SabreM1.UPLOADNO " _
            & "WHERE     ( (dbo.SabreM0.UPLOADNO = " & UploadNo & ") AND dbo.SabreM0.IU0TYP = '1' OR " _
            & "                      dbo.SabreM0.IU0TYP = '2' OR " _
            & "                      dbo.SabreM0.IU0TYP = '3' OR " _
            & "                      dbo.SabreM0.IU0TYP = '5' OR " _
            & "                      dbo.SabreM0.IU0TYP = 'D') AND (dbo.SabreM5.IU5VR4 IS NULL OR " _
            & "                      dbo.SabreM5.IU5VR4 <> 'BLK' AND dbo.SabreM5.IU5VR4 <> 'DIS' AND dbo.SabreM5.IU5VR4 <> 'TBR')"
        'by Abhi on 07-Nov-2013 for caseid 3512 PenGDS Sabre Duplication
        'vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC]) " _
            & "SELECT     'FIL' AS GIT_ID, 'SABRE' AS GIT_GDS, 'SABRE' AS GIT_INTERFACE, dbo.SabreM0.UPLOADNO AS GIT_UPLOADNO, SUBSTRING(dbo.SabreM0.IU0MDT, 1, 2) " _
            & "                      + '/' + SUBSTRING(dbo.SabreM0.IU0MDT, 3, 3) + '/' + '" & Year(Date) & "' AS GIT_PNRDATE, dbo.SabreM0.IU0PNR AS GIT_PNR, dbo.SabreM1.IU1PNM AS GIT_LASTNAME, " _
            & "                      '' AS GIT_FIRSTNAME, dbo.SabreM0.IU0IAG AS GIT_BOOKINGAGENT, dbo.SabreM5.IU5VR4 AS TKTNBRGIT_TICKETNUMBER, dbo.SabreM0.FNAME AS GIT_FILENAME, " _
            & "                       dbo.SabreM0.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.SabreM0.IU0PCC AS GIT_PCC " _
            & "FROM         dbo.SabreM0 INNER JOIN " _
            & "                      dbo.SabreM1 ON dbo.SabreM0.UPLOADNO = dbo.SabreM1.UPLOADNO LEFT OUTER JOIN " _
            & "                      dbo.SabreM5 ON CAST(CASE WHEN isnumeric(LEFT(SabreM5.IU5MIN, 2)) = 1 THEN LEFT(SabreM5.IU5MIN, 2) ELSE SabreM5.IU5PTY END AS int) " _
            & "                      = dbo.SabreM1.IU1PNO AND dbo.SabreM5.UPLOADNO = dbo.SabreM1.UPLOADNO " _
            & "WHERE     ( (dbo.SabreM0.UPLOADNO = " & UploadNo & ") AND dbo.SabreM0.IU0TYP = '1' OR " _
            & "                      dbo.SabreM0.IU0TYP = '2' OR " _
            & "                      dbo.SabreM0.IU0TYP = '3' OR " _
            & "                      dbo.SabreM0.IU0TYP = '5' OR " _
            & "                      dbo.SabreM0.IU0TYP = 'D' OR " _
            & "                      dbo.SabreM0.IU0TYP = 'A' OR " _
            & "                      dbo.SabreM0.IU0TYP = 'B' OR " _
            & "                      dbo.SabreM0.IU0TYP = 'C') AND (dbo.SabreM5.IU5VR4 IS NULL OR " _
            & "                      dbo.SabreM5.IU5VR4 <> 'BLK' AND dbo.SabreM5.IU5VR4 <> 'DIS' AND dbo.SabreM5.IU5VR4 <> 'TBR')"
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'vSQL_String = "" _
        '    & "INSERT INTO dbo.GDSINTRAYTABLE " _
        '    & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC]) " _
        '    & "SELECT     'FIL' AS GIT_ID, 'SABRE' AS GIT_GDS, 'SABRE' AS GIT_INTERFACE, dbo.SabreM0.UPLOADNO AS GIT_UPLOADNO, SUBSTRING(dbo.SabreM0.IU0MDT, 1, 2) " _
        '    & "                      + '/' + SUBSTRING(dbo.SabreM0.IU0MDT, 3, 3) + '/' + '" & Year(Date) & "' AS GIT_PNRDATE, dbo.SabreM0.IU0PNR AS GIT_PNR, dbo.SabreM1.IU1PNM AS GIT_LASTNAME, " _
        '    & "                      '' AS GIT_FIRSTNAME, dbo.SabreM0.IU0IAG AS GIT_BOOKINGAGENT, dbo.SabreM5.IU5VR4 AS TKTNBRGIT_TICKETNUMBER, dbo.SabreM0.FNAME AS GIT_FILENAME, " _
        '    & "                       dbo.SabreM0.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.SabreM0.IU0PCC AS GIT_PCC " _
        '    & "FROM         dbo.SabreM0 INNER JOIN " _
        '    & "                      dbo.SabreM1 ON dbo.SabreM0.UPLOADNO = dbo.SabreM1.UPLOADNO LEFT OUTER JOIN " _
        '    & "                      dbo.SabreM5 ON CAST(CASE WHEN isnumeric(LEFT(SabreM5.IU5MIN, 2)) = 1 THEN LEFT(SabreM5.IU5MIN, 2) ELSE SabreM5.IU5PTY END AS int) " _
        '    & "                      = dbo.SabreM1.IU1PNO AND dbo.SabreM5.UPLOADNO = dbo.SabreM1.UPLOADNO " _
        '    & "WHERE     (dbo.SabreM0.UPLOADNO = " & UploadNo & ") AND (dbo.SabreM0.IU0TYP = '1' OR " _
        '    & "                      dbo.SabreM0.IU0TYP = '2' OR " _
        '    & "                      dbo.SabreM0.IU0TYP = '3' OR " _
        '    & "                      dbo.SabreM0.IU0TYP = '5' OR " _
        '    & "                      dbo.SabreM0.IU0TYP = 'D' OR " _
        '    & "                      dbo.SabreM0.IU0TYP = 'A' OR " _
        '    & "                      dbo.SabreM0.IU0TYP = 'B' OR " _
        '    & "                      dbo.SabreM0.IU0TYP = 'C') AND (dbo.SabreM5.IU5VR4 IS NULL OR " _
        '    & "                      dbo.SabreM5.IU5VR4 <> 'BLK' AND dbo.SabreM5.IU5VR4 <> 'DIS' AND dbo.SabreM5.IU5VR4 <> 'TBR')"
        vDateLastModified_String = fsObj.GetFile(txtFileName).DateLastModified
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
        vDateLastModified_String = DateTime12hrsFormat(vDateLastModified_String)
        vFileDateCreated_String = fsObj.GetFile(txtFileName).DateCreated
        vFileDateCreated_String = DateTime12hrsFormat(vFileDateCreated_String)
        'by Abhi on 22-Jul-2017 for caseid 7651 GDS File delay checking separately for each GDS
        'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
        vSQL_String = "" _
            & "SELECT     TOP (1) BR " _
            & "FROM         dbo.SabreM9 " _
            & "WHERE     (UPLOADNO = " & UploadNo & ") AND (PEN = PEN) AND (BR <> '') " _
            & "ORDER BY ITNO"
        GIT_PENLINEBRID_String = getFromExecuted(vSQL_String, "BR")
        'by Abhi on 08-Aug-2019 for caseid 10556 GDS Tray filter by user or branch
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        ErrDetails_String = " GDSINTRAYTABLE."
        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        vSQL_String = "" _
            & "INSERT INTO dbo.GDSINTRAYTABLE " _
            & "([GIT_ID],[GIT_GDS],[GIT_INTERFACE],[GIT_UPLOADNO],[GIT_PNRDATE],[GIT_PNR],[GIT_LASTNAME],[GIT_FIRSTNAME],[GIT_BOOKINGAGENT],[GIT_TICKETNUMBER],[GIT_FILENAME],[GIT_GDSAUTOFAILED],[GIT_PCC],[GIT_FILEISSUEDATE],[GIT_CUSTOMERUSERCODE],[GIT_PENWAIT],[GIT_FILECREATEDDATE],[GIT_ISSUEDBY], [GIT_PENLINEBRID]) " _
            & "SELECT     'FIL' AS GIT_ID, 'SABRE' AS GIT_GDS, 'SABRE' AS GIT_INTERFACE, dbo.SabreM0.UPLOADNO AS GIT_UPLOADNO, SUBSTRING(dbo.SabreM0.IU0MDT, 1, 2) " _
            & "                      + '/' + SUBSTRING(dbo.SabreM0.IU0MDT, 3, 3) + '/' + '" & Year(Date) & "' AS GIT_PNRDATE, dbo.SabreM0.IU0PNR AS GIT_PNR, dbo.SabreM1.IU1PNM AS GIT_LASTNAME, " _
            & "                      '' AS GIT_FIRSTNAME, dbo.SabreM0.IU0IAG AS GIT_BOOKINGAGENT, dbo.SabreM5.IU5VR4 AS TKTNBRGIT_TICKETNUMBER, dbo.SabreM0.FNAME AS GIT_FILENAME, " _
            & "                       dbo.SabreM0.GDSAutoFailed AS GIT_GDSAUTOFAILED, dbo.SabreM0.IU0PCC AS GIT_PCC, '" & vDateLastModified_String & "' AS GIT_FILEISSUEDATE, " _
            & "                      '" & SkipChars(GIT_CUSTOMERUSERCODE_String) & "' AS GIT_CUSTOMERUSERCODE, '" & SkipChars(GIT_PENWAIT_String) & "' AS GIT_PENWAIT, " _
            & "                      '" & vFileDateCreated_String & "' AS GIT_FILECREATEDDATE, " _
            & "                      CASE " _
            & "                       WHEN dbo.SabreM2.IU2SN2 IS NULL OR LTRIM(RTRIM(dbo.SabreM2.IU2SN2)) = '' THEN dbo.SabreM0.IU0IAG " _
            & "                       ELSE dbo.SabreM2.IU2SN2 " _
            & "                      END AS GIT_ISSUEDBY, '" & GIT_PENLINEBRID_String & "' AS GIT_PENLINEBRID "
        vSQL_String = vSQL_String & "" _
            & "FROM         dbo.SabreM0 INNER JOIN " _
            & "                      dbo.SabreM1 ON dbo.SabreM0.UPLOADNO = dbo.SabreM1.UPLOADNO LEFT OUTER JOIN " _
            & "                      dbo.SabreM5 ON CAST(CASE WHEN isnumeric(LEFT(SabreM5.IU5MIN, 2)) = 1 THEN LEFT(SabreM5.IU5MIN, 2) ELSE SabreM5.IU5PTY END AS int) " _
            & "                      = dbo.SabreM1.IU1PNO AND dbo.SabreM5.UPLOADNO = dbo.SabreM1.UPLOADNO LEFT OUTER JOIN " _
            & "                      dbo.SabreM2 ON dbo.SabreM5.UPLOADNO = dbo.SabreM2.UPLOADNO AND left(dbo.SabreM5.IU5VR4,10) = dbo.SabreM2.IU2TNO " _
            & "WHERE     (dbo.SabreM0.UPLOADNO = " & UploadNo & ") AND (dbo.SabreM0.IU0TYP = '1' OR " _
            & "                      dbo.SabreM0.IU0TYP = '2' OR " _
            & "                      dbo.SabreM0.IU0TYP = '3' OR " _
            & "                      dbo.SabreM0.IU0TYP = '5' OR " _
            & "                      dbo.SabreM0.IU0TYP = 'D' OR " _
            & "                      dbo.SabreM0.IU0TYP = 'A' OR " _
            & "                      dbo.SabreM0.IU0TYP = 'B' OR " _
            & "                      dbo.SabreM0.IU0TYP = 'C') AND (dbo.SabreM5.IU5VR4 IS NULL OR " _
            & "                      dbo.SabreM5.IU5VR4 <> 'BLK' AND dbo.SabreM5.IU5VR4 <> 'DIS' AND dbo.SabreM5.IU5VR4 <> 'TBR')"
        'by Abhi on 21-May-2019 for caseid 9497 Show issued by user in the GDS Tray
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        'by Abhi on 05-Aug-2016 for caseid 6651 Splitting a PNR in case of bulk PNR
        'by Abhi on 07-Nov-2013 for caseid 3512 PenGDS Sabre Duplication
        'by Abhi on 19-Oct-2013 for caseid 3465 Sabre file loading issue pick TRANSACTION TYPE=A,B,C
        dbCompany.Execute vSQL_String
        FMain.stbUpload.Panels(1).Text = "File moving to Destination..."
        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
    End If
    'by Abhi on 09-Nov-2012 for caseid 2368 Sabre pnr missing checking sequence changed
    'commented by Abhi on 14-Jan-2013 for caseid 2786 PenGDS Sabre PNR missing - SabreM1 row not found. we need to skip this file
    'by Abhi on 08-Mar-2013 for caseid 2958 PenGDS Sabre pnr file not refeshing
'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
'    If Val(getFromTable("dbo.SabreM1", "UPLOADNO", "UPLOADNO", UploadNo)) <> UploadNo Then
'    'If Val(getFromTable("dbo.SabreM1", "UPLOADNO", "UPLOADNO", UploadNo)) = UploadNo Then
'        'commented by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
'        ''by Abhi on 25-Mar-2013 for caseid 2786 PenGDS Sabre PNR missing - SabreM1 row not found. we need to skip this file
'        'If Dir(txtDirName.Text & "\Error M1 Missing\", vbDirectory) = "" Then
'        '    MkDir txtDirName.Text & "\Error M1 Missing\"
'        'End If
'        'commented by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
'        fsObj.CopyFile txtFileName.Text, txtDirName.Text & "\Error M1 Missing\", True
''commented by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
''On Error Resume Next
''commented by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
'        fsObj.DeleteFile txtFileName.Text, True
''commented by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
''On Error GoTo 0
''commented by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
'        Err.Raise vbObjectError + 513, "Sabre", "PenGDSErr(Sabre PNR missing - SabreM1 row not found)"
'    End If
'    If CopyStatus = False Then
'        'fsObj.MoveFile txtFileName.Text, txtRelLocation.Text & "\"
'        fsObj.CopyFile txtFileName.Text, txtRelLocation.Text & "\", True
''commented by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
''On Error Resume Next
''commented by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
'        fsObj.DeleteFile txtFileName.Text, True
'        'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
'        INIWrite App.Path & "\PenSEARCH.ini", "Sabre", "LUFPNR", LUFPNR_String
'        INIWrite App.Path & "\PenSEARCH.ini", "Sabre", "LUFDate", DateFormat(Date)
'        INIWrite App.Path & "\PenSEARCH.ini", "Sabre", "LUFTime", TimeFormat12HRS(time) & "(" & TimeFormat(time) & ")"
'        'by Abhi on 18-May-2011 for caseid 1757 Added Events for GDSAuto
'        Call EventLog("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & LUFPNR_String & " - " & FMain.stbUpload.Panels(3).Text)
'        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
'        dbCompany.CommitTrans
'        'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
'    End If
    'by Abhi on 07-Nov-2013 for caseid 3512 PenGDS Sabre Duplication for Include Itinerary Only (DIT ENTRIES)
    'If Val(getFromTable("dbo.SabreM1", "UPLOADNO", "UPLOADNO", UploadNo)) <> UploadNo Then
    '    'fsObj.CopyFile txtFileName.Text, txtDirName.Text & "\Error M1 Missing\", True
    '    'fsObj.DeleteFile txtFileName.Text, True
    '    'Err.Raise vbObjectError + 513, "Sabre", "PenGDSErr(Sabre PNR missing - SabreM1 row not found)"
    '    fsObj.CopyFile txtFileName.Text, txtRelLocation.Text & "\", True
    '    fsObj.DeleteFile txtFileName.Text, True
    '    dbCompany.RollbackTrans
    'Else
    'by Abhi on 07-Nov-2013 for caseid 3512 PenGDS Sabre Duplication for Include Itinerary Only (DIT ENTRIES)
        If CopyStatus = False Then
            fsObj.CopyFile txtFileName.Text, txtRelLocation.Text & "\", True
            fsObj.DeleteFile txtFileName.Text, True
            'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
            INIWrite App.Path & "\PenSEARCH.ini", "Sabre", "LUFPNR", LUFPNR_String
            INIWrite App.Path & "\PenSEARCH.ini", "Sabre", "LUFDate", DateFormat(Date)
            INIWrite App.Path & "\PenSEARCH.ini", "Sabre", "LUFTime", TimeFormat12HRS(time) & "(" & TimeFormat(time) & ")"
            'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
            'Call EventLog("#PenGDS# " & FMain.stbUpload.Panels(2).Text & " - " & LUFPNR_String & " - " & FMain.stbUpload.Panels(3).Text)
            Call EventLog("#PenGDS# " & PadR(FMain.stbUpload.Panels(2).Text, 9) & " - " & PadR(LUFPNR_String, 6) & " - " & PadR(FMain.stbUpload.Panels(3).Text, 27))
            'by Abhi on 12-Jun-2015 for caseid 5313 PenGDS Error Multiple-step operation generated errors Check each status value
            'dbCompany.CommitTrans
            'by Abhi on 06-Dec-2018 for caseid 9584 Sabre Error: -2147168242 - No transaction is active
            If PENErr_BeginTrans = True Then
                dbCompany.CommitTrans
            End If
            'by Abhi on 06-Dec-2018 for caseid 9584 Sabre Error: -2147168242 - No transaction is active
            'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
            PENErr_BeginTrans = False
            'by Abhi on 16-Dec-2014 for caseid 4827 Warning(Sabre) in PengDS No transaction is active
        End If
    'by Abhi on 07-Nov-2013 for caseid 3512 PenGDS Sabre Duplication for Include Itinerary Only (DIT ENTRIES)
    'End If
    'by Abhi on 07-Nov-2013 for caseid 3512 PenGDS Sabre Duplication for Include Itinerary Only (DIT ENTRIES)
'by Abhi on 08-Oct-2013 for caseid 3373 PNR showing locked by admin
    Exit Sub
Note:
   Me.Hide
Exit Sub
PENErr:
    'by Abhi on 14-Feb-2014 for caseid 3741 Pengds -Galileo file stuck Incorrect syntax near ARCY
    If (tsObj Is Nothing) = False Then
        tsObj.Close
    End If
    'by Abhi on 14-Feb-2014 for caseid 3741 Pengds -Galileo file stuck Incorrect syntax near ARCY
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'If PENErr() = True Then GoTo DeadlockRETRY
    If PENErr() = True Then Resume DeadlockRETRY
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
End Sub
Public Sub SelectDetails(ln As String, id As String)

Dim Query As String
Dim lnTemp As String

Static i As Integer
Static j As Integer
Static k As Integer


Static M3_flag As Boolean
Static M2_flag As Boolean
Static M1_flag As Boolean
Static M0_flag As Boolean

Static M3_UploadNo As Long
Static M2_UploadNo As Long
Static M1_UploadNo As Long
Static M0_UploadNo As Long

Dim TempPEN, TempPENSplited
Dim TempPENi As Long, TempPENCount As Long
Dim TempPENPNO As Long
Dim CollPenSplited As New Collection
'by Abhi on 23-Mar-2011 for caseid 1648 Sabre - Passenger line connection to ticket line - Logic changed to IU5PTY to IU1M5C
Static vIU1PNO_Long As Long
'by Abhi on 21-Jun-2014 for caseid 4189 Sabre M5 Line for TASF with @ is not saving in gds table
Dim vAtSignPosition_Long As Long
'by Abhi on 21-Jun-2014 for caseid 4189 Sabre M5 Line for TASF with @ is not saving in gds table
'by Abhi on 08-Sep-2014 for caseid 4496 Sabre Airline Fee should be taken from MA Record
Dim aMA() As String
'by Abhi on 08-Sep-2014 for caseid 4496 Sabre Airline Fee should be taken from MA Record
'by Abhi on 10-Mar-2015 for caseid 5016 Sabre-Delivery address,phone number and email issue
Dim aMF() As String
'by Abhi on 10-Mar-2015 for caseid 5016 Sabre-Delivery address,phone number and email issue
'by Abhi on 06-Apr-2016 for caseid 6217 penfare line for sabre with supplier code
Dim vPENFARESELL_Currency As Currency
'by Abhi on 06-Apr-2016 for caseid 6217 penfare line for sabre with supplier code
'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking
Dim vPENFAREPASSTYPE_String As String
'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking

    If id = "AA" And Mid(ln, 3, 1) <> "/" And Len(ln) > 100 Then
        Preid_String = id
        M0_flag = True
        M0_UploadNo = UploadNo
        Query = "Insert into SabreHeader(UPLOADNO,Header1,Header2,Header3,Header4) " & _
                " Values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 1, 2)) & "','" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "', " & _
                " '" & Trim(fnMidValue(ln, 8, 4)) & "')"
        dbCompany.Execute Query
        If Trim(fnMidValue(ln, 14, 1)) <> 5 Then
            'by Abhi on 07-Jan-2012 for caseid 2009 Sabre file picking wrong customercode as 0
            Query = "Insert into SabreM0(UPLOADNO,IU0TYP,IU0VER,IU0DKN,IU0CTJ,IU0CAP,IU0PIV,IU0RRC,IU0QUE,IU0CDK,IU0IVN,IU0ATC,IU0PNR,IU0PNL,IU0OPT,IU0F00, " & _
                    " IU0F01,IU0F02,IU0F03,IU0F04,IU0F05,IU0F06,IU0F07,IU0F08,IU0F09,IU0F0A,IU0F0B,IU0F0C,IU0F0D,IU0F0E,IU0F0F,IU0ATB,IU0DUP,IU0PCC,IU0IDC,IU0IAG, " & _
                    " IU0LIN,IU0RPR,IU0PDT,IU0TIM,IU0IS4,IU0IS1,IU0IS3,IU0TCO,IU0IDB,IU0DEP,IU0ORG,IU0ONM,IU0DST,IU0DNM,IU0NM1,IU0NM2,IU0NM3,IU0NM4,IU0NM5,IU0NM6, " & _
                    " IU0NM7,IU0NM8,IU0NM9,IU0NMA,IU0ADC,IU0PHC,IU0TYM,IU0MDT,FNAME,GDSAutoFailed,IU0BL1,IU0BL2) " & _
                    " Values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 14, 1)) & "','" & Trim(fnMidValue(ln, 15, 2)) & "','" & SkipChars(fnMidValue(ln, 17, 10)) & "','" & Trim(fnMidValue(ln, 27, 4)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 31, 1)) & "','" & Trim(fnMidValue(ln, 32, 1)) & "','" & Trim(fnMidValue(ln, 33, 2)) & "','" & Trim(fnMidValue(ln, 35, 1)) & "','" & Trim(fnMidValue(ln, 36, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 37, 7)) & "','" & Trim(fnMidValue(ln, 44, 10)) & "','" & Trim(fnMidValue(ln, 54, 8)) & "','" & Trim(fnMidValue(ln, 62, 8)) & "','" & Trim(fnMidValue(ln, 70, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 71, 1)) & "','" & Trim(fnMidValue(ln, 72, 1)) & "','" & Trim(fnMidValue(ln, 73, 1)) & "','" & Trim(fnMidValue(ln, 74, 1)) & "','" & Trim(fnMidValue(ln, 75, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 76, 1)) & "','" & Trim(fnMidValue(ln, 77, 1)) & "','" & Trim(fnMidValue(ln, 78, 1)) & "','" & Trim(fnMidValue(ln, 79, 1)) & "','" & Trim(fnMidValue(ln, 80, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 81, 1)) & "','" & Trim(fnMidValue(ln, 82, 1)) & "','" & Trim(fnMidValue(ln, 83, 1)) & "','" & Trim(fnMidValue(ln, 84, 1)) & "','" & Trim(fnMidValue(ln, 85, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 86, 1)) & "','" & Trim(fnMidValue(ln, 87, 1)) & "','" & Trim(fnMidValue(ln, 88, 1)) & "','" & Trim(fnMidValue(ln, 89, 5)) & "','" & Trim(fnMidValue(ln, 94, 2)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 96, 3)) & "','" & Trim(fnMidValue(ln, 99, 8)) & "','" & Trim(fnMidValue(ln, 107, 10)) & "','" & Trim(fnMidValue(ln, 117, 5)) & "','" & Trim(fnMidValue(ln, 122, 5)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 127, 5)) & "','" & Trim(fnMidValue(ln, 132, 2)) & "','" & Trim(fnMidValue(ln, 134, 3)) & "','" & Trim(fnMidValue(ln, 137, 2)) & "','" & Trim(fnMidValue(ln, 139, 3)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 142, 5)) & "','" & Trim(fnMidValue(ln, 147, 3)) & "','" & Trim(fnMidValue(ln, 150, 17)) & "','" & Trim(fnMidValue(ln, 167, 3)) & "','" & Trim(fnMidValue(ln, 170, 17)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 187, 3)) & "','" & Trim(fnMidValue(ln, 190, 3)) & "','" & Trim(fnMidValue(ln, 193, 3)) & "','" & Trim(fnMidValue(ln, 196, 3)) & "','" & Trim(fnMidValue(ln, 199, 3)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 202, 3)) & "','" & Trim(fnMidValue(ln, 205, 3)) & "','" & Trim(fnMidValue(ln, 208, 3)) & "','" & Trim(fnMidValue(ln, 211, 3)) & "','" & Trim(fnMidValue(ln, 214, 3)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 217, 2)) & "','" & Trim(fnMidValue(ln, 219, 2)) & "','" & Trim(fnMidValue(ln, 221, 5)) & "','" & Trim(fnMidValue(ln, 226, 5)) & "','" & FNAME & "',0,'','')"
            'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
            LUFPNR_String = Trim(fnMidValue(ln, 54, 8))
        Else
            Query = "Insert into SabreM0(UPLOADNO,IU0TYP,IU0VER,IU0DKN,IU0TKI,IU0CNT,IU0PR1,IU0DUP,IU0SAC,FNAME,GDSAutoFailed,IU0BL1,IU0BL2) " & _
                    " Values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 14, 1)) & "','" & Trim(fnMidValue(ln, 15, 2)) & "','" & Trim(fnMidValue(ln, 17, 10)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 27, 14)) & "','" & Trim(fnMidValue(ln, 41, 2)) & "','" & Trim(fnMidValue(ln, 43, 8)) & "','" & Trim(fnMidValue(ln, 51, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 52, 15)) & "','" & FNAME & "',0,'','')"
        End If
        dbCompany.Execute Query
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        If Trim(Right(fnMidValue(ln, 17, 10), 7)) <> "" Then
            GIT_CUSTOMERUSERCODE_String = Right(fnMidValue(ln, 17, 10), 7)
        End If
        'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
        Exit Sub
    ElseIf id = "M1" And Mid(ln, 3, 1) <> "/" And Len(ln) > 100 Then
        Preid_String = id
        M1_flag = True
        M1_UploadNo = UploadNo
        j = 0
        'by Abhi on 23-Mar-2011 for caseid 1648 Sabre - Passenger line connection to ticket line - Logic changed to IU5PTY to IU1M5C
        vIU1PNO_Long = Val(Trim(fnMidValue(ln, 3, 2)))
        Query = "Insert into SabreM1(UPLOADNO,IU1PNO,IU1PNM,IU1PRK,IU1IND,IU1AVN,IU1MLV,IU1NM3,IU1TKT,IU1SSM,IU1NM5,IU1NM7, " & _
                " IU1NM8,IU1NM9,IU1NMA) " & _
                " Values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & ",'" & Trim(fnMidValue(ln, 5, 64)) & "','" & Trim(fnMidValue(ln, 69, 30)) & "', " & _
                " '" & Trim(fnMidValue(ln, 99, 1)) & "','" & Trim(fnMidValue(ln, 100, 20)) & "','" & Trim(fnMidValue(ln, 120, 5)) & "', " & _
                " '" & Trim(fnMidValue(ln, 125, 2)) & "','" & Trim(fnMidValue(ln, 127, 1)) & "','" & Trim(fnMidValue(ln, 128, 1)) & "'," & Val(fnMidValue(ln, 129, 2)) & ", " & _
                " '" & Trim(fnMidValue(ln, 131, 2)) & "','" & Trim(fnMidValue(ln, 133, 2)) & "','" & Trim(fnMidValue(ln, 135, 2)) & "','" & Trim(fnMidValue(ln, 137, 2)) & "')"
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M2" And Mid(ln, 3, 1) <> "/" Then
        Preid_String = id
        M2_flag = True
        M2_UploadNo = UploadNo
        i = 0
        Query = "Insert into SabreM2(UPLOADNO,IU2PNO,IU2PTY,IU2TCN,IU2INT,IU2SSI,IU2IND,IU2AP1,IU2AP2,IU2AP3,IU2AP4,IU2AP5,IU2AP6,IU2AP7,IU2AP8,IU2AP9,IU2APA,IU2APB,IU2APZ,IU2FSN,IU2FCC,IU2FAR,IU2T1S,IU2TX1,IU2ID1,IU2T2S, " & _
            " IU2TX2,IU2ID2,IU2T3S,IU2TX3,IU2ID3,IU2TFS,IU2TFC,IU2TFR,IU2PEN,IU2KXP,IU2SPX,IU2EPS,IU2EFC,IU2EFR,IU2PCT,IU2CSN,IU2COM,IU2NAS,IU2NET,IU2CDC,IU2CTT,IU2FTF,IU2TAT,IU2TUR,IU2TPC,IU2MIT,IU2CCB,IU2APC, " & _
            " IU2SN4,IU2SN3,IU2SN2,IU2PN4,IU2PN3,IU2PN2,IU2AVI,IU2ATH,IU2SPR,IU2EXP,IU2EXT,IU2M4C,IU2M6C,IU2VAL,IU2TNO,IU2DNO)" & _
            " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & ",'" & Trim(fnMidValue(ln, 5, 3)) & "','" & Trim(fnMidValue(ln, 8, 11)) & "','" & Trim(fnMidValue(ln, 19, 1)) & "','" & Trim(fnMidValue(ln, 20, 1)) & "','" & Trim(fnMidValue(ln, 21, 1)) & "', " & _
            " '" & Trim(fnMidValue(ln, 22, 1)) & "','" & Trim(fnMidValue(ln, 23, 1)) & "','" & Trim(fnMidValue(ln, 24, 1)) & "','" & Trim(fnMidValue(ln, 25, 1)) & "','" & Trim(fnMidValue(ln, 26, 1)) & "','" & Trim(fnMidValue(ln, 27, 1)) & "', " & _
            " '" & Trim(fnMidValue(ln, 28, 1)) & "','" & Trim(fnMidValue(ln, 29, 1)) & "','" & Trim(fnMidValue(ln, 30, 1)) & "','" & Trim(fnMidValue(ln, 31, 1)) & "','" & Trim(fnMidValue(ln, 32, 1)) & "','" & Trim(fnMidValue(ln, 33, 1)) & "', " & _
            " '" & Trim(fnMidValue(ln, 34, 1)) & "','" & Trim(fnMidValue(ln, 35, 3)) & "'," & Val(Trim(fnMidValue(ln, 38, 8))) & ",'" & Trim(fnMidValue(ln, 46, 1)) & "'," & Val(Trim(fnMidValue(ln, 47, 7))) & ",'" & Trim(fnMidValue(ln, 54, 2)) & "', " & _
            " '" & Trim(fnMidValue(ln, 56, 1)) & "'," & Val(Trim(fnMidValue(ln, 57, 7))) & ",'" & Trim(fnMidValue(ln, 64, 2)) & "','" & Trim(fnMidValue(ln, 66, 1)) & "'," & Val(Trim(fnMidValue(ln, 67, 7))) & ",'" & Trim(fnMidValue(ln, 74, 2)) & "', " & _
            " '" & Trim(fnMidValue(ln, 76, 1)) & "','" & Trim(fnMidValue(ln, 77, 3)) & "'," & Val(Trim(fnMidValue(ln, 80, 8))) & "," & Val(Trim(fnMidValue(ln, 88, 11))) & "," & Val(Trim(fnMidValue(ln, 99, 11))) & ",'" & Trim(fnMidValue(ln, 110, 6)) & "', " & _
            " '" & Trim(fnMidValue(ln, 116, 1)) & "','" & Trim(fnMidValue(ln, 117, 3)) & "'," & Val(Trim(fnMidValue(ln, 120, 8))) & ",'" & Trim(fnMidValue(ln, 128, 8)) & "','" & Trim(fnMidValue(ln, 136, 1)) & "'," & Val(Trim(fnMidValue(ln, 137, 8))) & ", " & _
            " '" & Trim(fnMidValue(ln, 145, 1)) & "'," & Val(Trim(fnMidValue(ln, 146, 8))) & ",'" & Trim(fnMidValue(ln, 154, 1)) & "'," & Val(Trim(fnMidValue(ln, 155, 1))) & ",'" & Trim(fnMidValue(ln, 156, 1)) & "','" & Trim(fnMidValue(ln, 157, 8)) & "', " & _
            " '" & Trim(fnMidValue(ln, 165, 15)) & "','" & Trim(fnMidValue(ln, 180, 4)) & "','" & Trim(fnMidValue(ln, 184, 1)) & "','" & Trim(fnMidValue(ln, 185, 1)) & "','" & Trim(fnMidValue(ln, 186, 1)) & "','" & Trim(fnMidValue(ln, 187, 5)) & "', " & _
            " '" & Trim(fnMidValue(ln, 192, 2)) & "','" & Trim(fnMidValue(ln, 194, 3)) & "','" & Trim(fnMidValue(ln, 197, 5)) & "','" & Trim(fnMidValue(ln, 202, 2)) & "','" & Trim(fnMidValue(ln, 204, 3)) & "','" & Trim(fnMidValue(ln, 207, 1)) & "', " & _
            " '" & Trim(fnMidValue(ln, 208, 9)) & "','" & Trim(fnMidValue(ln, 217, 5)) & "','" & Trim(fnMidValue(ln, 222, 4)) & "','" & Trim(fnMidValue(ln, 226, 2)) & "','" & Trim(fnMidValue(ln, 228, 2)) & "','" & Trim(fnMidValue(ln, 230, 2)) & "', " & _
            " '" & Trim(fnMidValue(ln, 232, 2)) & "','" & Trim(fnMidValue(ln, 234, 10)) & "','" & Trim(fnMidValue(ln, 244, 1)) & "')"
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M3" And Mid(ln, 3, 1) <> "/" Then
        Preid_String = id
        M3_flag = True
        M3_UploadNo = UploadNo
        k = 0
        If Trim(fnMidValue(ln, 5, 1)) = 1 Then
            'by Abhi on 18-May-2009 for vlocator
            'Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3BPI,IU3DCC,IU3DCY,IU3ACC,IU3ACY,IU3CAR,IU3FLT,IU3CLS,IU3DTM,IU3ATM,IU3ELT,IU3MLI,IU3SUP,IU3DCH,IU3NOS,IU3SCC,IU3CRT,IU3EQP,IU3ARM,IU3AVM,IU3SCT,IU3RTM,IU3GAT,IU3TMA,IU3GTA,IU3GAR,IU3RTM_1,IU3COG,IU3CRN,IU3TKT,IU3MCT,IU3SPX) " & _
                    " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & "," & Val(Trim(fnMidValue(ln, 5, 1))) & "," & Val(Trim(fnMidValue(ln, 6, 1))) & "," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "','" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 19, 3)) & "','" & Trim(fnMidValue(ln, 22, 17)) & "','" & Trim(fnMidValue(ln, 39, 3)) & "','" & Trim(fnMidValue(ln, 42, 17)) & "','" & Trim(fnMidValue(ln, 59, 2)) & "','" & Trim(fnMidValue(ln, 61, 5)) & "','" & Trim(fnMidValue(ln, 66, 2)) & "','" & Trim(fnMidValue(ln, 68, 5)) & "','" & Trim(fnMidValue(ln, 73, 5)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 78, 8)) & "','" & Trim(fnMidValue(ln, 86, 4)) & "','" & Trim(fnMidValue(ln, 90, 1)) & "','" & Trim(fnMidValue(ln, 91, 1)) & "','" & Trim(fnMidValue(ln, 92, 1)) & "','" & Trim(fnMidValue(ln, 93, 18)) & "','" & Trim(fnMidValue(ln, 111, 2)) & "','" & Trim(fnMidValue(ln, 113, 3)) & "','" & Trim(fnMidValue(ln, 116, 6)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 122, 6)) & "','" & Trim(fnMidValue(ln, 128, 2)) & "','" & Trim(fnMidValue(ln, 130, 26)) & "','" & Trim(fnMidValue(ln, 156, 4)) & "','" & Trim(fnMidValue(ln, 160, 26)) & "','" & Trim(fnMidValue(ln, 186, 5)) & "','" & Trim(fnMidValue(ln, 191, 4)) & "','" & Trim(fnMidValue(ln, 195, 5)) & "','" & Trim(fnMidValue(ln, 200, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 201, 37)) & "','" & Trim(fnMidValue(ln, 238, 1)) & "','" & Trim(fnMidValue(ln, 239, 2)) & "','" & Trim(fnMidValue(ln, 241, 12)) & "')"
            'commented and rewrtten by Abhi 20-May-2009 for new documenation removed IU3GTA, IU3SPX
            'Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3BPI,IU3DCC,IU3DCY,IU3ACC,IU3ACY,IU3CAR,IU3FLT,IU3CLS,IU3DTM,IU3ATM,IU3ELT,IU3MLI,IU3SUP,IU3DCH,IU3NOS,IU3SCC,IU3CRT,IU3EQP,IU3ARM,IU3AVM,IU3SCT,IU3RTM,IU3GAT,IU3TMA,IU3GTA,IU3GAR,IU3RTM_1,IU3COG,IU3CRN,IU3TKT,IU3MCT,IU3SPX) " & _
                    " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & "," & Val(Trim(fnMidValue(ln, 5, 1))) & "," & Val(Trim(fnMidValue(ln, 6, 1))) & "," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "','" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 19, 3)) & "','" & Trim(fnMidValue(ln, 22, 17)) & "','" & Trim(fnMidValue(ln, 39, 3)) & "','" & Trim(fnMidValue(ln, 42, 17)) & "','" & Trim(fnMidValue(ln, 59, 2)) & "','" & Trim(fnMidValue(ln, 61, 5)) & "','" & Trim(fnMidValue(ln, 66, 2)) & "','" & Trim(fnMidValue(ln, 68, 5)) & "','" & Trim(fnMidValue(ln, 73, 5)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 78, 8)) & "','" & Trim(fnMidValue(ln, 86, 4)) & "','" & Trim(fnMidValue(ln, 90, 1)) & "','" & Trim(fnMidValue(ln, 91, 1)) & "','" & Trim(fnMidValue(ln, 92, 1)) & "','" & Trim(fnMidValue(ln, 93, 18)) & "','" & Trim(fnMidValue(ln, 111, 2)) & "','" & Trim(fnMidValue(ln, 113, 3)) & "','" & Trim(fnMidValue(ln, 116, 6)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 122, 6)) & "','" & Trim(fnMidValue(ln, 128, 2)) & "','" & Trim(fnMidValue(ln, 130, 26)) & "','" & Trim(fnMidValue(ln, 156, 4)) & "','" & Trim(fnMidValue(ln, 160, 26)) & "','" & Trim(fnMidValue(ln, 186, 5)) & "','" & Trim(fnMidValue(ln, 191, 4)) & "','" & Trim(fnMidValue(ln, 195, 5)) & "','" & Trim(fnMidValue(ln, 200, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 201, 37)) & "','" & Trim(fnMidValue(ln, 238, 1)) & "','" & Trim(fnMidValue(ln, 239, 1)) & "','" & Trim(fnMidValue(ln, 240, 12)) & "')"
            'by Abhi on 24-Apr-2011 for caseid 2126 Sabre Rail Segment 2nd stage
            'Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3BPI,IU3DCC,IU3DCY,IU3ACC,IU3ACY,IU3CAR,IU3FLT,IU3CLS,IU3DTM,IU3ATM,IU3ELT,IU3MLI,IU3SUP,IU3DCH,IU3NOS,IU3SCC,IU3CRT,IU3EQP,IU3ARM,IU3AVM,IU3SCT,IU3RTM,IU3GAT,IU3TMA,IU3GAR,IU3RTM_1,IU3COG,IU3CRN,IU3TKT,IU3MCT,IU3YER,IU3OAL) " & _
                    " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & "," & Val(Trim(fnMidValue(ln, 5, 1))) & "," & Val(Trim(fnMidValue(ln, 6, 1))) & "," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "','" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 19, 3)) & "','" & Trim(fnMidValue(ln, 22, 17)) & "','" & Trim(fnMidValue(ln, 39, 3)) & "','" & Trim(fnMidValue(ln, 42, 17)) & "','" & Trim(fnMidValue(ln, 59, 2)) & "','" & Trim(fnMidValue(ln, 61, 5)) & "','" & Trim(fnMidValue(ln, 66, 2)) & "','" & Trim(fnMidValue(ln, 68, 5)) & "','" & Trim(fnMidValue(ln, 73, 5)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 78, 8)) & "','" & Trim(fnMidValue(ln, 86, 4)) & "','" & Trim(fnMidValue(ln, 90, 1)) & "','" & Trim(fnMidValue(ln, 91, 1)) & "','" & Trim(fnMidValue(ln, 92, 1)) & "','" & Trim(fnMidValue(ln, 93, 18)) & "','" & Trim(fnMidValue(ln, 111, 2)) & "','" & Trim(fnMidValue(ln, 113, 3)) & "','" & Trim(fnMidValue(ln, 116, 6)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 122, 6)) & "','" & Trim(fnMidValue(ln, 128, 2)) & "','" & Trim(fnMidValue(ln, 130, 26)) & "','" & Trim(fnMidValue(ln, 156, 4)) & "','" & Trim(fnMidValue(ln, 160, 26)) & "','" & Trim(fnMidValue(ln, 186, 4)) & "','" & Trim(fnMidValue(ln, 190, 5)) & "','" & Trim(fnMidValue(ln, 195, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 196, 37)) & "','" & Trim(fnMidValue(ln, 233, 1)) & "','" & Trim(fnMidValue(ln, 234, 2)) & "','" & Trim(fnMidValue(ln, 236, 4)) & "','" & Trim(fnMidValue(ln, 240, 8)) & "')"
            'by Abhi on 05-Mar-2014 for caseid 3789 Sabre Accounting line to others new logic connect the M3 LINE from M5 line IU5VR1 to IU3LNK If the product code is IU3PRC<>1 or 3 or 6
            'Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3BPI,IU3DCC,IU3DCY,IU3ACC,IU3ACY,IU3CAR,IU3FLT,IU3CLS,IU3DTM,IU3ATM,IU3ELT,IU3MLI,IU3SUP,IU3DCH,IU3NOS,IU3SCC,IU3CRT,IU3EQP,IU3ARM,IU3AVM,IU3SCT,IU3RTM,IU3GAT,IU3TMA,IU3GAR,IU3RTM_1,IU3COG,IU3CRN,IU3TKT,IU3MCT,IU3YER,IU3OAL) " & _
                    " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & "," & Val(Trim(fnMidValue(ln, 5, 1))) & ",'" & SkipChars(fnMidValue(ln, 6, 1)) & "'," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "','" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 19, 3)) & "','" & Trim(fnMidValue(ln, 22, 17)) & "','" & Trim(fnMidValue(ln, 39, 3)) & "','" & Trim(fnMidValue(ln, 42, 17)) & "','" & Trim(fnMidValue(ln, 59, 2)) & "','" & Trim(fnMidValue(ln, 61, 5)) & "','" & Trim(fnMidValue(ln, 66, 2)) & "','" & Trim(fnMidValue(ln, 68, 5)) & "','" & Trim(fnMidValue(ln, 73, 5)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 78, 8)) & "','" & Trim(fnMidValue(ln, 86, 4)) & "','" & Trim(fnMidValue(ln, 90, 1)) & "','" & Trim(fnMidValue(ln, 91, 1)) & "','" & Trim(fnMidValue(ln, 92, 1)) & "','" & Trim(fnMidValue(ln, 93, 18)) & "','" & Trim(fnMidValue(ln, 111, 2)) & "','" & Trim(fnMidValue(ln, 113, 3)) & "','" & Trim(fnMidValue(ln, 116, 6)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 122, 6)) & "','" & Trim(fnMidValue(ln, 128, 2)) & "','" & Trim(fnMidValue(ln, 130, 26)) & "','" & Trim(fnMidValue(ln, 156, 4)) & "','" & Trim(fnMidValue(ln, 160, 26)) & "','" & Trim(fnMidValue(ln, 186, 4)) & "','" & Trim(fnMidValue(ln, 190, 5)) & "','" & Trim(fnMidValue(ln, 195, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 196, 37)) & "','" & Trim(fnMidValue(ln, 233, 1)) & "','" & Trim(fnMidValue(ln, 234, 2)) & "','" & Trim(fnMidValue(ln, 236, 4)) & "','" & Trim(fnMidValue(ln, 240, 8)) & "')"
            Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3BPI,IU3DCC,IU3DCY,IU3ACC,IU3ACY,IU3CAR,IU3FLT,IU3CLS,IU3DTM,IU3ATM,IU3ELT,IU3MLI,IU3SUP,IU3DCH,IU3NOS,IU3SCC,IU3CRT,IU3EQP,IU3ARM,IU3AVM,IU3SCT,IU3RTM,IU3GAT,IU3TMA,IU3GAR,IU3RTM_1,IU3COG,IU3CRN,IU3TKT,IU3MCT,IU3YER,IU3OAL) " & _
                    " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & ",'" & Trim(fnMidValue(ln, 5, 1)) & "','" & SkipChars(fnMidValue(ln, 6, 1)) & "'," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "','" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 19, 3)) & "','" & Trim(fnMidValue(ln, 22, 17)) & "','" & Trim(fnMidValue(ln, 39, 3)) & "','" & Trim(fnMidValue(ln, 42, 17)) & "','" & Trim(fnMidValue(ln, 59, 2)) & "','" & Trim(fnMidValue(ln, 61, 5)) & "','" & Trim(fnMidValue(ln, 66, 2)) & "','" & Trim(fnMidValue(ln, 68, 5)) & "','" & Trim(fnMidValue(ln, 73, 5)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 78, 8)) & "','" & Trim(fnMidValue(ln, 86, 4)) & "','" & Trim(fnMidValue(ln, 90, 1)) & "','" & Trim(fnMidValue(ln, 91, 1)) & "','" & Trim(fnMidValue(ln, 92, 1)) & "','" & Trim(fnMidValue(ln, 93, 18)) & "','" & Trim(fnMidValue(ln, 111, 2)) & "','" & Trim(fnMidValue(ln, 113, 3)) & "','" & Trim(fnMidValue(ln, 116, 6)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 122, 6)) & "','" & Trim(fnMidValue(ln, 128, 2)) & "','" & Trim(fnMidValue(ln, 130, 26)) & "','" & Trim(fnMidValue(ln, 156, 4)) & "','" & Trim(fnMidValue(ln, 160, 26)) & "','" & Trim(fnMidValue(ln, 186, 4)) & "','" & Trim(fnMidValue(ln, 190, 5)) & "','" & Trim(fnMidValue(ln, 195, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 196, 37)) & "','" & Trim(fnMidValue(ln, 233, 1)) & "','" & Trim(fnMidValue(ln, 234, 2)) & "','" & Trim(fnMidValue(ln, 236, 4)) & "','" & Trim(fnMidValue(ln, 240, 8)) & "')"
            'by Abhi on 05-Mar-2014 for caseid 3789 Sabre Accounting line to others new logic connect the M3 LINE from M5 line IU5VR1 to IU3LNK If the product code is IU3PRC<>1 or 3 or 6
        ElseIf Trim(fnMidValue(ln, 5, 1)) = 3 Then
            
            If fnMidValue(ln, 15, 3) = "HHL" Then
                'by Abhi on 24-Apr-2011 for caseid 2126 Sabre Rail Segment 2nd stage
                'by Abhi on 05-Mar-2014 for caseid 3789 Sabre Accounting line to others new logic connect the M3 LINE from M5 line IU5VR1 to IU3LNK If the product code is IU3PRC<>1 or 3 or 6
                'Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3NOS,IU3SCC,IU3CCN,IU3PTY,IU3CFN,IU3CAC,IU3TDG,IU3OUT,IU3PRP,IU3ITT,IU3NRM,IU3VR2) " & _
                          " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & "," & Val(Trim(fnMidValue(ln, 5, 1))) & ",'" & SkipChars(fnMidValue(ln, 6, 1)) & "'," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "','','','" & Trim(fnMidValue(ln, 19, 2)) & "','" & Trim(fnMidValue(ln, 21, 15)) & "','" & Trim(fnMidValue(ln, 36, 3)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 39, 2)) & "','" & Trim(fnMidValue(ln, 41, 14)) & "','" & Trim(fnMidValue(ln, 55, 6)) & "','" & Trim(fnMidValue(ln, 61, 32)) & "','" & Trim(fnMidValue(ln, 93, 11)) & "','" & Trim(fnMidValue(ln, 104, Len(ln))) & "')"
                Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3NOS,IU3SCC,IU3CCN,IU3PTY,IU3CFN,IU3CAC,IU3TDG,IU3OUT,IU3PRP,IU3ITT,IU3NRM,IU3VR2) " & _
                          " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & ",'" & Trim(fnMidValue(ln, 5, 1)) & "','" & SkipChars(fnMidValue(ln, 6, 1)) & "'," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "','','','" & Trim(fnMidValue(ln, 19, 2)) & "','" & Trim(fnMidValue(ln, 21, 15)) & "','" & Trim(fnMidValue(ln, 36, 3)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 39, 2)) & "','" & Trim(fnMidValue(ln, 41, 14)) & "','" & Trim(fnMidValue(ln, 55, 6)) & "','" & Trim(fnMidValue(ln, 61, 32)) & "','" & Trim(fnMidValue(ln, 93, 11)) & "','" & Trim(fnMidValue(ln, 104, Len(ln))) & "')"
                'by Abhi on 05-Mar-2014 for caseid 3789 Sabre Accounting line to others new logic connect the M3 LINE from M5 line IU5VR1 to IU3LNK If the product code is IU3PRC<>1 or 3 or 6
            Else
                'by Abhi on 24-Apr-2011 for caseid 2126 Sabre Rail Segment 2nd stage
                'by Abhi on 05-Mar-2014 for caseid 3789 Sabre Accounting line to others new logic connect the M3 LINE from M5 line IU5VR1 to IU3LNK If the product code is IU3PRC<>1 or 3 or 6
                'Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3CCN,IU3PTY,IU3CFN,IU3VR4) " & _
                          " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & "," & Val(Trim(fnMidValue(ln, 5, 1))) & ",'" & SkipChars(fnMidValue(ln, 6, 1)) & "'," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "','" & Trim(fnMidValue(ln, 19, 2)) & "','" & Trim(fnMidValue(ln, 21, 15)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 36, Len(ln))) & "')"
                Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3CCN,IU3PTY,IU3CFN,IU3VR4) " & _
                          " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & ",'" & Trim(fnMidValue(ln, 5, 1)) & "','" & SkipChars(fnMidValue(ln, 6, 1)) & "'," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "','" & Trim(fnMidValue(ln, 19, 2)) & "','" & Trim(fnMidValue(ln, 21, 15)) & "', " & _
                          " '" & Trim(fnMidValue(ln, 36, Len(ln))) & "')"
                'by Abhi on 05-Mar-2014 for caseid 3789 Sabre Accounting line to others new logic connect the M3 LINE from M5 line IU5VR1 to IU3LNK If the product code is IU3PRC<>1 or 3 or 6
            End If
        
        Else
            'by Abhi on 24-Apr-2011 for caseid 2126 Sabre Rail Segment 2nd stage
            'by Abhi on 05-Mar-2014 for caseid 3789 Sabre Accounting line to others new logic connect the M3 LINE from M5 line IU5VR1 to IU3LNK If the product code is IU3PRC<>1 or 3 or 6
            'Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3CCN,IU3PTY,IU3CFN,IU3VR4) " & _
                    " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & "," & Val(Trim(fnMidValue(ln, 5, 1))) & ",'" & SkipChars(fnMidValue(ln, 6, 1)) & "'," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "','" & Trim(fnMidValue(ln, 19, 2)) & "','" & Trim(fnMidValue(ln, 21, 15)) & "','" & Trim(fnMidValue(ln, 36, Len(ln))) & "')"
            Query = "Insert into SabreM3(UPLOADNO,IU3ITN,IU3PRC,IU3LNK,IU3CRL,IU3AAC,IU3DDT,IU3PC2,IU3CCN,IU3PTY,IU3CFN,IU3VR4) " & _
                    " values(" & UploadNo & "," & Val(Trim(fnMidValue(ln, 3, 2))) & ",'" & Trim(fnMidValue(ln, 5, 1)) & "','" & SkipChars(fnMidValue(ln, 6, 1)) & "'," & Val(Trim(fnMidValue(ln, 7, 1))) & ",'" & Trim(fnMidValue(ln, 8, 2)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 15, 3)) & "','" & Trim(fnMidValue(ln, 18, 1)) & "','" & Trim(fnMidValue(ln, 19, 2)) & "','" & Trim(fnMidValue(ln, 21, 15)) & "','" & Trim(fnMidValue(ln, 36, Len(ln))) & "')"
            'by Abhi on 05-Mar-2014 for caseid 3789 Sabre Accounting line to others new logic connect the M3 LINE from M5 line IU5VR1 to IU3LNK If the product code is IU3PRC<>1 or 3 or 6
        End If
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M4" And Mid(ln, 3, 1) <> "/" Then
        Preid_String = id
        Query = "Insert into SabreM4(UPLOADNO,IU4SEG,IU4TYP,IU4CNI,IU4ETP,IU4NVB,IU4NVA,IU4AAC,IU4AWL,IU4FBS,IU4TDS,IU4ACL,IU4AMT,IU4ETK,IU4FB2,IU4TD2,IU4CUR,IU4SP2) " & _
                " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "','" & Trim(fnMidValue(ln, 8, 1)) & "', " & _
                " '" & Trim(fnMidValue(ln, 9, 1)) & "','" & Trim(fnMidValue(ln, 10, 5)) & "','" & Trim(fnMidValue(ln, 15, 5)) & "','" & Trim(fnMidValue(ln, 20, 2)) & "', " & _
                " '" & Trim(fnMidValue(ln, 22, 3)) & "','" & Trim(fnMidValue(ln, 25, 7)) & "','" & Trim(fnMidValue(ln, 32, 6)) & "','" & Trim(fnMidValue(ln, 38, 2)) & "', " & _
                " '" & Trim(fnMidValue(ln, 40, 8)) & "','" & Trim(fnMidValue(ln, 48, 1)) & "','" & Trim(fnMidValue(ln, 49, 12)) & "','" & Trim(fnMidValue(ln, 61, 10)) & "', " & _
                " '" & Trim(fnMidValue(ln, 71, 3)) & "','" & Trim(fnMidValue(ln, 74, 13)) & "')"
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M5" And Mid(ln, 3, 1) <> "/" Then
        Preid_String = id
        lnTemp = Replace(ln, "#", "/")
        'by Abhi on 19-May-2014 for caseid 4044 Exchange Ticket value picking issue for Sabre
        'lnTemp = Left(lnTemp, 6) & Replace(Right(lnTemp, Len(lnTemp) - 6), "@", "/") 'Replace(lnTemp, "@", "/")
        'by Abhi on 21-Jun-2014 for caseid 4189 Sabre M5 Line for TASF with @ is not saving in gds table
        vAtSignPosition_Long = InStr(1, Left(lnTemp, 21), "@", vbTextCompare)
        If vAtSignPosition_Long > 0 Then
            lnTemp = Left(lnTemp, vAtSignPosition_Long - 1) & "/" & Right(lnTemp, Len(lnTemp) - vAtSignPosition_Long)
        End If
        'by Abhi on 21-Jun-2014 for caseid 4189 Sabre M5 Line for TASF with @ is not saving in gds table
        'by Abhi on 19-May-2014 for caseid 4044 Exchange Ticket value picking issue for Sabre
        lnTemp = Replace(lnTemp, "C0/", "")
        avpItems = Split(lnTemp, "/")
        'by Abhi on 08-Mar-2014 for caseid 3469 Sabre EMD Add one more field for EMD indicator as a sub field of IU5VR1
        'ReDim Preserve avpItems(8)
        ReDim Preserve avpItems(9)
        'by Abhi on 08-Mar-2014 for caseid 3469 Sabre EMD Add one more field for EMD indicator as a sub field of IU5VR1
'        Query = "Insert into SabreM5(UPLOADNO,IU5PTY,IU5MIN,IU5VR1,IU5VR2,IU5VR3,IU5VR4,IU5VR5,IU5VR6,IU5VR7,IU5VR8) " & _
                " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 2)) & "','" & Trim(fnMidValue(ln, 7, 1)) & "', " & _
                " '" & Trim(fnMidValue(ln, 8, 1)) & "','" & Trim(fnMidValue(ln, 9, 2)) & "','" & Trim(fnMidValue(ln, 12, 10)) & "', " & _
                " '" & Trim(fnMidValue(ln, 23, 7)) & "','" & Trim(fnMidValue(ln, 31, 8)) & "','" & Trim(fnMidValue(ln, 41, 6)) & "', " & _
                " '" & Trim(fnMidValue(ln, 48, Len(ln))) & "')"
        vlTCount = Val(avpItems(7))
        'by Abhi on 24-Apr-2011 for caseid 2126 Sabre Rail Segment 2nd stage
        'If Left(Trim(avpItems(1)), 3) = "BLK" Or Left(Trim(avpItems(1)), 3) = "FPT" Or Left(Trim(avpItems(1)), 3) = "DIS" Then
        'by Abhi on 15-May-2012 for caseid 2238 Sabre Hotel Not picking the transaction fee to others
        'If Left(Trim(avpItems(1)), 3) = "BLK" Or Left(Trim(avpItems(1)), 3) = "FPT" Or Left(Trim(avpItems(1)), 3) = "DIS" Or Left(Trim(avpItems(1)), 3) = "TKT" Then
        'by Abhi on 07-Jun-2012 for caseid 2253 Pick "VCH" also to m5 line for hotel for checking "Customer Pay Locally"
        If Left(Trim(avpItems(1)), 3) = "BLK" Or Left(Trim(avpItems(1)), 3) = "FPT" Or Left(Trim(avpItems(1)), 3) = "DIS" Or Left(Trim(avpItems(1)), 3) = "TKT" Or Left(Trim(avpItems(1)), 3) = "TFH" Or Left(Trim(avpItems(1)), 3) = "VCH" Then
            'by Abhi on 24-Apr-2011 for caseid 2126 Sabre Rail Segment 2nd stage
            If Left(Trim(avpItems(1)), 3) = "TKT" Then
            
            Else
                avpItems(1) = Left(avpItems(1), 3)
            End If
            'by Abhi on 24-Jun-2017 for caseid 7213 Sabre Conjunction Ticket loading issue
            'If vlTCount = 0 Then vlTCount = 1
            'by Abhi on 24-Jun-2017 for caseid 7213 Sabre Conjunction Ticket loading issue
        End If
        'by Abhi on 24-Jun-2017 for caseid 7213 Sabre Conjunction Ticket loading issue
        'For vlTI = 1 To vlTCount
        'by Abhi on 24-Jun-2017 for caseid 7213 Sabre Conjunction Ticket loading issue
            'by Abhi on 27-Aug-2010 for caseid 708 BLK should affect on above tickets for PER PP
            'Query = "Insert into SabreM5(UPLOADNO,IU5PTY,IU5MIN,IU5VR1,IU5VR2,IU5VR3,IU5VR4,IU5VR5,IU5VR6,IU5VR7,IU5VR8,IU5VR3_2) " & _
                    " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 2)) & "','" & Trim(fnMidValue(ln, 7, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 8, 1)) & "','" & Trim(fnMidValue(ln, 9, 2)) & "','" & Trim(avpItems(1)) & "', " & _
                    " '" & Trim(avpItems(2)) & "','" & Trim(avpItems(3)) & "','" & Trim(avpItems(4)) & "', " & _
                    " '" & Trim(avpItems(5)) & "/" & Trim(avpItems(6)) & "/" & Trim(avpItems(7)) & "/" & Trim(avpItems(8)) & "','" & Trim(fnMidValue(ln, 11, 4)) & "')"
            'by Abhi on 13-Nov-2010 for caseid 1549 Sabre M5 Vendor Code BMI etc
            'Query = "Insert into SabreM5(UPLOADNO,IU5PTY,IU5MIN,IU5VR1,IU5VR2,IU5VR3,IU5VR4,IU5VR5,IU5VR6,IU5VR7,IU5VR8,IU5VR3_2) " & _
                    " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 2)) & "','" & Trim(fnMidValue(ln, 7, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 8, 1)) & "','" & Trim(fnMidValue(ln, 9, 2)) & "','" & Trim(avpItems(1)) & "', " & _
                    " '" & Trim(avpItems(2)) & "','" & Trim(avpItems(3)) & "','" & Trim(avpItems(4)) & "', " & _
                    " '" & Trim(avpItems(5)) & "/" & Trim(avpItems(6)) & "/" & Trim(avpItems(7)) & "/" & Trim(avpItems(8)) & "/" & fNOBLKPP_String & "','" & Trim(fnMidValue(ln, 11, 4)) & "')"
            'by Abhi on 08-Mar-2014 for caseid 3469 Sabre EMD Add one more field for EMD indicator as a sub field of IU5VR1
            'Query = "Insert into SabreM5(UPLOADNO,IU5PTY,IU5MIN,IU5VR1,IU5VR2,IU5VR3,IU5VR4,IU5VR5,IU5VR6,IU5VR7,IU5VR8,IU5VR3_2) " & _
                    " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 2)) & "','" & Trim(fnMidValue(ln, 7, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 8, 1)) & "','" & Trim(fnMidValue(ln, 9, 2)) & "','" & Trim(avpItems(1)) & "', " & _
                    " '" & Trim(avpItems(2)) & "','" & Trim(avpItems(3)) & "','" & Trim(avpItems(4)) & "', " & _
                    " '" & Trim(avpItems(5)) & "/" & Trim(avpItems(6)) & "/" & Trim(avpItems(7)) & "/" & Trim(avpItems(8)) & "/" & fNOBLKPP_String & "','" & Trim(fnMidValue(ln, 9, 6)) & "')"
            'by Abhi on 24-Jun-2017 for caseid 7213 Sabre Conjunction Ticket loading issue
            If vlTCount > 1 Then
                avpItems(1) = avpItems(1) & "-" & Right(Val(avpItems(1)) + (Val(avpItems(7)) - 1), 2)
            End If
            'by Abhi on 24-Jun-2017 for caseid 7213 Sabre Conjunction Ticket loading issue
            Query = "Insert into SabreM5(UPLOADNO,IU5PTY,IU5MIN,IU5VR1,IU5VR2,IU5VR3,IU5VR4,IU5VR5,IU5VR6,IU5VR7,IU5VR8,IU5VR3_2,IU5VR1N) " & _
                    " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 2)) & "','" & Trim(fnMidValue(ln, 7, 1)) & "', " & _
                    " '" & Trim(fnMidValue(ln, 8, 1)) & "','" & Trim(fnMidValue(ln, 9, 2)) & "','" & Trim(avpItems(1)) & "', " & _
                    " '" & Trim(avpItems(2)) & "','" & Trim(avpItems(3)) & "','" & Trim(avpItems(4)) & "', " & _
                    " '" & Trim(avpItems(5)) & "/" & Trim(avpItems(6)) & "/" & Trim(avpItems(7)) & "/" & Trim(avpItems(8)) & "/" & Trim(avpItems(9)) & "/" & fNOBLKPP_String & "','" & Trim(fnMidValue(ln, 9, 6)) & "','" & SkipChars(avpItems(9)) & "')"
            
            'by Abhi on 08-Mar-2014 for caseid 3469 Sabre EMD Add one more field for EMD indicator as a sub field of IU5VR1
            'by Abhi on 25-Apr-2014 for caseid 3901 Sabre -Skip Refund Tickets
            'dbCompany.Execute Query
            If Trim(fnMidValue(ln, 8, 1)) <> "R" Then
                dbCompany.Execute Query
            End If
            'by Abhi on 25-Apr-2014 for caseid 3901 Sabre -Skip Refund Tickets
            'by Abhi on 27-Aug-2010 for caseid 708 BLK should affect on above tickets for PER PP
            If avpItems(1) = "BLK" And Trim(Mid(avpItems(0), 5, 2)) = "PP" Then
                fNOBLKPP_String = "NOBLKPP"
            End If
            avpItems(1) = Val(avpItems(1)) + 1
            avpItems(2) = 0
            avpItems(3) = 0
            avpItems(4) = 0
            avpItems(7) = 0
            DoEvents
        'by Abhi on 24-Jun-2017 for caseid 7213 Sabre Conjunction Ticket loading issue
        'Next vlTI
        'by Abhi on 24-Jun-2017 for caseid 7213 Sabre Conjunction Ticket loading issue
        Exit Sub
    ElseIf id = "M6" And Mid(ln, 3, 1) <> "/" Then
        Preid_String = id
        Query = "Insert into SabreM6(UPLOADNO,IU6RTY,IU6FCT,IU6CPR,IU6FC12) " & _
                " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 3)) & "','" & Trim(fnMidValue(ln, 6, 1)) & "', " & _
                " '" & Trim(fnMidValue(ln, 7, 1)) & "','" & Trim(fnMidValue(ln, 8, Len(ln))) & "')"
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M7" And Mid(ln, 3, 1) <> "/" Then
        Preid_String = id
        Query = "Insert into SabreM7(UPLOADNO,IU7RKN,IU7RMK) " & _
                " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, Len(ln))) & "')"
        dbCompany.Execute Query
        Exit Sub
    ElseIf id = "M9" And Mid(ln, 3, 1) <> "/" Then
        Preid_String = id
        'by Abhi on 03-Mar-2010 for caseid 1236 PENPART IN penline for Sabre
        If InStr(UCase(ln), "PENPART") > 0 Or InStr(UCase(ln), "PENPARTEND") > 0 Then
            'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
            If InStr(UCase(ln), PENLINEID_String & "PENPART") > 0 Or InStr(UCase(ln), PENLINEID_String & "PENPARTEND") > 0 Then
                If InStr(UCase(ln), PENLINEID_String & "PENPART") > 0 Then
                    ln = Left(ln, 4) & "PENPART" & Mid(ln, InStr(1, ln, "PENPART", vbTextCompare) + Len("PENPART"))
                ElseIf InStr(UCase(ln), PENLINEID_String & "PENPARTEND") > 0 Then
                    ln = Left(ln, 4) & "PENPARTEND" & Mid(ln, InStr(1, ln, "PENPARTEND", vbTextCompare) + Len("PENPARTEND"))
                End If
                aPEN = SplitWithLengths(ln, 2, 2, 15)
                ReDim Preserve aPEN(3)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN) " & _
                        " values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 15-Feb-2010 for caseid 1205 new penline for sabre
        ElseIf InStr(UCase(ln), "PENFARE") > 0 Then
            'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
            If InStr(UCase(ln), PENLINEID_String & "PENFARE") > 0 Then
                ln = Left(ln, 4) & "PENFARE/" & Mid(ln, InStr(1, ln, "PENFARE/", vbTextCompare) + Len("PENFARE/"))
                aPENFARE = Split(ln, "/")
                'by Abhi on 19-Mar-2015 for caseid 2836 Penline PENFARE Sell amt without N
                'ReDim Preserve aPENFARE(4)
                'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
                'ReDim Preserve aPENFARE(5)
                ReDim Preserve aPENFARE(6)
                'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
                'by Abhi on 06-Apr-2016 for caseid 6217 penfare line for sabre with supplier code
                vPENFARESELL_Currency = 0
                'by Abhi on 06-Apr-2016 for caseid 6217 penfare line for sabre with supplier code
                'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking
                'If InStr(1, aPENFARE(4), "N", vbTextCompare) = 0 Then
                vPENFAREPASSTYPE_String = ""
                'by Abhi on 15-Oct-2018 for caseid 9304 Passenger Type picking issue -Sabre file
                'If InStr(1, aPENFARE(4), "N", vbTextCompare) = 0 And InStr(1, aPENFARE(4), ".", vbTextCompare) = 0 Then
                faPENFARE4_String = aPENFARE(4)
                faPENFARE4_String = Replace(faPENFARE4_String, "N", "", , , vbTextCompare)
                If (InStr(1, aPENFARE(4), "N", vbTextCompare) = 0 And InStr(1, aPENFARE(4), ".", vbTextCompare) = 0) Or Val(faPENFARE4_String) = 0 Then
                'by Abhi on 15-Oct-2018 for caseid 9304 Passenger Type picking issue -Sabre file
                'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking
                    'by Abhi on 06-Apr-2016 for caseid 6217 penfare line for sabre with supplier code
                    'aPENFARE(5) = aPENFARE(3)
                    vPENFARESELL_Currency = Val(aPENFARE(3))
                    'by Abhi on 06-Apr-2016 for caseid 6217 penfare line for sabre with supplier code
                    aPENFARE(3) = 0
                    'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking
                    If IsNumeric(aPENFARE(4)) = False Then
                        vPENFAREPASSTYPE_String = aPENFARE(4)
                        aPENFARE(4) = ""
                    End If
                    'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking
                End If
                'by Abhi on 19-Mar-2015 for caseid 2836 Penline PENFARE Sell amt without N
                'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
                If Trim(aPENFARE(6)) = "" Then
                    aPENFARE(6) = "TKT"
                End If
                'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
                aPEN = SplitWithLengths(aPENFARE(0), 2, 2, 7)
                ReDim Preserve aPEN(3)
                'by Abhi on 19-Nov-2010 for caseid 1205 Penline PENFARE for Sabre passenger number picking changed to N1.1
                'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP) " & _
                        " values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "','" _
                        & "','','','" _
                        & "','','',''," & Val(aPEN(1)) & "," & Val(aPENFARE(1)) & "," & Val(aPENFARE(2)) & "," & Val(aPENFARE(3)) & ")"
                'by Abhi on 19-Mar-2015 for caseid 2836 Penline PENFARE Sell amt without N
                'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP) " & _
                        " values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "','" _
                        & "','','','" _
                        & "','','',''," & Val(Replace(aPENFARE(4), "N", "")) & "," & Val(aPENFARE(1)) & "," & Val(aPENFARE(2)) & "," & Val(aPENFARE(3)) & ")"
                'by Abhi on 06-Apr-2016 for caseid 6217 penfare line for sabre with supplier code
                'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,PENFARESELL) " & _
                        " values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "','" _
                        & "','','','" _
                        & "','','',''," & Val(Replace(aPENFARE(4), "N", "")) & "," & Val(aPENFARE(1)) & "," & Val(aPENFARE(2)) & "," & Val(aPENFARE(3)) & "," & Val(aPENFARE(5)) & ")"
                'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking
                'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,PENFARESELL,PENFARESUPPID) " & _
                        " values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "','" _
                        & "','','','" _
                        & "','','',''," & Val(Replace(aPENFARE(4), "N", "")) & "," & Val(aPENFARE(1)) & "," & Val(aPENFARE(2)) & "," & Val(aPENFARE(3)) & "," & vPENFARESELL_Currency & ",'" & SkipChars(aPENFARE(5)) & "')"
                'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
                'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,PENFARESELL,PENFARESUPPID,PENFAREPASSTYPE) " & _
                        " values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "','" _
                        & "','','','" _
                        & "','','',''," & Val(Replace(aPENFARE(4), "N", "")) & "," & Val(aPENFARE(1)) & "," & Val(aPENFARE(2)) & "," & Val(aPENFARE(3)) & "," & vPENFARESELL_Currency & ",'" _
                        & SkipChars(aPENFARE(5)) & "','" & SkipChars(vPENFAREPASSTYPE_String) & "')"
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,PENFARESELL,PENFARESUPPID,PENFAREPASSTYPE,PENFARETICKETTYPE) " & _
                        " values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "','" _
                        & "','','','" _
                        & "','','',''," & Val(Replace(aPENFARE(4), "N", "")) & "," & Val(aPENFARE(1)) & "," & Val(aPENFARE(2)) & "," & Val(aPENFARE(3)) & "," & vPENFARESELL_Currency & ",'" _
                        & SkipChars(aPENFARE(5)) & "','" & SkipChars(vPENFAREPASSTYPE_String) & "','" & SkipChars(aPENFARE(6)) & "')"
                'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
                'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking
                'by Abhi on 06-Apr-2016 for caseid 6217 penfare line for sabre with supplier code
                'by Abhi on 19-Mar-2015 for caseid 2836 Penline PENFARE Sell amt without N
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 23-Feb-2010 for caseid 1222 Accounting line for sabre in penline
        ElseIf InStr(UCase(ln), "PENOTH") > 0 Then
            'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
            If InStr(UCase(ln), PENLINEID_String & "PENOTH") > 0 Then
                ln = Left(ln, 4) & "PENOTH/" & Mid(ln, InStr(1, ln, "PENOTH/", vbTextCompare) + Len("PENOTH/"))
                aPEN = Split(ln, "/")
                ReDim Preserve aPEN(10)
                'by Abhi on 25-Nov-2010 for caseid 1549 Sabre M9 PENOTH product code also picking 6 characters
                'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'" & Mid(aPEN(1), 1, 2) & "','" & aPEN(2) & "','" & aPEN(3) & "','" & aPEN(4) & "','" & aPEN(5) & "','" & aPEN(7) & "/" & aPEN(8) & "/" & aPEN(9) & "/" & aPEN(10) & "','" & Mid(aPEN(1), 3, 4) & "')"
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'" & Mid(aPEN(1), 1, 2) & "','" & aPEN(2) & "','" & aPEN(3) & "','" & aPEN(4) & "','" & aPEN(5) & "','" & aPEN(7) & "/" & aPEN(8) & "/" & aPEN(9) & "/" & aPEN(10) & "','" & aPEN(1) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 27-Aug-2010 for caseid 1473 Penline PENO
'        ElseIf InStr(UCase(ln), "PENO") > 0 Then
'            'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
'            If InStr(UCase(ln), PENLINEID_String & "PENO") > 0 Then
'                ln = Left(ln, 4) & "PENO/" & Mid(ln, InStr(1, ln, "PENO/", vbTextCompare) + Len("PENO/"))
'                aPENO = Split(ln, "/")
'                ReDim Preserve aPENO(5)
'                aPEN = SplitWithLengths(aPENFARE(0), 2, 2, 7)
'                ReDim Preserve aPEN(3)
'                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP) " & _
'                        " values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "','" _
'                        & "','','','" _
'                        & "','','',''," & Val(aPEN(1)) & "," & Val(aPENFARE(1)) & "," & Val(aPENFARE(2)) & "," & Val(aPENFARE(3)) & ")"
'                dbCompany.Execute Query
'            End If
'            Exit Sub
        'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking
        ElseIf InStr(UCase(ln), "PENO") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENO") > 0 Then
                ln = Left(ln, 4) & "PENO/" & Mid(ln, InStr(1, ln, "PENO/", vbTextCompare) + Len("PENO/"))
                aPENO = Split(ln, "/")
                ReDim Preserve aPENO(6)
                aPEN = SplitWithLengths(aPENO(0), 2, 2, 7)
                ReDim Preserve aPEN(3)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,PENOPRDID,PENOQTY,PENORATE,PENOSELL,PENOSUPPID,PENOPAYMETHODID) " & _
                        "values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "', " & _
                        "'" & SkipChars(aPENO(1)) & "'," & Val(aPENO(2)) & "," & Val(aPENO(3)) & "," & Val(aPENO(4)) & ",'" & SkipChars(aPENO(5)) & "','" & SkipChars(aPENO(6)) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        ElseIf InStr(UCase(ln), "PENAGROSS") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENAGROSS") > 0 Then
                ln = Left(ln, 4) & "PENAGROSS/" & Mid(ln, InStr(1, ln, "PENAGROSS/", vbTextCompare) + Len("PENAGROSS/"))
                aPENAGROSS = Split(ln, "/")
                ReDim Preserve aPENAGROSS(5)
                aPEN = SplitWithLengths(aPENAGROSS(0), 2, 2, Len("PENAGROSS"))
                ReDim Preserve aPEN(3)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,PENAGROSSADULT,PENAGROSSCHILD,PENAGROSSINFANT,PENAGROSSPACKAGE,PENAGROSSYOUTH) " & _
                        "values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "', " & _
                        "" & Val(aPENAGROSS(1)) & "," & Val(aPENAGROSS(2)) & "," & Val(aPENAGROSS(3)) & "," & Val(aPENAGROSS(4)) & "," & Val(aPENAGROSS(5)) & ")"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 11-Apr-2016 for caseid 6240 Sabre Passenger Type picking
        
        'by Abhi on 31-Aug-2010 for caseid 1433 Penline PENAUTOOFF
        ElseIf InStr(UCase(ln), "PENAUTOOFF") > 0 Then
            'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
            If InStr(UCase(ln), PENLINEID_String & "PENAUTOOFF") > 0 Then
                ln = Left(ln, 4) & "PENAUTOOFF" & Mid(ln, InStr(1, ln, "PENAUTOOFF", vbTextCompare) + Len("PENAUTOOFF"))
                aPEN = SplitWithLengths(ln, 2, 2, 10)
                ReDim Preserve aPEN(3)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP) " & _
                        " values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0)"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
        ElseIf InStr(UCase(ln), "PENATOL") > 0 Then
            'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
            If InStr(UCase(ln), PENLINEID_String & "PENATOL") > 0 Then
                ln = Left(ln, 4) & "PENATOL" & Mid(ln, InStr(1, ln, "PENATOL", vbTextCompare) + Len("PENATOL"))
                aPEN = Split(ln, "/")
                ReDim Preserve aPEN(2)
                'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','','" & Left(aPEN(1), 50) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENRT") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENRT") > 0 Then
                ln = Left(ln, 4) & "PENRT" & Mid(ln, InStr(1, ln, "PENRT", vbTextCompare) + Len("PENRT"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'" & Left(aPEN(1), 50) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENPOL") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENPOL") > 0 Then
                ln = Left(ln, 4) & "PENPOL" & Mid(ln, InStr(1, ln, "PENPOL", vbTextCompare) + Len("PENPOL"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','" & Left(aPEN(1), 50) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENPROJ") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENPROJ") > 0 Then
                ln = Left(ln, 4) & "PENPROJ" & Mid(ln, InStr(1, ln, "PENPROJ", vbTextCompare) + Len("PENPROJ"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','" & Left(aPEN(1), 50) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENEID") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENEID") > 0 Then
                ln = Left(ln, 4) & "PENEID" & Mid(ln, InStr(1, ln, "PENEID", vbTextCompare) + Len("PENEID"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','" & Left(aPEN(1), 50) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENHFRC") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENHFRC") > 0 Then
                ln = Left(ln, 4) & "PENHFRC" & Mid(ln, InStr(1, ln, "PENHFRC", vbTextCompare) + Len("PENHFRC"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','','" & Left(aPEN(1), 50) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENLFRC") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENLFRC") > 0 Then
                ln = Left(ln, 4) & "PENLFRC" & Mid(ln, InStr(1, ln, "PENLFRC", vbTextCompare) + Len("PENLFRC"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC,PENLFRC) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','','','" & Left(aPEN(1), 50) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENHIGHF") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENHIGHF") > 0 Then
                ln = Left(ln, 4) & "PENHIGHF" & Mid(ln, InStr(1, ln, "PENHIGHF", vbTextCompare) + Len("PENHIGHF"))
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                ln = Replace(ln, "/", "-", , , vbTextCompare)
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                aPEN = Split(ln, "-")
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                'ReDim Preserve aPEN(2)
                ReDim Preserve aPEN(3)
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC,PENLFRC,PENHIGHF,PENHIGHFTKTNO) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','','',''," & Val(aPEN(1)) & ",'" & Left(SkipChars(aPEN(3)), 50) & "')"
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENLOWF") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENLOWF") > 0 Then
                ln = Left(ln, 4) & "PENLOWF" & Mid(ln, InStr(1, ln, "PENLOWF", vbTextCompare) + Len("PENLOWF"))
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                ln = Replace(ln, "/", "-", , , vbTextCompare)
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                aPEN = Split(ln, "-")
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                'ReDim Preserve aPEN(2)
                ReDim Preserve aPEN(3)
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC,PENLFRC,PENHIGHF,PENLOWF,PENLOWFTKTNO) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','','','',0," & Val(aPEN(1)) & ",'" & Left(SkipChars(aPEN(3)), 50) & "')"
                'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENUC1") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENUC1") > 0 Then
                ln = Left(ln, 4) & "PENUC1" & Mid(ln, InStr(1, ln, "PENUC1", vbTextCompare) + Len("PENUC1"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC,PENLFRC,PENHIGHF,PENLOWF,PENUC1) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','','','',0,0,'" & Left(aPEN(1), 50) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENUC2") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENUC2") > 0 Then
                ln = Left(ln, 4) & "PENUC2" & Mid(ln, InStr(1, ln, "PENUC2", vbTextCompare) + Len("PENUC2"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC,PENLFRC,PENHIGHF,PENLOWF,PENUC1,PENUC2) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','','','',0,0,'','" & Left(aPEN(1), 50) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENUC3") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENUC3") > 0 Then
                ln = Left(ln, 4) & "PENUC3" & Mid(ln, InStr(1, ln, "PENUC3", vbTextCompare) + Len("PENUC3"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC,PENLFRC,PENHIGHF,PENLOWF,PENUC1,PENUC2,PENUC3) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','','','',0,0,'','','" & Left(aPEN(1), 50) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENCC") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENCC") > 0 Then
                ln = Left(ln, 4) & "PENCC" & Mid(ln, InStr(1, ln, "PENCC", vbTextCompare) + Len("PENCC"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC,PENLFRC,PENHIGHF,PENLOWF,PENUC1,PENUC2,PENUC3) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "'," _
                        & "'','','','" & Left(aPEN(1), 255) & "',''" _
                        & ",'','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','','','',0,0,'','','')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENPO") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENPO") > 0 Then
                ln = Left(ln, 4) & "PENPO" & Mid(ln, InStr(1, ln, "PENPO", vbTextCompare) + Len("PENPO"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC,PENLFRC,PENHIGHF,PENLOWF,PENUC1,PENUC2,PENUC3) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "'," _
                        & "'','','" & Left(aPEN(1), 255) & "','',''" _
                        & ",'','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','','','',0,0,'','','')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 28-Apr-2011 for caseid 2123 Penline Corporate booking for airplus and barclays
        ElseIf InStr(UCase(ln), "PENBB") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENBB") > 0 Then
                ln = Left(ln, 4) & "PENBB" & Mid(ln, InStr(1, ln, "PENBB", vbTextCompare) + Len("PENBB"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC,PENLFRC,PENHIGHF,PENLOWF,PENUC1,PENUC2,PENUC3) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "'," _
                        & "'','','','',''" _
                        & ",'','',0,0,0,0,'','','','','','','','','" & Left(aPEN(1), 50) & "',''" _
                        & ",'','','','','','',0,0,'','','')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 01-Oct-2012 for caseid 2511 Penline for PENAGENTCOM
        ElseIf InStr(UCase(ln), "PENAGENTCOM") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENAGENTCOM") > 0 Then
                ln = Left(ln, 4) & "PENAGENTCOM/" & Mid(ln, InStr(1, ln, "PENAGENTCOM/", vbTextCompare) + Len("PENAGENTCOM/"))
                aPEN = Split(ln, "/")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC,PENLFRC,PENHIGHF,PENLOWF,PENUC1,PENUC2,PENUC3,PENAGENTCOMSUM,PENAGENTCOMVAT) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "'," _
                        & "'','','','',''" _
                        & ",'','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','','','',0,0,'','',''," & Val(aPEN(1)) & "," & Val(aPEN(2)) & ")"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
        ElseIf InStr(UCase(ln), "PENAIRTKT") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENAIRTKT") > 0 Then
                'ln = Left(ln, 4) & "PENAIRTKT/" & Mid(ln, InStr(1, ln, "PENAIRTKT/", vbTextCompare) + Len("PENAIRTKT/"))
                Set CollPenSplited = RecM9_PENAIRTKT(ln)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AIRTKTPAX,AIRTKTTKT,AIRTKTDATE) " & _
                        " values(" & UploadNo & ",'','','" & CollPenSplited("ITNO") & "','PENAIRTKT'" _
                        & ",'" & CollPenSplited("PAX") & "','" & CollPenSplited("TKT") & "','" & CollPenSplited("DATE") & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
        'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
        ElseIf InStr(UCase(ln), "PENBILLCUR") > 0 Then 'M905P7PENBILLCUR-USD
            If InStr(UCase(ln), PENLINEID_String & "PENBILLCUR") > 0 Then
                ln = Left(ln, 4) & "PENBILLCUR" & Mid(ln, InStr(1, ln, "PENBILLCUR", vbTextCompare) + Len("PENBILLCUR"))
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(2)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,PENBILLCUR)" & _
                        " values(" & UploadNo & ",'','','" & SkipChars(Mid(aPEN(0), 3, 2)) & "','" & SkipChars(Mid(aPEN(0), 5)) & "','" & SkipChars(Left(aPEN(1), 4)) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        ElseIf InStr(UCase(ln), "PENWAIT") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENWAIT") > 0 Then
                ln = Left(ln, 4) & "PENWAIT" & Mid(ln, InStr(1, ln, "PENWAIT", vbTextCompare) + Len("PENWAIT"))
                aPEN = SplitWithLengths(ln, 2, 2, 7)
                ReDim Preserve aPEN(3)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP) " & _
                        " values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0)"
                dbCompany.Execute Query
                GIT_PENWAIT_String = "Y"
            End If
            Exit Sub
        'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
        'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
        ElseIf InStr(UCase(ln), "PENVC") > 0 Then
            'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
            If InStr(UCase(ln), PENLINEID_String & "PENVC") > 0 Then
                ln = Left(ln, 4) & "PENVC" & Mid(ln, InStr(1, ln, "PENVC", vbTextCompare) + Len("PENVC"))
                aPEN = SplitWithLengths(ln, 2, 2, 5)
                ReDim Preserve aPEN(3)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP) " & _
                        " values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0)"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
        'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
        ElseIf InStr(UCase(ln), "PENCS") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENCS") > 0 Then
                ln = Left(ln, 4) & "PENCS/" & Mid(ln, InStr(1, ln, "PENCS/", vbTextCompare) + Len("PENCS/"))
                aPENCS = Split(ln, "/")
                ReDim Preserve aPENCS(3)
                aPEN = SplitWithLengths(aPENCS(0), 2, 2, Len("PENCS"))
                ReDim Preserve aPEN(3)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,PENCSLABELID,PENCSLISTID) " & _
                        "values(" & UploadNo & ",'','','" & aPEN(1) & "','" & aPEN(2) & "', " & _
                        "N'" & SkipChars(Left(Trim(aPENCS(1)), 20)) & "', N'" & SkipChars(Left(Trim(aPENCS(2)), 20)) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
        'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
        ElseIf InStr(UCase(ln), "PENRC") > 0 Then
            If InStr(UCase(ln), PENLINEID_String & "PENRC") > 0 Then
                ln = Left(ln, 4) & "PENRC" & Mid(ln, InStr(1, ln, "PENRC", vbTextCompare) + Len("PENRC"))
                ln = Replace(ln, "/", "-", , , vbTextCompare)
                aPEN = Split(ln, "-")
                ReDim Preserve aPEN(3)
                Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,FARE,TAXES,MARKUP,M9VR3,M9VR4,M9VR5,M9VR6,M9VR7,M9VR8,M9VR3_2,MC,BB,PENATOLTYPE" & _
                        ",PENRT,PENPOL,PENPROJ,PENEID,PENHFRC,PENLFRC,PENHIGHF,PENRC,PENRCTKTNO) " & _
                        " values(" & UploadNo & ",'','','" & Mid(aPEN(0), 3, 2) & "','" & Mid(aPEN(0), 5) & "','" _
                        & "','','','" _
                        & "','','','',0,0,0,0,'','','','','','','','','',''" _
                        & ",'','','','','','',0,'" & Left(SkipChars(aPEN(1)), 50) & "','" & Left(SkipChars(aPEN(3)), 50) & "')"
                dbCompany.Execute Query
            End If
            Exit Sub
        'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
        ElseIf InStr(UCase(ln), "PEN") > 0 Then
            'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
            If InStr(UCase(ln), PENLINEID_String & "PEN") > 0 Then
                For TempPENPNO = 1 To 99
                    'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
                    'If InStr(1, ln, "P" & TempPENPNO & "E-", vbTextCompare) > 0 Or InStr(1, ln, "P" & TempPENPNO & "T-", vbTextCompare) > 0 Then
                    'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
                    'If InStr(1, ln, PENLINEID_String & "E" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, ln, PENLINEID_String & "T" & TempPENPNO & "-", vbTextCompare) > 0 Then
                    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
                    'If InStr(1, ln, "E" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, ln, "T" & TempPENPNO & "-", vbTextCompare) > 0 Then
                    If InStr(1, ln, "E" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, ln, "T" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, ln, "DOB" & TempPENPNO & "-", vbTextCompare) > 0 Then
                    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
                        Set CollPenSplited = RecM9_PENLINEPassenger(ln, TempPENPNO)
                        'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
                        'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO) " & _
                                " values(" & UploadNo & ",'','','" & CollPenSplited("ITNO") & "','PEN','" _
                                & "','','','" _
                                & "','','" & CollPenSplited("PEMAIL") & "','" & CollPenSplited("PTELE") & "'," & CollPenSplited("PNO") & ")"
                        Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,PDOB) " & _
                                " values(" & UploadNo & ",'','','" & CollPenSplited("ITNO") & "','PEN','" _
                                & "','','','" _
                                & "','','" & CollPenSplited("PEMAIL") & "','" & CollPenSplited("PTELE") & "'," & CollPenSplited("PNO") & ",'" & CollPenSplited("PDOB") & "')"
                        'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
                        dbCompany.Execute Query
                    End If
                    DoEvents
                Next
                'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
                ln = Left(ln, 4) & "PEN\" & Mid(ln, InStr(1, ln, "PEN\", vbTextCompare) + Len("PEN\"))
                'by Abhi on 13-Dec-2011 for caseid 1591 Sabre customer field picking the phone no if the telephone tag is wrong
                'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                'If InStr(1, ln, "AC-") > 0 Or InStr(1, ln, "DE-") > 0 Or InStr(1, ln, "PO-") > 0 Or InStr(1, ln, "CC-") > 0 Or InStr(1, ln, "BR-") > 0 Or InStr(1, ln, "MC-") > 0 Or InStr(1, ln, "BB-") > 0 Then
                'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                'If InStr(1, ln, "AC-") > 0 Or InStr(1, ln, "DE-") > 0 Or InStr(1, ln, "PO-") > 0 Or InStr(1, ln, "CC-") > 0 Or InStr(1, ln, "BR-") > 0 Or InStr(1, ln, "MC-") > 0 Or InStr(1, ln, "BB-") > 0 Or InStr(1, ln, "INETREF-") > 0 Then
                'by Abhi on 21-Nov-2014 for caseid 4746 Delivery address from penline is not picking for Sabre File
                'If InStr(1, ln, "AC-") > 0 Or InStr(1, ln, "DE-") > 0 Or InStr(1, ln, "PO-") > 0 Or InStr(1, ln, "CC-") > 0 Or InStr(1, ln, "BR-") > 0 Or InStr(1, ln, "MC-") > 0 Or InStr(1, ln, "BB-") > 0 Or InStr(1, ln, "INETREF-") > 0 Or InStr(1, ln, "TKTDEADLINE-") > 0 Then
                'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                'If InStr(1, ln, "AC-") > 0 Or InStr(1, ln, "DE-") > 0 Or InStr(1, ln, "PO-") > 0 Or InStr(1, ln, "CC-") > 0 Or InStr(1, ln, "BR-") > 0 Or InStr(1, ln, "MC-") > 0 Or InStr(1, ln, "BB-") > 0 Or InStr(1, ln, "INETREF-") > 0 Or InStr(1, ln, "TKTDEADLINE-") > 0 Or InStr(1, ln, "ADD-") > 0 Then
                'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                'If InStr(1, ln, "AC-") > 0 Or InStr(1, ln, "DE-") > 0 Or InStr(1, ln, "PO-") > 0 Or InStr(1, ln, "CC-") > 0 Or InStr(1, ln, "BR-") > 0 Or InStr(1, ln, "MC-") > 0 Or InStr(1, ln, "BB-") > 0 Or InStr(1, ln, "INETREF-") > 0 Or InStr(1, ln, "TKTDEADLINE-") > 0 Or InStr(1, ln, "ADD-") > 0 Or InStr(1, ln, "DEPT-") > 0 Then
                If InStr(1, ln, "AC-") > 0 Or InStr(1, ln, "DE-") > 0 Or InStr(1, ln, "PO-") > 0 Or InStr(1, ln, "CC-") > 0 _
                    Or InStr(1, ln, "BR-") > 0 Or InStr(1, ln, "MC-") > 0 Or InStr(1, ln, "BB-") > 0 Or InStr(1, ln, "INETREF-") > 0 _
                    Or InStr(1, ln, "TKTDEADLINE-") > 0 Or InStr(1, ln, "ADD-") > 0 Or InStr(1, ln, "DEPT-") > 0 _
                    Or InStr(1, ln, "DEPOSITAMT-") > 0 Or InStr(1, ln, "DEPOSITDUEDATE-") > 0 Then
                'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                'by Abhi on 21-Nov-2014 for caseid 4746 Delivery address from penline is not picking for Sabre File
                'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    avpItems = Split(ln, "/")
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'ReDim Preserve avpItems(7)
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'ReDim Preserve avpItems(8)
                    'by Abhi on 21-Nov-2014 for caseid 4746 Delivery address from penline is not picking for Sabre File
                    'aiTo_Long = 9
                    'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                    'aiTo_Long = 10
                    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                    'aiTo_Long = 11
                    aiTo_Long = 13
                    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                    'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                    'by Abhi on 21-Nov-2014 for caseid 4746 Delivery address from penline is not picking for Sabre File
                    ReDim Preserve avpItems(aiTo_Long)
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
                    avpItems(0) = Mid(avpItems(0), 9)
                    
                    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
                    'For ai = 1 To 5
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'For ai = 0 To 5
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'For ai = 0 To 8
                    For ai = 0 To aiTo_Long
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                        If UCase(Left(avpItems(ai), 2)) = "AC" And ai <> 1 Then
                            avpItemsTemp = avpItems(1)
                            avpItems(1) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
                    'For ai = 1 To 5
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'For ai = 0 To 5
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'For ai = 0 To 8
                    For ai = 0 To aiTo_Long
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                        If UCase(Left(avpItems(ai), 2)) = "DE" And ai <> 2 Then
                            avpItemsTemp = avpItems(2)
                            avpItems(2) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
                    'For ai = 1 To 5
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'For ai = 0 To 5
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'For ai = 0 To 8
                    For ai = 0 To aiTo_Long
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                        If UCase(Left(avpItems(ai), 2)) = "PO" And ai <> 3 Then
                            avpItemsTemp = avpItems(3)
                            avpItems(3) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
                    'For ai = 1 To 5
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'For ai = 0 To 5
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'For ai = 0 To 8
                    For ai = 0 To aiTo_Long
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                        If UCase(Left(avpItems(ai), 2)) = "CC" And ai <> 4 Then
                            avpItemsTemp = avpItems(4)
                            avpItems(4) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
                    'For ai = 1 To 5
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'For ai = 0 To 5
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'For ai = 0 To 8
                    For ai = 0 To aiTo_Long
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                        If UCase(Left(avpItems(ai), 2)) = "BR" And ai <> 5 Then
                            avpItemsTemp = avpItems(5)
                            avpItems(5) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    
                    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'For ai = 0 To 5
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'For ai = 0 To 8
                    For ai = 0 To aiTo_Long
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                        If UCase(Left(avpItems(ai), 2)) = "MC" And ai <> 6 Then
                            avpItemsTemp = avpItems(6)
                            avpItems(6) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'For ai = 0 To 5
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'For ai = 0 To 8
                    For ai = 0 To aiTo_Long
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                        If UCase(Left(avpItems(ai), 2)) = "BB" And ai <> 7 Then
                            avpItemsTemp = avpItems(7)
                            avpItems(7) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'For ai = 0 To 8
                    For ai = 0 To aiTo_Long
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                        If UCase(Left(avpItems(ai), Len("INETREF"))) = "INETREF" And ai <> 8 Then
                            avpItemsTemp = avpItems(8)
                            avpItems(8) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    For ai = 0 To aiTo_Long
                        If UCase(Left(avpItems(ai), Len("TKTDEADLINE"))) = "TKTDEADLINE" And ai <> 9 Then
                            avpItemsTemp = avpItems(9)
                            avpItems(9) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'by Abhi on 21-Nov-2014 for caseid 4746 Delivery address from penline is not picking for Sabre File
                    For ai = 0 To aiTo_Long
                        If UCase(Left(avpItems(ai), Len("ADD"))) = "ADD" And ai <> 10 Then
                            avpItemsTemp = avpItems(10)
                            avpItems(10) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    'by Abhi on 21-Nov-2014 for caseid 4746 Delivery address from penline is not picking for Sabre File
                    'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                    For ai = 0 To aiTo_Long
                        If UCase(Left(avpItems(ai), Len("DEPT"))) = "DEPT" And ai <> 11 Then
                            avpItemsTemp = avpItems(11)
                            avpItems(11) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                    For ai = 0 To aiTo_Long
                        If UCase(Left(avpItems(ai), Len("DEPOSITAMT"))) = "DEPOSITAMT" And ai <> 12 Then
                            avpItemsTemp = avpItems(12)
                            avpItems(12) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    
                    For ai = 0 To aiTo_Long
                        If UCase(Left(avpItems(ai), Len("DEPOSITDUEDATE"))) = "DEPOSITDUEDATE" And ai <> 13 Then
                            avpItemsTemp = avpItems(13)
                            avpItems(13) = avpItems(ai)
                            avpItems(ai) = avpItemsTemp
                        End If
                        DoEvents
                    Next ai
                    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                    
                    'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC) " & _
                    '        " values(" & UploadNo & ",'','','" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "','" _
                    '        & Trim(fnMidValue(ln, 12, 4)) & "','" & Trim(fnMidValue(ln, 20, 3)) & "','" & Trim(fnMidValue(ln, 27, 4)) & "','" _
                    '        & Trim(fnMidValue(ln, 35, 2)) & "')"
                    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
                    'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO) " & _
                            " values(" & UploadNo & ",'','','" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "','" _
                            & Trim(Mid(avpItems(1), 4)) & "','" & Trim(Mid(avpItems(2), 4)) & "','" & Trim(Mid(avpItems(3), 4)) & "','" _
                            & Trim(Mid(avpItems(4), 4)) & "','" & Trim(Mid(avpItems(5), 4)) & "','','',0)"
                    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,MC,BB) " & _
                            " values(" & UploadNo & ",'','','" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "','" _
                            & Trim(Mid(avpItems(1), 4)) & "','" & Trim(Mid(avpItems(2), 4)) & "','" & Trim(Mid(avpItems(3), 4)) & "','" _
                            & Trim(Mid(avpItems(4), 4)) & "','" & Trim(Mid(avpItems(5), 4)) & "','','',0,'" & Left(Trim(Mid(avpItems(6), 4)), 50) & "','" & Left(Trim(Mid(avpItems(7), 4)), 50) & "')"
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,MC,BB,INETREF) " & _
                            " values(" & UploadNo & ",'','','" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "','" _
                            & Trim(Mid(avpItems(1), 4)) & "','" & Trim(Mid(avpItems(2), 4)) & "','" & Trim(Mid(avpItems(3), 4)) & "','" _
                            & Trim(Mid(avpItems(4), 4)) & "','" & Trim(Mid(avpItems(5), 4)) & "','','',0,'" & Left(Trim(Mid(avpItems(6), 4)), 50) & "','" & Left(Trim(Mid(avpItems(7), 4)), 50) & "','" _
                            & Left(Trim(Mid(avpItems(8), 9)), 100) & "')"
                    'by Abhi on 21-Nov-2014 for caseid 4746 Delivery address from penline is not picking for Sabre File
                    'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                    'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,MC,BB,INETREF,TKTDEADLINE,DeliAdd) " & _
                            " values(" & UploadNo & ",'','','" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "','" _
                            & Trim(Mid(avpItems(1), 4)) & "','" & Trim(Mid(avpItems(2), 4)) & "','" & Trim(Mid(avpItems(3), 4)) & "','" _
                            & Trim(Mid(avpItems(4), 4)) & "','" & Trim(Mid(avpItems(5), 4)) & "','','',0,'" & Left(Trim(Mid(avpItems(6), 4)), 50) & "','" & Left(Trim(Mid(avpItems(7), 4)), 50) & "','" _
                            & Left(Trim(Mid(avpItems(8), 9)), 100) & "','" & Left(Trim(Mid(avpItems(9), 13)), 9) & "','" & Left(Trim(Mid(avpItems(10), 5)), 500) & "')"
                    'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
                    'Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,MC,BB,INETREF,TKTDEADLINE,DeliAdd,DEPT) " & _
                            " values(" & UploadNo & ",'','','" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "','" _
                            & Trim(Mid(avpItems(1), 4)) & "','" & Trim(Mid(avpItems(2), 4)) & "','" & Trim(Mid(avpItems(3), 4)) & "','" _
                            & Trim(Mid(avpItems(4), 4)) & "','" & Trim(Mid(avpItems(5), 4)) & "','','',0,'" & Left(Trim(Mid(avpItems(6), 4)), 50) & "','" & Left(Trim(Mid(avpItems(7), 4)), 50) & "','" _
                            & Left(Trim(Mid(avpItems(8), 9)), 100) & "','" & Left(Trim(Mid(avpItems(9), 13)), 9) & "','" & Left(Trim(Mid(avpItems(10), 5)), 500) & "','" & Left(Trim(Mid(avpItems(11), 6)), 50) & "')"
                    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                    Query = "Insert into SabreM9(UPLOADNO,Details1,Details2,ITNO,PEN,AC,FD,PO,CC,BR,PEMAIL,PTELE,PNO,MC,BB,INETREF,TKTDEADLINE,DeliAdd,DEPT,DEPOSITAMT,DEPOSITDUEDATE) " & _
                            " values(" & UploadNo & ",'','','" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, 3)) & "','" _
                            & Left(Trim(Mid(avpItems(1), 4)), 50) & "','" & Trim(Mid(avpItems(2), 4)) & "','" & Trim(Mid(avpItems(3), 4)) & "','" _
                            & Trim(Mid(avpItems(4), 4)) & "','" & Trim(Mid(avpItems(5), 4)) & "','','',0,'" & Left(Trim(Mid(avpItems(6), 4)), 50) & "','" & Left(Trim(Mid(avpItems(7), 4)), 50) & "','" _
                            & Left(Trim(Mid(avpItems(8), 9)), 100) & "','" & Left(Trim(Mid(avpItems(9), 13)), 9) & "','" & Left(Trim(Mid(avpItems(10), 5)), 500) & "','" & Left(Trim(Mid(avpItems(11), 6)), 50) & "'," _
                            & Val(Trim(Mid(avpItems(12), 12))) & ", '" & Left(Trim(Mid(avpItems(13), 16)), 9) & "')"
                    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
                    'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
                    'by Abhi on 26-Dec-2014 for caseid 4846 Business Area penline for Galileo and Sabre
                    'by Abhi on 21-Nov-2014 for caseid 4746 Delivery address from penline is not picking for Sabre File
                    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
                    'by Abhi on 13-Dec-2013 for caseid 3559 Penline for INETREF for Sabre
                    dbCompany.Execute Query
                    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
                    If Trim(Mid(avpItems(1), 4)) <> "" Then
                        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
                        'GIT_CUSTOMERUSERCODE_String = Mid(avpItems(1), 4)
                        GIT_CUSTOMERUSERCODE_String = Left(Mid(avpItems(1), 4), 50)
                        'by Abhi on 02-Aug-2017 for caseid 7675 Error: -2147217833 - String or binary data would be truncated due to PENLINE AC data more than 20 characters
                    End If
                    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
                End If
            End If
        Else
            Query = "Insert into SabreM9(UPLOADNO,Details1,Details2) " & _
                    " values(" & UploadNo & ",'" & Trim(fnMidValue(ln, 3, 2)) & "','" & Trim(fnMidValue(ln, 5, Len(ln))) & "')"
            dbCompany.Execute Query
        End If
        Exit Sub
    'by Abhi on 08-Sep-2014 for caseid 4496 Sabre Airline Fee should be taken from MA Record
    ElseIf id = "MA" And Mid(ln, 3, 1) <> "/" Then
        Preid_String = id
        aMA = SplitWithLengths(ln, 2, 2, 2, 1, 13, 3, 2, 3, 3, 1, 1, 1, 1, 2, 25, 1, 2, 11, 3)
        Query = "Insert into SabreMA(UPLOADNO, IUAOMID, IUAOSEQ, IUAOITN, IUAOCHA, IUAOTKN, IUAOVAL, IUAOTYP, IUAOSUB, IUAOSP3, IUAOSYS, IUAOREF, IUAOCMA, IUAOITL, IUAOSP4, IUAODES, IUAOSP5, IUAOXCN, IUAOAMT, IUAOCUR) " & _
                " values(" & UploadNo & ", '" & SkipChars(aMA(0)) & "', " & Val(aMA(1)) & ", '" & SkipChars(aMA(2)) & "', '" & SkipChars(aMA(3)) & "', '" & SkipChars(aMA(4)) & "', '" & SkipChars(aMA(5)) & "', '" & SkipChars(aMA(6)) & "', '" & SkipChars(aMA(7)) & "', '" & SkipChars(aMA(8)) & "', '" & SkipChars(aMA(9)) & "', '" & SkipChars(aMA(10)) & "', '" & SkipChars(aMA(11)) & "', '" & SkipChars(aMA(12)) & "', '" & SkipChars(aMA(13)) & "', '" & SkipChars(aMA(14)) & "', '" & SkipChars(aMA(15)) & "', '" & SkipChars(aMA(16)) & "', '" & SkipChars(aMA(17)) & "', '" & SkipChars(aMA(18)) & "')"
        dbCompany.Execute Query
    'by Abhi on 08-Sep-2014 for caseid 4496 Sabre Airline Fee should be taken from MA Record
    'by Abhi on 03-Nov-2014 for caseid 4487 ME tag String or binary data would be truncated skipped error
    ElseIf id = "ME" Then
        Preid_String = id
        Exit Sub
    'by Abhi on 03-Nov-2014 for caseid 4487 ME tag String or binary data would be truncated skipped error
'    ElseIf Left(ln, 4) = "PEN\" Then
        
'        Exit Sub
    'by Abhi on 10-Mar-2015 for caseid 5016 Sabre-Delivery address,phone number and email issue
    ElseIf id = "MF" And Mid(ln, 3, 1) <> "/" Then
        Preid_String = id
        aMF = SplitWithLengthsPlus(ln, 2, 2)
        ReDim Preserve aMF(2)
        aMF(2) = ExtractBetween(aMF(2), "#", "#")
        Query = "Insert into SabreMF(UPLOADNO, IUFMID, IUFITM, IUFEMA) " & _
                " values(" & UploadNo & ", '" & SkipChars(aMF(0)) & "', '" & SkipChars(aMF(1)) & "', '" & Left(SkipChars(aMF(2)), 63) & "')"
        dbCompany.Execute Query
    'by Abhi on 10-Mar-2015 for caseid 5016 Sabre-Delivery address,phone number and email issue
    End If
    
    If M0_flag = True Then
        Query = ""
        Select Case PreidNO_Long
            Case 0
                Query = "update SabreM0 set IU0TPR = '" & Trim(ln) & "' where UPLOADNO = " & M0_UploadNo & " "
            Case 1 To 5
                Query = "update SabreM0 set IU0BL1 =IU0BL1 + '" & Replace(Trim(ln), "'", "''") & vbCrLf & "' where UPLOADNO = " & M0_UploadNo & " "
            Case 6 To 10
                Query = "update SabreM0 set IU0BL2 =IU0BL2 + '" & Replace(Trim(ln), "'", "''") & vbCrLf & "' where UPLOADNO = " & M0_UploadNo & " "
            Case 11
                Query = "update SabreM0 set IU0PH1 = '" & Replace(Trim(ln), "'", "''") & "' where UPLOADNO = " & M0_UploadNo & " "
            Case 12
                Query = "update SabreM0 set IU0PH2 = '" & Replace(Trim(ln), "'", "''") & "' where UPLOADNO = " & M0_UploadNo & " "
            Case 13
                Query = "update SabreM0 set IU0PH3 = '" & Replace(Trim(ln), "'", "''") & "' where UPLOADNO = " & M0_UploadNo & " "
            Case 14
                Query = "update SabreM0 set IU0RCV = '" & Replace(Trim(ln), "'", "''") & "' where UPLOADNO = " & M0_UploadNo & " "
        End Select
        If Trim(Query) <> "" Then
            dbCompany.Execute Query
        End If
        If PreidNO_Long = 14 Then
            M0_flag = False
        End If
        PreidNO_Long = PreidNO_Long + 1
        'by Abhi on 10-Mar-2015 for caseid 5016 Sabre-Delivery address,phone number and email issue
        If M0_flag = True Then
            Preid_String = "AA"
        End If
        'by Abhi on 10-Mar-2015 for caseid 5016 Sabre-Delivery address,phone number and email issue
    End If
    If M1_flag = True And j < 5 Then
        
        If j = 0 Then
            Query = "update SabreM1 set IU1M3C = '" & Trim(ln) & "' where UPLOADNO = " & M1_UploadNo & " "
        ElseIf j = 1 Then
            'by Abhi on 23-Mar-2011 for caseid 1648 Sabre - Passenger line connection to ticket line - Logic changed to IU5PTY to IU1M5C
            'Query = "update SabreM1 set IU1M5C = '" & Trim(ln) & "' where UPLOADNO = " & M1_UploadNo & ""
            Query = "update SabreM1 set IU1M5C = '" & Right(ln, 2) & "' where UPLOADNO = " & M1_UploadNo & " AND IU1PNO = " & vIU1PNO_Long
        ElseIf j = 2 Then
            Query = "update SabreM1 set IU1M7C = '" & Trim(ln) & "' where UPLOADNO = " & M1_UploadNo & ""
        ElseIf j = 3 Then
            Query = "update SabreM1 set IU1M8C = '" & Trim(ln) & "' where UPLOADNO = " & M1_UploadNo & ""
        ElseIf j = 4 Then
            Query = "update SabreM1 set IU1M9C = '" & Trim(ln) & "' where UPLOADNO = " & M1_UploadNo & ""
        End If
        dbCompany.Execute Query
        If j = 4 Then
            M1_flag = False
        End If
        j = j + 1
    
    End If
    
    If M2_flag = True And i < 11 Then
        
        If i = 0 Then
            Query = "update SabreM2 set IU2NM4 = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & " "
        ElseIf i = 1 Then
            Query = "update SabreM2 set IU2NM6 = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 2 Then
            Query = "update SabreM2 set IU2FOP = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 3 Then
            Query = "update SabreM2 set IU2ORG = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 4 Then
            Query = "update SabreM2 set IU2IIE = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 5 Then
            Query = "update SabreM2 set IU2END = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 6 Then
            Query = "update SabreM2 set IU2JIC = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 7 Then
            Query = "update SabreM2 set IU2NRT = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 8 Then
            Query = "update SabreM2 set IU2REA = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 9 Then
            Query = "update SabreM2 set IU2DCD = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        ElseIf i = 10 Then
            Query = "update SabreM2 set IU2SCN = '" & Trim(ln) & "' where UPLOADNO = " & M2_UploadNo & ""
        End If
        dbCompany.Execute Query
        If i = 10 Then
            M2_flag = False
        End If
        i = i + 1
    End If
    
    If M3_flag = True And id <> "M4" And id <> "M5" And id <> "M6" And id <> "M7" And id <> "M9" Then
        If k = 0 Then
            Query = "update SabreM3 set IU3RRC = '" & Trim(ln) & "' where UPLOADNO = " & M3_UploadNo & " and IU3PC2 = 'HHL'"
        ElseIf k = 1 Then
            Query = "update SabreM3 set IU3HCP = '" & Trim(ln) & "' where UPLOADNO = " & M3_UploadNo & " and IU3PC2 = 'HHL'"
        End If
        k = k + 1
        If k = 2 Then
            k = 0
            M3_flag = False
        End If
        dbCompany.Execute Query
    Else
        M3_flag = False
    End If
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Public Function fnMidValue(str As String, Start As Integer, Ending As Integer) As String
    fnMidValue = Mid(str, Start, Ending)
End Function
Private Sub Form_Load()
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'On Error GoTo Note
    Dim rsSelect As New ADODB.Recordset
    Me.Icon = FMain.Icon
    If rsSelect.State = 1 Then rsSelect.Close
    'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsSelect.Open "Select SABREUPLOADDIRNAME,SABREDESTDIRNAME From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    'rsSelect.Open "Select SABREUPLOADDIRNAME,SABREDESTDIRNAME From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
    'If rsSelect.EOF = False Then
    '    txtDirName = IIf(IsNull(rsSelect!SABREUPLOADDIRNAME), "", rsSelect!SABREUPLOADDIRNAME)
    '    txtRelLocation = IIf(IsNull(rsSelect!SABREDESTDIRNAME), "", rsSelect!SABREDESTDIRNAME)
    'End If
    'rsSelect.Close
    'If rsSelect.State = 1 Then rsSelect.Close
    ''by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    ''rsSelect.Open "Select SABRESTATUS From [File]", dbCompany, adOpenDynamic, adLockBatchOptimistic
    'rsSelect.Open "Select SABRESTATUS From [File] WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
    'chkSabre.value = IIf((IsNull(rsSelect!SABRESTATUS) = True), 0, rsSelect!SABRESTATUS)
    txtDirName = txtSource
    txtRelLocation = txtDest
    chkSabre.value = chkSabre
    'by Abhi on 23-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
    File1.Path = txtDirName.Text
    File1.Pattern = FMain.txtSabreExt
    Exit Sub
'by Abhi on 13-Apr-2010 for caseid 1302 begin trans for PenGDS
'Note:
'    If Me.Visible = True Then Unload Me
End Sub
Public Function fnChecktheDatabase(CheckStr As String) As Boolean
Dim rsSelect As New ADODB.Recordset
Dim RecLoc As String

RecLoc = fnMidValue(CheckStr, 54, 8)
If rsSelect.State = 1 Then rsSelect.Close
'by Abhi on 20-Sep-2013 for caseid 3390 Optimisation for reading path and settings from file table
'rsSelect.Open "Select * From SabreM0 Where IU0PNR='" & Trim(RecLoc) & "'", dbCompany, adOpenDynamic, adLockBatchOptimistic
rsSelect.Open "Select IU0PNR From SabreM0 Where IU0PNR='" & Trim(RecLoc) & "'", dbCompany, adOpenForwardOnly, adLockReadOnly
If rsSelect.EOF = False Then
    fnChecktheDatabase = True
Else
    fnChecktheDatabase = False
End If
rsSelect.Close
End Function

Private Function RecM9_PENLINEPassenger(Data1, ByVal PNO As Long) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
        Coll.Add "", "Details1"
        Coll.Add "", "Details2"
        Coll.Add Mid(Data1, 3, 2), "ITNO"
        Coll.Add "", "PEN"
        Coll.Add "", "AC"
        Coll.Add "", "FD"
        Coll.Add "", "PO"
        Coll.Add "", "CC"
        Coll.Add "", "BR"
        
    'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
    'Temp1 = ExtractBetween(Data, "P" & PNO & "E-", "\")
    'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
    'Temp1 = ExtractBetween(Data, PENLINEID_String & "E" & PNO & "-", "\")
    Temp1 = ExtractBetween(Data, "E" & PNO & "-", "\")
        'by Abhi on 08-Mar-2018 for caseid 8331 Amadeus,Worldspan and Sabre - Pengds penline email replacement
        'Temp1 = Replace(Temp1, "*", "@")
        Temp1 = PenlineEmailReplacement(Temp1)
        'by Abhi on 08-Mar-2018 for caseid 8331 Amadeus,Worldspan and Sabre - Pengds penline email replacement
        'by Abhi on 11-Oct-2017 for caseid 7862 Passenger Email id validation in folder should also save multiple email id with coma separator
        'Coll.Add Temp1, "PEMAIL"
        Coll.Add Left(Temp1, 300), "PEMAIL"
        'by Abhi on 11-Oct-2017 for caseid 7862 Passenger Email id validation in folder should also save multiple email id with coma separator
    'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
    'Temp1 = ExtractBetween(Data, "P" & PNO & "T-", "\")
    'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
    'Temp1 = ExtractBetween(Data, PENLINEID_String & "T" & PNO & "-", "\")
    Temp1 = ExtractBetween(Data, "T" & PNO & "-", "\")
        Coll.Add Temp1, "PTELE"
        Coll.Add PNO, "PNO"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Temp1 = ExtractBetween(Data, "DOB" & PNO & "-", "\")
        Coll.Add Trim(Temp1), "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
Data1 = Data
Set RecM9_PENLINEPassenger = Coll
End Function

Private Function ExtractBetween(Data, startText, endText, Optional AllifNoValue As Boolean = True)
Dim aa, ab, ac
Dim result
result = ""
aa = InStr(1, Data, startText, vbTextCompare)
ac = IIf(aa > 0, aa, 1)
'by Abhi on 10-Mar-2015 for caseid 5016 Sabre-Delivery address,phone number and email issue
ac = IIf(aa < Len(Data), aa + 1, 1)
'by Abhi on 10-Mar-2015 for caseid 5016 Sabre-Delivery address,phone number and email issue
ab = InStr(ac, Data, endText, vbTextCompare)
If AllifNoValue = True Then
    If ab = 0 Then ab = Len(Data) + 1
End If
If aa > 0 And ab > 0 And ab > aa Then

aa = aa + Len(startText)
'ab = ab + Len(endText)
result = Mid(Data, aa, (ab - aa))
Data = Replace(Data, startText & result & endText, "")
Data = Replace(Data, startText & result, "")
End If
ExtractBetween = result
End Function

Private Function SplitWithLengths(Data, ParamArray Lengths())
Dim Nos As Integer
Dim pos As Integer, j As Integer
Nos = UBound(Lengths)
Dim temp() As String
ReDim temp(Nos) As String
pos = 1

For j = 0 To Nos
    temp(j) = Mid(Data, pos, Lengths(j))
    pos = pos + Lengths(j)
    DoEvents
Next
SplitWithLengths = temp
End Function

'by Abhi on 10-Mar-2015 for caseid 5016 Sabre-Delivery address,phone number and email issue
Private Function SplitWithLengthsPlus(Data, ParamArray Lengths())
On Error Resume Next
Dim Nos As Integer
Dim pos As Integer, j As Integer
Nos = UBound(Lengths)
Dim temp() As String
ReDim temp(Nos + 1) As String
pos = 1

For j = 0 To Nos
    temp(j) = Mid(Data, pos, Lengths(j))
    pos = pos + Lengths(j)
    DoEvents
Next
If Len(Data) > 0 Then
    temp(Nos + 1) = Mid(Data, pos, SkipNull(Len(Data) - pos + 1))
End If
SplitWithLengthsPlus = temp
End Function
'by Abhi on 10-Mar-2015 for caseid 5016 Sabre-Delivery address,phone number and email issue

'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
Private Function RecM9_PENAIRTKT(ByVal Data1) As Collection
Dim Splited
'Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
'splited = SplitForce(Data, "/", 9)
Coll.Add Mid(Data1, 3, 2), "ITNO"

Temp1 = ExtractBetween(Data, "PAX-", "/")
    Coll.Add Left(Temp1, 150), "PAX"
Temp1 = ExtractBetween(Data, "TKT-", "/")
    Coll.Add Left(Temp1, 50), "TKT"
Temp1 = ExtractBetween(Data, "DATE-", "/")
    Coll.Add Left(Temp1, 9), "DATE"

Data1 = Data
Set RecM9_PENAIRTKT = Coll
End Function
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
