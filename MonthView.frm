VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MonthView 
   BorderStyle     =   0  'None
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2505
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox mskDATE 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "dd/mmm/yyyyy"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   91619329
      CurrentDate     =   38311
   End
   Begin VB.Shape Shape1 
      Height          =   2445
      Left            =   0
      Top             =   0
      Width           =   2835
   End
End
Attribute VB_Name = "MonthView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PFrm As Form
'by Abhi on 17-Nov-2014 for caseid 4710 Month end closing advanced for ACR and ACP closing
Public fLedgerType_MonthEndClosingLedgerType As MonthEndClosingLedgerType
'by Abhi on 17-Nov-2014 for caseid 4710 Month end closing advanced for ACR and ACP closing

Private Sub Form_Activate()
On Error Resume Next
MonthView1.Year = Year(DateAppTimeZone)
MonthView1.Month = Month(DateAppTimeZone)
MonthView1.Day = Day(DateAppTimeZone)
'by Abhi on 03-Nov-2010 for caseid 1536 Selected date should be show in month view
If IsDate(mSelectedDate_String) = True Then
    MonthView1.value = mSelectedDate_String
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
'mskDATE.text = dateFormat(MonthView1.Value)
'If Not mFnDateValidated(dateFormat(MonthView1.Value)) Then MonthView1.SetFocus: Exit Sub
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
'If fnIsFinancial = True Then
''If (mFormName = "FPurch2" And mTxtCallReturn = "DtpDate") Then
'    'by Abhi on 17-Nov-2014 for caseid 4710 Month end closing advanced for ACR and ACP closing
'    'If Not mFnDateValidated(dateFormat(MonthView1.value)) Then
'    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
'    If Not mFnDateValidated(DateFormat(MonthView1.value), fLedgerType_MonthEndClosingLedgerType) Then
'    'by Abhi on 17-Nov-2014 for caseid 4710 Month end closing advanced for ACR and ACP closing
'        Exit Sub
'    Else
'        mskDATE.Text = DateFormat(MonthView1.value)
'        Unload Me
'    End If
'Else
'    mskDATE.Text = DateFormat(MonthView1.value)
'    Unload Me
'End If
mskDATE.Text = DateFormat(MonthView1.value)
Unload Me
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
End Sub
Private Sub MonthView1_DblClick()
'mskDATE.Text = dateFormat(MonthView1.Value)
'If Not mFnDateValidated(dateFormat(MonthView1.Value)) Then
'    Exit Sub
'Else
'    Unload Me
'End If
'MonthView1.SetFocus: Exit Sub
End Sub
Private Sub MonthView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        MonthView1_DateClick (DateAppTimeZone)
    End If
End Sub

Private Sub MonthView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = SHIFT_MASK + CTRL_MASK + ALT_MASK Then
        If Button = vbRightButton Then
            Call ShowFormName(Me)
        End If
    End If
End Sub

Private Sub mskDATE_Change()
On Error Resume Next
Select Case mFormName
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    Case "FMain"
        If mTxtCallReturn = "OutOfBookingsDateAmadeusText" Then PFrm.OutOfBookingsDateAmadeusText.Text = DateFormat(mskDATE)
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
End Select
End Sub

Function fnIsFinancial() As Boolean
Dim vblIsFin As Boolean
vblIsFin = False
    'By Sandhya    on   07/10/2014   for case id : 4606
    'By Sandhya    on   18/11/2014   for case id : 4707
'    If mFormName = "SupplierOSbyFolderCombinedFilterForm" Then
    If mFormName = "SupplierOSbyFolderCombinedFilterForm" Or mFormName = "FolderActivityReportbyProductFilterForm" Then
        fnIsFinancial = False
        Exit Function
    End If
    
    'By Sandhya    on   05/05/2014   for case id : 3815
    If mFormName = "StaffPerformanceSummary" Or mFormName = "SlidingCommnReportFilter" Or mFormName = "BranchCommissionReportForm" Then
        fnIsFinancial = False
        Exit Function
    End If
    
    
    'By Susan on 02/09/2014  for case id: 4386
    If mFormName = "CustOutstandingAsOnDateForm" Then
        fnIsFinancial = False
        Exit Function
    End If
    'by jiji on /Sep/ for caseid 9085
    If mFormName = "FFolderInfoCollectionDetailsForm" Then
        fnIsFinancial = False
        Exit Function
    End If
    'by jiji on /Sep/ for caseid 9085
    
''''        Sandhya   19/02/2008
'If mFormName <> "FPurch2" And mTxtCallReturn <> "DtpBDate" And mFormName <> "FFolder" Then
'By jiji on 24/jan/2013 for caseid 2843
'If mFormName <> "FPurch2" And mFormName <> "FPassenger" And mTxtCallReturn <> "DtpBDate" And mFormName <> "FFolder" And mTxtCallReturn <> "StaffCommissionFrom" And mTxtCallReturn <> "StaffCommissionTo" Then
If mFormName <> "FPurch2" And mFormName <> "FPassenger" And mTxtCallReturn <> "DtpBDate" And mFormName <> "FFolder" And mTxtCallReturn <> "StaffCommissionFrom" And mTxtCallReturn <> "StaffCommissionTo" And mFormName <> "ExternalInterfaceForm" Then
    vblIsFin = True
    If mFormName <> "FAtolestimate" And mFormName <> "FInvTravelRPt" And mFormName <> "FInvTravelRPt2" And mFormName <> "FInvTravelRPt3" And mFormName <> "FHotelBView" And (mFormName <> "FInvTravel" And (mTxtCallReturn <> "txtBSPReStmtDateFrom" Or mTxtCallReturn <> "TxtBSPReStmtDateTo")) And mFormName <> "FRoomRates" And (mFormName <> "FList" And (mTxtCallReturn <> "txtChqReStmtDateFrom" Or mTxtCallReturn <> "TxtChqReStmtDateTo")) And mFormName <> "FProject" And mFormName <> "FAgentCommission" And (mFormName <> "FInvTravelRPt4" And (mTxtCallReturn <> "txtBSPReStmtDateFrom" Or mTxtCallReturn <> "TxtBSPReStmtDateTo")) Then
        vblIsFin = True
    Else
        vblIsFin = False
    End If
''''        Sandhya   24/04/2008   case id : 494
    ''''Or mTxtCallReturn = "CRFromDateTxt" Or mTxtCallReturn = "CRToDate" add for Case Id 971 By Jinu on 15-Aug-2009
    If (mFormName = "FChequeregister" And (mTxtCallReturn = "dtpDocDateFrom" Or mTxtCallReturn = "dtpDocDateTo" Or mTxtCallReturn = "dtpChqDateFrom" Or mTxtCallReturn = "dtpChqDateTo")) Or (mFormName = "FAgeIng" And (mTxtCallReturn = "dtpChkDate" Or mTxtCallReturn = "dtpNCDate")) Or (mFormName = "FBSPReconcile" And mTxtCallReturn <> "txtIDATE") Or (mFormName = "FCreditRecon" And (mTxtCallReturn = "txtDateFrom" Or mTxtCallReturn = "txtDateTo" Or mTxtCallReturn = "CRFromDateTxt" Or mTxtCallReturn = "CRToDate")) Or (mFormName = "FChequePayinslip" And mTxtCallReturn <> "txtDateBanked") Then
'    If (mFormName = "FAgeIng" And (mTxtCallReturn = "dtpChkDate" Or mTxtCallReturn = "dtpNCDate")) Or (mFormName = "FBSPReconcile" And mTxtCallReturn <> "txtIDATE") Or (mFormName = "FCreditRecon" And (mTxtCallReturn = "txtDateFrom" Or mTxtCallReturn = "txtDateTo")) Or (mFormName = "FChequePayinslip" And mTxtCallReturn <> "txtDateBanked") Then
        vblIsFin = False
    End If
    If mFormName = "FCustSalesAnlysRPT" Or mFormName = "FIataRpt" Or mFormName = "FInvTravelRPt5" Or mFormName = "FTopCityPairRpt" Or mFormName = "LedgerFilterForm" Or mFormName = "WriteoffFilterForm" Or mFormName = "VATFilterForm" Or mFormName = "VATUKFilterForm" Or mFormName = "JournalFilterForm" Or mFormName = "LedgerFilterForm" Or mFormName = "AirsegmentReportForm" Then
        vblIsFin = False
    End If
    ''''' for caseId 971 by Jinu on 15-Aug-09
    If mFormName = "FBankReconcile" And mTxtCallReturn = "ChequeDateFromTxt" Or mTxtCallReturn = "ChequeDateToTxt" Then
            vblIsFin = False
    End If ''''' for caseId 971 by Jinu on 15-Aug-09
      ''By jiji on 15/05/2012 for case id=2250
    If mFormName = "FolderCustomerListRPTForm" And mTxtCallReturn = "FolderCustListDateFromText" Or mTxtCallReturn = "FolderCustListDateToText" Then
        vblIsFin = False
    End If
    ''''caseid 1038 binu 05/12/09
    If mFormName = "SlidingCommissionCalculation" And mTxtCallReturn = "DateFromTxt" Or mTxtCallReturn = "DateToTxt" Then
        vblIsFin = False
    End If
    ''''''''''''''''''''''''''
    
    'by Abhi on 03-Nov-2010 for caseid 1535 Step1 Date filter
    If mFormName = "EPDQFilterForm" Or mFormName = "CustomerStatementRPTForm" Or mFormName = "CustStmtWithPassDetailsRpt" Or mFormName = "InvoiceListingFilterForm" Or mFormName = "AirTicketReportsForm" Or mFormName = "FAgeIngFilterForm" Then ''''Case Id 929 jinu on 24-06-09
        vblIsFin = False
    End If ''''Case Id 929 jinu on 24-06-09
    '''''For CaseID 953 by Jinu on 10-10-09
    If mFormName = "FBank" And mTxtCallReturn = "SpChequeDateBtn" Then
        vblIsFin = False
    End If
    'by  Sandhya  on  07/05/2015   for  caseid : 4857  vvvvvvvvvvvvvvvvvv
    If mFormName = "FBank" And (mTxtCallReturn = "SBRChequeDateTxt" Or mTxtCallReturn = "dtpChkDate") Then
        vblIsFin = False
    End If
    'by  Sandhya  on  07/05/2015   for  caseid : 4857  ^^^^^^^^^^^^^^^^^^
     
     '''''For CaseID 953 by Jinu on 10-10-09
    ''''for CaseId 1057 on 17-12-2009 jnu
    If mFormName = "FBranchCommission" Then
        If mTxtCallReturn = "BranchComm" Then
            vblIsFin = False
        ElseIf mTxtCallReturn = "BranchCommissionTo" Then
            vblIsFin = False
        End If
    End If
    
    If mFormName = "FInvoice" Or mFormName = "FInvPackage" Then
        If mTxtCallReturn = "dtpChkDate" Or mTxtCallReturn = "dtpNCDate" Then
            vblIsFin = False
        End If
    End If
    'by Abhi on 04-Sep-2010 for caseid 1415 remove financial year checking for Folder option date
    If mFormName = "FFolderOptions" Then
        'by Abhi on 14-May-2018 for caseid 8611 GDPRConsent, OptinDate field in Foldermaster
        'If mTxtCallReturn = "OptionDate" Then
        '    vblIsFin = False
        'End If
        vblIsFin = False
        'by Abhi on 14-May-2018 for caseid 8611 GDPRConsent, OptinDate field in Foldermaster
    End If
    ''''for CaseId 1057 on 17-12-2009 jnu

    'Sandhya    Case id : 1575    24/12/2010
    'by Abhi on 22-Feb-2011 for caseid 1616 Option time Address Check and 3D check for Sagepay EPDQ
    If mFormName = "FolderPassengerListForm" Or mFormName = "FFolderEPayMsgBoxForm" Then
        vblIsFin = False
    ElseIf mTxtCallReturn = "AllocationDateText" Then 'Sandhya    Case id : 1619    09/03/2011
        vblIsFin = False
    End If
    'Sandhya    Case id : 2591    11/12/2012    'DailyReportsAirTicketsFilterForm
    'Sandhya    Case id : 2644    05/12/2012    'StaffBookingReportForm
    'Sandhya    Case id : 1629    07/04/2011
    'Sandhya    Case id : 1491    08/04/2011
    'Sandhya    Case id : 1804    01/07/2011    'CustSalesSummaryFilterForm
    'by Abhi on 14-Jun-2011 for caseid 1784 Folder Query Performance Report
    'by Abhi on 09-Oct-2012 for caseid 2534 remove the financial yera check from ATOL Return-Quarterly and ATOL Report(Detailed by Invoice)
    'By jiji on 23/Oct/2012 for Please remove the financial year check from this report-Date filter.
'    If mFormName = "SuppStmtWithPassengerDetailsForm" Or mFormName = "AutoInvoiceEmailForm" Or mFormName = "CustSalesSummaryFilterForm" Or mFormName = "FolderQueryPerformanceReportFilterForm" Or mFormName = "FFolderRailSegDetailsForm" _
'    Or mFormName = "ATOLReturnQuarterlyForm" Or mFormName = "FAtolReportStandard" Then
    'by Abhi on 01-Jul-2013 for caseid 3197 folder search change
    If mFormName = "SuppStmtWithPassengerDetailsForm" Or mFormName = "AutoInvoiceEmailForm" Or mFormName = "CustSalesSummaryFilterForm" Or mFormName = "FolderQueryPerformanceReportFilterForm" Or mFormName = "FFolderRailSegDetailsForm" _
     Or mFormName = "ATOLReturnQuarterlyForm" Or mFormName = "FAtolReportStandard" Or mFormName = "TopDestinationReport" Or mFormName = "StaffBookingReportForm" Or mFormName = "DailyReportsAirTicketsFilterForm" Or mFormName = "FTravelSearch" Then
        vblIsFin = False
    End If
    If mFormName = "FPurchMultiInv" Then 'By   Sandhya  for  Case id : 3085    on   02/05/2013
        If mTxtCallReturn = "PODateFrom" Or mTxtCallReturn = "PODateTo" Then
            vblIsFin = False
        End If
    End If
    'by Abhi on 26-Apr-2013 for caseid 3070 Trade invoice optimization for loading
    If mFormName = "FPurchTradeInvoiceBulkApproval" Then
        If mTxtCallReturn = "PODateFrom" Or mTxtCallReturn = "PODateTo" Then
            vblIsFin = False
        End If
    End If
    'If mFormName = "ProductAnalysisHotelReportForm" Then 'By  Sandhya   on  08/05/2013   for Case id : 3096
    If mFormName = "ProductAnalysisHotelReportForm" Or mFormName = "ProductAnalysisAirReportForm" Then    'By jiji on 14/Jun/2013 for caseid 3136
        'If mTxtCallReturn = "CrDateFromText" Or mTxtCallReturn = "CrDateToText" Then
         If mTxtCallReturn = "CrDateFromText" Or mTxtCallReturn = "CrDateToText" Or mTxtCallReturn = "ChkInDteFromText" Or mTxtCallReturn = "ChkInDteToText" Or mTxtCallReturn = "ChkOutDteFromText" Or mTxtCallReturn = "ChkOutDteToText" Then 'By jiji on 21/Aug/2013 for caseid 3297
            vblIsFin = False
        End If
    End If
    'by Abhi on 09-Nov-2013 for caseid 3501 New BSP Report
    'By jiji on 28/Nov/2013 for caseid 3544
    'If mFormName = "FolderNotInvoicedFilterForm" Or mFormName = "FolderInvoicedNotyetApprovedFilterForm" Or mFormName = "FolderApprovedNotyetPaidFilterForm" Or mFormName = "BSPReStmtForm" Then  'By Sandhya on 28/Sep/2013 for caseid 3378
    If mFormName = "FolderNotInvoicedFilterForm" Or mFormName = "FolderInvoicedNotyetApprovedFilterForm" Or mFormName = "FolderApprovedNotyetPaidFilterForm" Or mFormName = "BSPReStmtForm" Or mFormName = "DraftagentbillinganalysisForm" Then  'By Sandhya on 28/Sep/2013 for caseid 3378
        vblIsFin = False
    End If
    
    'By Sandhya    on   25/03/2014    for case id : 3612
'    If mFormName = "ALTurnoverReportFilterForm" Then 'By Sandhya on 01/11/2013 for Case id : 2350
    If mFormName = "ALTurnoverReportFilterForm" Or mFormName = "CustomerRecPayFilterForm" Then
        vblIsFin = False
    End If
    
    If mFormName = "VATRPTForm" Then 'By Sandhya    on   22/11/2013    for case id : 3517
        vblIsFin = True
    End If
    'by Abhi on 20-Jul-2015 for caseid 5413 Birth Day in User Master
    'If mFormName = "FFolderRailSegsForm" Or mFormName = "FFolderRailDetailsForm" Then 'By Sandhya    on   12/02/2014    for case id : 3723
    'by Abhi on 10-Mar-2016 for caseid 4849 Balance due date for Supplier-Folder change
    'If mFormName = "FFolderRailSegsForm" Or mFormName = "FFolderRailDetailsForm" Or mFormName = "FUsers" Then  'By Sandhya    on   12/02/2014    for case id : 3723
    'by Abhi on 25-Mar-2016 for caseid 6155 Balance due date calculation for supplier-Air, Car, Tour, Rail, Cruise, Insurance, Others and Transfers
    'by Abhi on 04-Jun-2018 for caseid 8666 Folder form changes for GDPR
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    If mFormName = "FFolderRailSegsForm" Or mFormName = "FFolderRailDetailsForm" Or mFormName = "FUsers" Or mFormName = "FHotelCRS" Or mFormName = "FFolderAirTicketAdd" Or mFormName = "FFolderCarDetailsForm" _
        Or mFormName = "FFolderTourDetailsForm" Or mFormName = "FFolderCruiseDetailsForm" Or mFormName = "FFolderInsuranceDetailsForm" Or mFormName = "FFolderOthersPackageForm" _
        Or mFormName = "FFolderOtherDetailForm" Or mFormName = "FFolderTransfersDetailsForm" Or mFormName = "FFolderGDPRDetails" Or mFormName = "FTrParameter" Then        'By Sandhya    on   12/02/2014    for case id : 3723
    'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
    'by Abhi on 04-Jun-2018 for caseid 8666 Folder form changes for GDPR
    'by Abhi on 25-Mar-2016 for caseid 6155 Balance due date calculation for supplier-Air, Car, Tour, Rail, Cruise, Insurance, Others and Transfers
    'by Abhi on 10-Mar-2016 for caseid 4849 Balance due date for Supplier-Folder change
    'by Abhi on 20-Jul-2015 for caseid 5413 Birth Day in User Master
        vblIsFin = False
    End If
End If
fnIsFinancial = vblIsFin
End Function
