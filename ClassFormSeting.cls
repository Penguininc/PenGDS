VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFormSeting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private nFormHeight      As Integer
Private nFormWidth       As Integer
Private nNumOfControls   As Integer
Private nTop()           As Integer
Private nLeft()          As Integer
Private nHeight()        As Integer
Private nWidth()         As Integer
Private nFontSize()      As Integer
Private nRightMargin()   As Integer
Private bFirstTime       As Boolean

Public Sub Init(frm As Form, Optional nWindState As Variant)
   Dim i          As Integer
   Dim bWinMax    As Boolean
  
   bWinMax = Not IsMissing(nWindState)
  
   nFormHeight = frm.Height
   nFormWidth = frm.Width
   nNumOfControls = frm.Controls.Count - 1
   bFirstTime = True
   ReDim nTop(nNumOfControls)
   ReDim nLeft(nNumOfControls)
   ReDim nHeight(nNumOfControls)
   ReDim nWidth(nNumOfControls)
   ReDim nFontSize(nNumOfControls)
  
   ReDim nRightMargin(nNumOfControls)
   On Error Resume Next
   For i = 0 To nNumOfControls
      If TypeOf frm.Controls(i) Is Line Then
         nTop(i) = frm.Controls(i).Y1
         nLeft(i) = frm.Controls(i).X1
         nHeight(i) = frm.Controls(i).Y2
         nWidth(i) = frm.Controls(i).X2
      Else
         nTop(i) = frm.Controls(i).Top
         nLeft(i) = frm.Controls(i).Left
         nHeight(i) = frm.Controls(i).Height
         nWidth(i) = frm.Controls(i).Width
         nFontSize(i) = frm.FontSize
         nRightMargin(i) = frm.Controls(i).RightMargin
      End If
      DoEvents
   Next
  
   If bWinMax Or frm.WindowState = 2 Then ' maxim
      frm.Height = Screen.Height
      frm.Width = Screen.Width
   Else
      frm.Height = frm.Height * Screen.Height / Screen.Height
      frm.Width = frm.Width * Screen.Width / Screen.Width
   End If
  
   bFirstTime = True
  
End Sub
Public Sub FormResize(frm As Form)
  
   Dim i             As Integer
   Dim nCaptionSize  As Integer
   Dim dRatioX       As Double
   Dim dRatioY       As Double
   Dim nSaveRedraw   As Long
  
   On Error Resume Next
   nSaveRedraw = frm.AutoRedraw
  
   frm.AutoRedraw = True
  
   If bFirstTime Then
      bFirstTime = False
      Exit Sub
   End If
  
   If frm.Height < nFormHeight / 2 Then frm.Height = nFormHeight / 2
  
   If frm.Width < nFormWidth / 2 Then frm.Width = nFormWidth / 2
   nCaptionSize = 400
   dRatioY = 1 ' * (nFormHeight - nCaptionSize) / (frm.Height - nCaptionSize)
   dRatioX = 1 ' * nFormWidth / frm.Width
   On Error Resume Next ' for comboboxes, timeres and other nonsizible controls
  
   For i = 0 To nNumOfControls
      If TypeOf frm.Controls(i) Is Line Then
         frm.Controls(i).Y1 = Int(nTop(i) / dRatioY)
         frm.Controls(i).X1 = Int(nLeft(i) / dRatioX)
         frm.Controls(i).Y2 = Int(nHeight(i) / dRatioY)
         frm.Controls(i).X2 = Int(nWidth(i) / dRatioX)
      Else
         frm.Controls(i).Top = Int(nTop(i) / dRatioY)
         frm.Controls(i).Left = Int(nLeft(i) / dRatioX)
         frm.Controls(i).Height = Int(nHeight(i) / dRatioY)
         frm.Controls(i).Width = Int(nWidth(i) / dRatioX)
         frm.Controls(i).FontSize = Int(nFontSize(i) / dRatioX) + Int(nFontSize(i) / dRatioX) Mod 2
         frm.Controls(i).RightMargin = Int(nRightMargin(i) / dRatioY)
      End If
      DoEvents
   Next
  
   frm.AutoRedraw = nSaveRedraw
  
End Sub

'Public Sub SetFormDetails(mForm As Form, mStBarPannel As Integer)
'Dim lblHeadItem  As ComboItem
'Set lblHeadItem = mForm.lblHead.ComboItems.Add(1, , "Customer", 3)
'Set lblHeadItem = mForm.lblHead.ComboItems.Add(2, , "Supplier", 4)
'Set lblHeadItem = mForm.lblHead.ComboItems.Add(3, , "Nominal", 5)
'Set lblHeadItem = mForm.lblHead.ComboItems.Add(4, , "Bank", 9)
''Set lblHeadItem = mForm.lblHead.ComboItems.Add(5, , "Ageing Receivable", 10)
''Set lblHeadItem = mForm.lblHead.ComboItems.Add(6, , "Ageing Payable", 10)
'
''Set lblHeadItem = mForm.lblHead.ComboItems.Add()
'SendKeys "{HOME}+{END}"
'mForm.Caption = "Exploring - " & FMain.StBar.Panels(4).Text & " . . ."
'mForm.Icon = FMain.Icon
'
'Select Case FMain.StBar.Panels(mStBarPannel).Text
'    Case "Customer"
'        FList.lblHead.SelectedItem = mForm.lblHead.ComboItems.Item(1)
'    Case "Supplier"
'        FList.lblHead.SelectedItem = mForm.lblHead.ComboItems.Item(2)
'    Case "Nominal"
'        FList.lblHead.SelectedItem = mForm.lblHead.ComboItems.Item(3)
'    Case "Bank"
'        FList.lblHead.SelectedItem = mForm.lblHead.ComboItems.Item(4)
'End Select
'
'mForm.pct1.Picture = mForm.ImageList2.ListImages(9).Picture: mForm.pct1.BackColor = mForm.BackColor
''mForm.pct2.Picture = mForm.ImageList2.ListImages(8).Picture: mForm.pct2.BackColor = mForm.BackColor: mForm.pct2.Left = mForm.lblHead.Left + mForm.lblHead.Width + 100: mForm.lbl1(1).Left = mForm.pct2.Left - 5000
'End Sub

'Public Sub SetCalculator(mCalcTop As Long, mCalcLeft As Long, mForm As Form, mTxt As String)
'mFormName = mForm.Name
'Calculator.Top = mCalcTop: Calculator.Left = mCalcLeft
'Calculator.Show 1
'End Sub

'by Abhi on 03-Nov-2010 for caseid 1536 Selected date should be show in month view
'by Abhi on 17-Nov-2014 for caseid 4710 Month end closing advanced for ACR and ACP closing
'Public Sub SetMonthView(mMonthVTop As Long, mMonthVLeft As Long, mForm As Form, mTxt As String, Optional pSelectedDate_String As String)
Public Sub SetMonthView(mMonthVTop As Long, mMonthVLeft As Long, mForm As Form, mTxt As String, Optional pSelectedDate_String As String, Optional ByVal pLedgerType_MonthEndClosingLedgerType As MonthEndClosingLedgerType = MonthEndClosingLedgerTypeGeneral)
'by Abhi on 17-Nov-2014 for caseid 4710 Month end closing advanced for ACR and ACP closing
mFormName = mForm.Name
Set MonthView.PFrm = mForm
mTxtCallReturn = mTxt
MonthView.Top = mMonthVTop: MonthView.Left = mMonthVLeft
'by Abhi on 03-Nov-2010 for caseid 1536 Selected date should be show in month view
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
'mSelectedDate_String = DateFormat(pSelectedDate_String)
If Trim(pSelectedDate_String) <> "" Then
    pSelectedDate_String = DateFormat(pSelectedDate_String)
    If mFnCheckValidDate(pSelectedDate_String, , False) = True Then
        mSelectedDate_String = DateFormat(pSelectedDate_String)
    End If
End If
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date

'by Abhi on 17-Nov-2014 for caseid 4710 Month end closing advanced for ACR and ACP closing
MonthView.fLedgerType_MonthEndClosingLedgerType = pLedgerType_MonthEndClosingLedgerType
'by Abhi on 17-Nov-2014 for caseid 4710 Month end closing advanced for ACR and ACP closing
MonthView.Show 1
End Sub

