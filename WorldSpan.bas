Attribute VB_Name = "WorldSpan"
Dim PassengerToAirFairID As Long
Dim PassengerToAirFairFlag As Boolean
Public WorldSpan_AirFareSLNO As Long
'by Abhi on 10-Nov-2009 for VLOCATOR in AirSegDetails
Dim mITNRYSEGNO_Long As Long
'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
Dim fOperatingITNRYSEGNO_Long As Long
Dim fOperatingITNRYSEGNO_Table_String As String
'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan

Private Function SplitDelims(Data, ParamArray Delims())
Dim temp, mData, result(), mCount
mData = Data
mCount = UBound(Delims)
ReDim result(mCount + 1)
For j = 0 To mCount
    temp = SplitFirstTwo(mData, Delims(j))
    mData = temp(1)
    result(j) = temp(0)
    DoEvents
Next
    result(j) = mData
SplitDelims = result
End Function

Private Function SplitTwo(Data, First As Long)
On Error Resume Next
Dim temp(1) As String
temp(0) = Left(Data, First)
temp(1) = Mid(Data, First + 1, Len(Data) - First)
SplitTwo = temp
End Function

Private Function SplitTwoReverse(Data, First As Long)
On Error Resume Next
Dim temp(1) As String
temp(1) = Right(Data, First)
temp(0) = Mid(Data, 1, Len(Data) - First)
SplitTwoReverse = temp
End Function

Private Function SplitFirstTwo(Data, find)
Dim temp
Dim tm(1) As String
temp = Split(CStr(Data), find, 2)
If UBound(temp) > -1 Then
tm(0) = temp(0)
End If
If UBound(temp) > 0 Then
tm(1) = temp(1)
End If

SplitFirstTwo = tm
End Function

Private Function SplitField(Data, Start As String, Finish As String)
Dim rStart, rMid, rEnd As String
Dim retn(2) As String
Dim temp, temp2, temp3
    temp = SplitFirstTwo(Data, "(")
    rStart = temp(0)
    temp2 = SplitFirstTwo(temp(1), ")")
    rMid = temp2(0)
    rEnd = temp2(1)
retn(0) = rStart
retn(1) = rMid
retn(2) = rEnd
SplitField = retn
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
temp(Nos + 1) = Mid(Data, pos, SkipNegative(Len(Data) - pos + 1))
SplitWithLengthsPlus = temp

End Function

Private Function SplitForce(Data, delimiter As String, minNos As Integer)
Dim temp() As String
Dim Nos As Integer
temp = Split(Data, delimiter)
Nos = UBound(temp) + 1
If Nos < minNos Then
ReDim Preserve temp(minNos)
End If
SplitForce = temp
End Function

Private Function SplitForcePlus(Data, delimiter As String, minNos As Integer)
Dim temp() As String
Dim Nos As Integer
temp = Split(Data, delimiter, minNos + 1)
Nos = UBound(temp) + 1
If Nos < minNos Then
ReDim Preserve temp(minNos)
End If
SplitForcePlus = temp
End Function
Private Function ExtractBetween(Data, startText, endText, Optional AllifNoValue As Boolean = True)
Dim aa, ab, ac
Dim result
result = ""
aa = InStr(1, Data, startText, vbTextCompare)
ac = IIf(aa > 0, aa, 1)
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

Private Function ExtractBetweenFrom(Data, startText, endText, Optional AllifNoValue As Boolean = True, Optional FromNoofChar As Long = 0)
Dim aa, ab, ac
Dim result
result = ""
aa = InStr(1, Data, startText, vbTextCompare)
aa = aa - FromNoofChar
ac = IIf(aa > 0, aa, 1)
ab = InStr(ac, Data, endText, vbTextCompare)
If AllifNoValue = True Then
    'by Abhi on 10-Oct-2015 for caseid 5645 PenGDS for Worldspan stuck due to missing backslash in PENFARE penline tag
    'If ab = 0 Then ab = Len(Data) + 1
    If ab = 0 Then
        ab = Len(Data) + 1
        endText = ""
    End If
    'by Abhi on 10-Oct-2015 for caseid 5645 PenGDS for Worldspan stuck due to missing backslash in PENFARE penline tag
End If
If aa > 0 And ab > 0 And ab > aa Then

'aa = aa + Len(startText)
'ab = ab + Len(endText)
result = Mid(Data, aa, (ab - aa))
Data = Replace(Data, startText & result & endText, "")
'Data = Replace(Data, startText & result, "")
Data = Replace(Data, result & endText, "")
End If
ExtractBetweenFrom = result
End Function
Private Function ExtractFrom(Data, startText)
Dim aa, ab, ac
Dim result
result = ""
aa = InStr(1, Data, startText, vbTextCompare)
ac = IIf(aa > 0, aa, 1)
If aa > 0 Then

aa = aa + Len(startText)
'ab = ab + Len(endText)
result = Mid(Data, aa)
Data = Replace(Data, startText & result, "")
End If
ExtractFrom = result
End Function

Private Function InsertData(Data As Collection, TableName As String, LineIdent As String, UploadNo As Long) As Boolean
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
On Error GoTo PENErr
Dim ErrNumber As String
Dim ErrDescription As String
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
Dim rs As New ADODB.Recordset
Dim j As Integer
'by Abhi on 23-Oct-2010 for caseid 1516 PenGDS Amadeus slow reading
'Rs.Open "Select * from " + TableName, dbCompany, adOpenDynamic, adLockPessimistic
'by Abhi on 12-Nov-2010 for caseid 1546 PenGDS Optimistic concurrency check failed
'Rs.Open "Select * from " + TableName & " WHERE UpLoadNo=" & UploadNo, dbCompany, adOpenForwardOnly, adLockOptimistic
rs.Open "Select * from " + TableName & " WHERE UpLoadNo=" & UploadNo, dbCompany, adOpenForwardOnly, adLockPessimistic
rs.AddNew
rs.Fields(0) = UploadNo
rs.Fields(1) = LineIdent
For j = 2 To rs.Fields.Count - 1
    rs.Fields(j) = Data(j - 1)
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    'If Len(Data(j - 1)) > rs.Fields(j).DefinedSize Then
    '    ErrDetails_String = vbCrLf & "Table=" & TableName & ", Field=" & rs.Fields(j).Name & ", FieldSize=" & rs.Fields(j).DefinedSize & "<>" & Len(Data(j - 1)) & "," & vbCrLf & Data(j - 1) & vbCrLf
    'End If
    'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    If Len(Data(j - 1)) > rs.Fields(j).DefinedSize Then
        ErrDetails_String = vbCrLf & "RecID=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(j).Name & ", FieldSize=" & rs.Fields(j).DefinedSize & "<>" & Len(Data(j - 1)) & "," & vbCrLf & Data(j - 1) & vbCrLf
    End If
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    DoEvents
Next
rs.Update
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
Exit Function
PENErr:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    'If ErrNumber = ErrNumber Then 'Subscript out of range
    If KeyExists(Data, j - 1) = False Then
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
        ErrDetails_String = vbCrLf & "RecID=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(j).Name & ", Field does not exist in Collection!," & vbCrLf
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    ElseIf Trim(ErrDetails_String) <> "" Then '-2147467259 Data provider or other service returned an E_FAIL status.
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    ElseIf Len(Data(j - 1)) > rs.Fields(j).DefinedSize Then '-2147217887 Multiple-step operation generated errors. Check each status value.
        ErrDetails_String = vbCrLf & "RecID=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(j).Name & ", FieldSize=" & rs.Fields(j).DefinedSize & "<>" & Len(Data(j - 1)) & "," & vbCrLf & Data(j - 1) & vbCrLf
    Else
        ErrDetails_String = vbCrLf & "RecID=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(j).Name & "," & vbCrLf & Data(j - 1) & vbCrLf
    End If
    Err.Raise ErrNumber, , ErrDescription
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
End Function

Private Function InsertDataByFieldName(Data As Collection, TableName As String, LineIdent As String, UploadNo As Long) As Boolean
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
On Error GoTo PENErr
Dim ErrNumber As String
Dim ErrDescription As String
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
Dim rs As New ADODB.Recordset
Dim temp
Dim j As Integer
'by Abhi on 23-Oct-2010 for caseid 1516 PenGDS Amadeus slow reading
'Rs.Open "Select * from " + TableName, dbCompany, adOpenDynamic, adLockPessimistic
'by Abhi on 12-Nov-2010 for caseid 1546 PenGDS Optimistic concurrency check failed
'Rs.Open "Select * from " + TableName & " WHERE UpLoadNo=" & UploadNo, dbCompany, adOpenForwardOnly, adLockOptimistic
rs.Open "Select * from " + TableName & " WHERE UpLoadNo=" & UploadNo, dbCompany, adOpenForwardOnly, adLockPessimistic
rs.AddNew
rs.Fields("UpLoadNo") = UploadNo
rs.Fields("RecID") = LineIdent
For j = 2 To rs.Fields.Count - 1
    temp = rs.Fields(j).Name
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'If Rs.Fields(temp).DefinedSize < Len(Data(temp)) Then
    '    Debug.Print TableName & "." & temp
    'End If
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    rs.Fields(temp) = Data(temp)
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    'If Len(Data(temp)) > rs.Fields(temp).DefinedSize Then
    '    ErrDetails_String = vbCrLf & "Table=" & TableName & ", Field=" & rs.Fields(temp).Name & ", FieldSize=" & rs.Fields(temp).DefinedSize & "<>" & Len(Data(temp)) & "," & vbCrLf & Data(temp) & vbCrLf
    'End If
    'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    'by Abhi on 09-Apr-2016 for caseid 6243 Pnr was Stuck in Pengds interface
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    If Len(Data(temp)) > rs.Fields(temp).DefinedSize Then '-2147217887 Multiple-step operation generated errors. Check each status value.
        ErrDetails_String = vbCrLf & "RecID=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(temp).Name & ", FieldSize=" & rs.Fields(temp).DefinedSize & "<>" & Len(Data(temp)) & "," & vbCrLf & Data(temp) & vbCrLf
    End If
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    DoEvents
Next
rs.Update
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
Exit Function
PENErr:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    If KeyExists(Data, temp) = False Then
        'ErrDetails_String = vbCrLf & "Table=" & TableName & ", Field=" & temp & ", Key Does NOT Exist!," & vbCrLf & Data(temp) & vbCrLf
        ErrDetails_String = vbCrLf & "RecID=" & LineIdent & ", Table=" & TableName & ", Field=" & temp & ", Field does not exist in Collection!," & vbCrLf
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    ElseIf Trim(ErrDetails_String) <> "" Then '-2147467259 Data provider or other service returned an E_FAIL status.
    'by Abhi on 12-Aug-2016 for caseid 6672 PenGDS Error Data provider or other service returned an E_FAIL status due to AmdLineKFT. FAREREMARKS field size
    ElseIf Len(Data(temp)) > rs.Fields(temp).DefinedSize Then '-2147217887 Multiple-step operation generated errors. Check each status value.
        ErrDetails_String = vbCrLf & "RecID=" & LineIdent & ", Table=" & TableName & ", Field=" & rs.Fields(temp).Name & ", FieldSize=" & rs.Fields(temp).DefinedSize & "<>" & Len(Data(temp)) & "," & vbCrLf & Data(temp) & vbCrLf
    Else
        ErrDetails_String = vbCrLf & "RecID=" & LineIdent & ", Table=" & TableName & ", Field=" & temp & "," & vbCrLf & Data(temp) & vbCrLf
    End If
    Err.Raise ErrNumber, , ErrDescription
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
End Function


Public Function PostLine(Line As String, UploadNo As Long, Optional FNAME = "") As String
Dim temp As String, NewLine As String
Dim Coll As New Collection
Dim CollPen As New Collection
Dim temp2nd As String

Dim vi As Long
temp = Left(Line, 1)
temp2nd = Left(Line, 2)
NewLine = Right(Line, SkipNegative(Len(Line) - 1))
Select Case temp
    Case "1"
        Set Coll = Rec1_PNRFileAddress(NewLine, FNAME)
        InsertData Coll, "WSPPNRADD", "1", UploadNo
    Case "2"
        Set Coll = Rec_ClientAcNo(NewLine)
        InsertDataByFieldName Coll, "WSPCLNTACNO", "2", UploadNo
    Case "6"
        Set Coll = Rec_PhoneContact(NewLine)
        InsertDataByFieldName Coll, "WSPPHONECONTACT", "6", UploadNo

    Case "3"
        Set Coll = Rec_BranchAgentSines(NewLine)
        InsertDataByFieldName Coll, "WSPAGNTSINE", temp, UploadNo
    Case "7"
        Set Coll = Rec_FormOfPayment(NewLine)
        InsertDataByFieldName Coll, "WSPFORMOFPYMNT", temp, UploadNo
    Case "8"
        Set Coll = Rec8_AirCommission(NewLine)
        InsertDataByFieldName Coll, "WSPAIRCOM", "8", UploadNo
    Case "9"
        Set Coll = Rec9_TicketingCarrier(NewLine)
        InsertData Coll, "WSPTCARRIER", "9", UploadNo
    Case "0", "*"
    '------------------------------------------------------------
        Dim temp2
        temp2 = Mid(Line, 2, 1)
        NewLine = Right(NewLine, Len(NewLine) - 1)
        Select Case temp2
        Case "1"
            Set Coll = Rec1_TicketableSegment(NewLine)
            InsertData Coll, "WSPTKTSEG", temp, UploadNo
            'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
            fOperatingITNRYSEGNO_Table_String = "WSPTKTSEG"
            'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
        Case "2"
            Set Coll = Rec1_NonTicketableSegment(NewLine)
            InsertData Coll, "WSPNTKTSEG", temp, UploadNo
            'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
            fOperatingITNRYSEGNO_Table_String = "WSPNTKTSEG"
            'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
        Case "3"
'            Set Coll = Rec33_TARUNKSegment(NewLine)
'            InsertData Coll, "WSPARUNK", temp, UploadNo
        Case "4"
            Set Coll = Rec4_4_AdditionalPTCs(NewLine)
            InsertData Coll, "WSPADDPTC", temp, UploadNo
        'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
        Case "6"
            Call Rec1_TicketableSegment_NonTicketableSegment_Operating(NewLine, UploadNo)
        'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
        End Select
    '------------------------------------------------------------
    Case "A"
        If temp2nd = "A\" Then
            Set Coll = RecA_NameDocumentNumbers(NewLine)
            InsertData Coll, "WSPPNAME", "A", UploadNo
        End If
    Case "D"
        Set Coll = RecD_OriginalIssueData(NewLine)
        InsertData Coll, "WSPORGISSDATA", "D", UploadNo
    Case "G"
        Set Coll = RecG_AirFair(NewLine)
        InsertData Coll, "WSPAIRFARE", "G", UploadNo
    Case "H"
        Set Coll = Rec1_HotelSegment(NewLine)
        InsertDataByFieldName Coll, "WSPHTLSEG", "H", UploadNo
    Case "E"
            Set Coll = Rec_Endorsement(NewLine)
            InsertDataByFieldName Coll, "WSPENDORSEMENT", "E", UploadNo
    Case "M"
            Set Coll = Rec_SSRData(NewLine)
            InsertDataByFieldName Coll, "WSPSSRDATA", "M", UploadNo
    Case "T"
            Set Coll = RecT_WSPTVLSEG(NewLine)
            InsertDataByFieldName Coll, "WSPTVLSEG", "T", UploadNo
    Case "U"
            Set Coll = Rec_UInput(NewLine)
            InsertDataByFieldName Coll, "WSPUINPUT", "U", UploadNo
    Case "N"
            Set Coll = Rec_Remarks(NewLine)
            InsertDataByFieldName Coll("GEN"), "WSPGENREMARKS", "N", UploadNo
            InsertDataCollectionKey Coll("SEL"), "WSPN_SELLRATE", "NS", UploadNo
            InsertDataCollectionKey Coll("VL"), "WSPN_VLOCATOR", "NS", UploadNo
            InsertDataByFieldName Coll("REF"), "WSPN_REF", "NS", UploadNo
            InsertDataByFieldName Coll("REM"), "WSPN_REMARKS", "NS", UploadNo
            Set CollPen = Coll("PEN")
            For vi = 1 To CollPen.Count
                InsertDataByFieldName CollPen(vi), "WSPPENLINE", "PEN", UploadNo
                DoEvents
            Next
            Set CollPen = Coll("PENPASSENGER")
            For vi = 1 To CollPen.Count
                InsertDataByFieldName CollPen(vi), "WSPPENLINE", "PEN", UploadNo
                DoEvents
            Next
            'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
            Set CollPen = Coll("PENFARE")
            For vi = 1 To CollPen.Count
                InsertDataByFieldName CollPen(vi), "WSPPENLINE", "PENFARE", UploadNo
                DoEvents
            Next
            'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
            Set CollPen = Coll("PENLINK")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENLINK", UploadNo
            End If
            'by Abhi on 10-Aug-2010 for caseid 1433 Penline PENAUTOOFF
            Set CollPen = Coll("PENAUTOOFF")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENAUTOOFF", UploadNo
            End If
            'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
            Set CollPen = Coll("PENO")
            For vi = 1 To CollPen.Count
                InsertDataByFieldName CollPen(vi), "WSPPENLINE", "PENO", UploadNo
                DoEvents
            Next
            'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
            Set CollPen = Coll("PENATOL")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENATOL", UploadNo
            End If
            'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
            Set CollPen = Coll("PENRT")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENRT", UploadNo
            End If
            Set CollPen = Coll("PENPOL")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENPOL", UploadNo
            End If
            Set CollPen = Coll("PENPROJ")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENPROJ", UploadNo
            End If
            Set CollPen = Coll("PENCC")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENCC", UploadNo
            End If
            Set CollPen = Coll("PENEID")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENEID", UploadNo
            End If
            Set CollPen = Coll("PENPO")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENPO", UploadNo
            End If
            Set CollPen = Coll("PENHFRC")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENHFRC", UploadNo
            End If
            Set CollPen = Coll("PENLFRC")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENLFRC", UploadNo
            End If
            Set CollPen = Coll("PENHIGHF")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENHIGHF", UploadNo
            End If
            Set CollPen = Coll("PENLOWF")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENLOWF", UploadNo
            End If
            Set CollPen = Coll("PENUC1")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENUC1", UploadNo
            End If
            Set CollPen = Coll("PENUC2")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENUC2", UploadNo
            End If
            Set CollPen = Coll("PENUC3")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENUC3", UploadNo
            End If
            Set CollPen = Coll("PENBB")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENBB", UploadNo
            End If
            Set CollPen = Coll("PENAGROSS")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENAGROSS", UploadNo
            End If
            'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
            Set CollPen = Coll("PENAIRTKT")
            For vi = 1 To CollPen.Count
                InsertDataByFieldName CollPen(vi), "WSPPENLINE", "PENAIRTKT", UploadNo
                DoEvents
            Next
            'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
            'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
            Set CollPen = Coll("PENBILLCUR")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENBILLCUR", UploadNo
            End If
            'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
            'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
            Set CollPen = Coll("PENWAIT")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENWAIT", UploadNo
            End If
            'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
            'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
            Set CollPen = Coll("PENVC")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENVC", UploadNo
            End If
            'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
            'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
            Set CollPen = Coll("PENCS")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENCS", UploadNo
            End If
            'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
            'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
            Set CollPen = Coll("PENRC")
            If CollPen.Count > 0 Then
                InsertDataByFieldName CollPen(1), "WSPPENLINE", "PENRC", UploadNo
            End If
            'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Case "X"
        Set Coll = RecX_EXCHVALUE(NewLine)
        InsertDataByFieldName Coll, "WSPEXCHVALUE", "X", UploadNo
    'by Abhi on 23-Jun-2017 for caseid 4527 OB ticketing fee mapping for Worldspan
    Case "Y"
        Set Coll = RecY_OBFees(NewLine)
        InsertDataByFieldName Coll, "WSPOBFees", "Y", UploadNo
    'by Abhi on 23-Jun-2017 for caseid 4527 OB ticketing fee mapping for Worldspan
    Case "Q"
        If FWorldSpan.RECIDQHeader = False Then
            FWorldSpan.RECIDQHeader = True
        Else
            Set Coll = RecQ_DELIVERYADD(NewLine)
            InsertDataByFieldName Coll, "WSPDELIVERYADD", "Q", UploadNo
        End If
    Case "J"
        Set Coll = RecJ_REFUND(NewLine)
        InsertDataByFieldName Coll, "WSPREFUND", "J", UploadNo
    Case "V"
        Set Coll = RecV_TRIPVALUE(NewLine)
        InsertDataByFieldName Coll, "WSPTRIPVALUE", "V", UploadNo
    'by Abhi on 31-Aug-2017 for caseid 7742 Worldspan EMD Ticktets from Record -E – For an Issued EMD
    Case "-"
        Select Case temp2nd
            Case "-E"
                Set Coll = RecE_IssuedEMD(NewLine)
                InsertDataByFieldName Coll, "WSPIssuedEMD", "-E", UploadNo
        End Select
    'by Abhi on 31-Aug-2017 for caseid 7742 Worldspan EMD Ticktets from Record -E – For an Issued EMD
    'Case Else
    '    If UCase(Left(Line, 4)) = UCase("PEN\") Then
    '        Set Coll = RecPEN_PENLINE(Line)
    '        InsertDataByFieldName Coll, "WSPPENLINE", "PEN", UploadNo
    '    End If
End Select

End Function


Private Function Rec1_TicketableSegment(Data) As Collection
Dim Splited
Dim Coll As New Collection
Dim SubSplit, SubSplit2
Splited = SplitForce(Data, "\", 17)

Coll.Add "1", "AIRSEGCODE"
Coll.Add Splited(1), "ITNRYSEGNO"
'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
fOperatingITNRYSEGNO_Long = Val(Splited(1))
'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
'by Abhi on 10-Nov-2009 for VLOCATOR in AirSegDetails
If Val(Splited(1)) > mITNRYSEGNO_Long Then
    mITNRYSEGNO_Long = Val(Splited(1))
End If
SubSplit = SplitWithLengths(Splited(2), 1, 1, 2, 4)
    Coll.Add SubSplit(0), "NOINTSTOP"
    Coll.Add SubSplit(1), "SHDESIGNINDI"
    Coll.Add SubSplit(2), "AIRCODE"
    Coll.Add SubSplit(3), "FLTNO"
Coll.Add Splited(3), "CLASS"
SubSplit = SplitWithLengths(Splited(4), 3, 5, 5)
    Coll.Add SubSplit(0), "ORGAIRCODE"
    Coll.Add SubSplit(1), "DEPDATE"
    Coll.Add SubSplit(2), "DEPTIME"
SubSplit = SplitWithLengths(Splited(5), 3, 5, 5)
    Coll.Add SubSplit(0), "DESTAIRCODE"
    Coll.Add SubSplit(1), "ARRDATE"
    Coll.Add SubSplit(2), "ARRTIME"
Coll.Add Splited(6), "BDATE"
Coll.Add Splited(7), "ADATE"
Coll.Add Splited(8), "BAGG"
SubSplit = SplitWithLengths(Splited(9), 2, 1, 1) 'Some ModificationNeeded
    Coll.Add SubSplit(0), "STATUS"
    Coll.Add ToMealsServiceCode(SubSplit(1)), "MEALSERVCODE" 'Mealsdetails
    Coll.Add SubSplit(2), "SEGSTOPCODE"
Coll.Add Splited(10), "EQUIPTYPE"
Coll.Add Splited(11), "FBASISCODE"
Coll.Add Splited(12), "SEGMILEAGE"
Coll.Add Splited(13), "INTSTOP"
Coll.Add Splited(14), "AFLTTIME"
Coll.Add Splited(15), "SEGOVRRDIND"
Coll.Add Splited(16), "DEPTRMNLCODE"
Coll.Add Splited(17), "ARRTRMNLCODE"
SubSplit = SplitWithLengthsPlus(Splited(9), 3)
Coll.Add SubSplit(0), "BPAYAMT"
'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
Coll.Add "", "OperatingAIRNAME"
Coll.Add "", "OperatingAIRID"
Coll.Add "", "OperatingFlightNo"
'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
    
Set Rec1_TicketableSegment = Coll
End Function

'Non-Ticketable Segment
Private Function Rec1_NonTicketableSegment(Data) As Collection

Dim Splited
Dim Coll As New Collection
Dim SubSplit
Splited = Split(Data, "\")

Coll.Add "2", "AIRSEGCODE"

Coll.Add Splited(1), "ITNRYSEGNO"
'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
fOperatingITNRYSEGNO_Long = Val(Splited(1))
'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
'by Abhi on 10-Nov-2009 for VLOCATOR in AirSegDetails
If Val(Splited(1)) > mITNRYSEGNO_Long Then
    mITNRYSEGNO_Long = Val(Splited(1))
End If
SubSplit = SplitWithLengths(Splited(2), 1, 1, 2, 4)
    Coll.Add SubSplit(0), "NOINTERSTOP"
    Coll.Add SubSplit(1), "SHDESIIND"
    Coll.Add SubSplit(2), "AIRCODE"
    Coll.Add SubSplit(3), "FLTNO"
Coll.Add Splited(3), "Class of Service"
SubSplit = SplitWithLengths(Splited(4), 3, 5, 5)
    Coll.Add SubSplit(0), "ORGAIRCODE"
    Coll.Add SubSplit(1), "DEPDATE"
    Coll.Add SubSplit(2), "DEPTIME"
SubSplit = SplitWithLengths(Splited(5), 3, 5, 5)
    Coll.Add SubSplit(0), "DESTAIRCODE"
    Coll.Add SubSplit(1), "ARRDATE"
    Coll.Add SubSplit(2), "ARRTIME"

Coll.Add Splited(6), "AS1"
Coll.Add Splited(7), "AS2"
Coll.Add Splited(8), "AS3"

SubSplit = SplitWithLengths(Splited(9), 2, 42, 1)
    Coll.Add SubSplit(0), "STATUS"
    Coll.Add SubSplit(1), "MEALSERVCODE"
    Coll.Add SubSplit(2), "SEGSTOPCODE"


Coll.Add Splited(10), "Equipment type"
Coll.Add Splited(11), "AS4*"
Coll.Add Splited(12), "Segment Mileage"
Coll.Add Splited(13), "Intermediate Stops"
Coll.Add Splited(14), "Accumulated Elapsed Flight Time"
Coll.Add Splited(15), "Segment Override Indicator"
Coll.Add Splited(16), "Departure airport terminal code"
Coll.Add Splited(17), "Arrival airport terminal code"
'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
Coll.Add "", "OperatingAIRNAME"
Coll.Add "", "OperatingAIRID"
Coll.Add "", "OperatingFlightNo"
'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan

Set Rec1_NonTicketableSegment = Coll
End Function


Private Function Rec1_PNRFileAddress(Data, Optional FNAME = "")
Dim Splited
Dim SubSplited
Dim Temp1, temp2
Dim Coll As New Collection

Splited = SplitForcePlus(Data, "\", 3)
Coll.Add Splited(1), "INTLVL"
Temp1 = Splited(2)

temp2 = ExtractBetween(Temp1, "FA-", "\")
    Coll.Add temp2, "PNRADD"
    'by Abhi on 14-Nov-2010 for caseid 1551 PenGDS last uploaded pnr and date time monitoring
    LUFPNR_String = temp2
temp2 = ExtractBetween(Temp1, "IN-", "\")
    SubSplited = SplitForce(temp2, "-", 2)
        Coll.Add SubSplited(0), "FINVNO"
        Coll.Add SubSplited(1), "LINVNO"
temp2 = ExtractBetween(Temp1, "NC-", "\")
    Coll.Add temp2, "ITNRYCHNGE"
temp2 = ExtractBetween(Temp1, "LC-", "\")
    SubSplited = SplitForce(temp2, "/", 2)
        Coll.Add SubSplited(0), "DLCI"
        Coll.Add SubSplited(1), "TLCI"
        
        Coll.Add FNAME, "FNAME"
        Coll.Add Now, "LUPDATE"
        Coll.Add 0, "GDSAutoFailed"

Set Rec1_PNRFileAddress = Coll
End Function

Private Function Rec8_AirCommission(Data)
'Some Problems with MCP- & CP-

Dim Splited
Dim SubSplited, SubSplitedPreserve
Dim Temp1, temp2
Dim Coll As New Collection

Temp1 = Data

temp2 = ExtractBetween(Temp1, "MCP-", "\")
    Coll.Add temp2, "MNYCOLAMT"


temp2 = ExtractBetween(Temp1, "OCA-", "\")
    Coll.Add temp2, "OCOMAMT"
temp2 = ExtractBetween(Temp1, "ACA-", "\")
    Coll.Add temp2, "ADMCOMAMT"
temp2 = ExtractBetween(Temp1, "NCA-", "\")
    Coll.Add temp2, "NCOMAMT"
temp2 = ExtractBetween(Temp1, "CP-", "\")
    Coll.Add temp2, "COMPERC"
temp2 = ExtractBetween(Temp1, "CA-", "\")
    SubSplited = SplitWithLengthsPlus(temp2, 1)
        Coll.Add SubSplited(0), "PNORBANK"
        Coll.Add SubSplited(1), "PCOM"
        Coll.Add PassengerToAirFairID, "PassengerID"
        
        
'--------Begin Details

Splited = SplitForcePlus(Temp1, "\", 3)
Coll.Add Splited(1), "PTYPECODE"
Coll.Add Splited(2), "EXCCOMID"
'-----------------
Set Rec8_AirCommission = Coll
End Function


Private Function Rec9_TicketingCarrier(Data)
Dim Splited
Dim SubSplited
Dim Temp1, temp2
Dim Coll As New Collection

Temp1 = Data
temp2 = ExtractBetween(Temp1, "V-", "\")
    SubSplited = SplitForce(temp2, "/", 2)
        Coll.Add SubSplited(0), "VAIRCODE"
        Coll.Add SubSplited(1), "VAIRNO"
temp2 = ExtractBetween(Temp1, "S-", "\")
    Coll.Add temp2, "TKTIND"
temp2 = ExtractBetween(Temp1, "I-", "\")
    Coll.Add temp2, "INTIND"
temp2 = ExtractBetween(Temp1, "DP-", "\")
    SubSplited = SplitForce(temp2, "/", 2)
        Coll.Add SubSplited(0), "DESTCDE"
        Coll.Add SubSplited(1), "PTCODE"
        
Set Rec9_TicketingCarrier = Coll
End Function


Private Function Rec33_TARUNKSegment(Data)
Dim Splited
Dim SubSplited
Dim Coll As New Collection
Splited = SplitForce(Data, "\", 3)

Coll.Add "3", "ARUNK segment code"
Coll.Add Splited(1), "Itinerary segment number"
Coll.Add Splited(2), "Arrival Unknown"

Set Rec33_TARUNKSegment = Coll
End Function

Private Function RecA_NameDocumentNumbers(Data)
Dim Splited
Dim SubSplited, SubSplited2
Dim temp
Dim Temp1, temp2
Dim Coll As New Collection
Temp1 = Data
'by Abhi on 13-Jul-2015 for caseid 5398 worldspan reissue tickets
Dim vSAC_String As String, vSAC2_String As String, vSAC3_String As String
'by Abhi on 13-Jul-2015 for caseid 5398 worldspan reissue tickets

temp2 = ExtractBetween(Temp1, "E-", "\")
temp2 = ExtractBetween(Temp1, "D-", "\")
'by Abhi on 13-Jul-2015 for caseid 5398 worldspan reissue tickets
'temp2 = ExtractBetween(Temp1, "SAC-", "\")
vSAC_String = ExtractBetween(Temp1, "SAC-", "\")
vSAC2_String = ExtractBetween(Temp1, "SAC2-", "\")
vSAC3_String = ExtractBetween(Temp1, "SAC3-", "\")
'by Abhi on 13-Jul-2015 for caseid 5398 worldspan reissue tickets

Splited = SplitForcePlus(Temp1, "\", 10)

SubSplited = SplitForce(Splited(2), "/", 2)
    Coll.Add SubSplited(0), "SURNAME"
    SubSplited2 = SplitFirstNameAndInitial(SubSplited(1))
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'Coll.Add Replace(SubSplited2(0), "*CHD", ""), "FRSTNAME" 'Some Records Found unwanted '*CHD' which is skipped
    'Coll.Add Replace(SubSplited2(1), "*CHD", ""), "PTitle"
    SubSplited2(0) = Replace(SubSplited2(0), "*CHD", "")
    SubSplited2(1) = Replace(SubSplited2(1), "*CHD", "")
    SubSplited2(0) = Replace(SubSplited2(0), "*YTH", "")
    SubSplited2(1) = Replace(SubSplited2(1), "*YTH", "")
    Coll.Add SubSplited2(0), "FRSTNAME"
    Coll.Add SubSplited2(1), "PTitle"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Apr-2010 for caseid 1293 PENFARE modification for Worldspan and Amadeus
    'If UCase(splited(3)) = "ADT" Then
    '    temp = "Adult"
    'ElseIf UCase(splited(3)) = "CNN" Then
    '    temp = "Child"
    'Else
    '    temp = ""
    'End If
    'by Abhi on 28-Oct-2010 for caseid 1531 PenGDS should take *CHD in passenger name as Child
    'temp = GetPassengerTypefromShort(splited(3), True, splited(2))
    temp = GetPassengerTypefromShort(Splited(3), True, Splited(2))
Coll.Add temp, "PTYPE"
Coll.Add Splited(4), "CUSTNO"
Coll.Add Splited(5), "CUSTCMNTS"
Coll.Add Splited(6), "DOCNO"
Coll.Add Splited(7), "ISSUEDATE"
Coll.Add Splited(8), "INVNO"
'by Abhi on 13-Jul-2015 for caseid 5398 worldspan reissue tickets
'SubSplited = SplitForce(splited(9), "D-\E-\SAC-", 2)
'Coll.Add SubSplited(1), "SETTMNTCDENO"
Coll.Add vSAC_String, "SETTMNTCDENO"
'by Abhi on 13-Jul-2015 for caseid 5398 worldspan reissue tickets

If PassengerToAirFairFlag = False Then
    PassengerToAirFairID = PassengerToAirFairID + 1
    PassengerToAirFairFlag = True
End If
Coll.Add PassengerToAirFairID, "PassengerID"
If InStr(1, Data, "E-") > 0 Then
    Coll.Add "E-", "ETKTIND"
Else
    Coll.Add "N", "ETKTIND"
End If
'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
'Coll.Add splited(1), "SURNAMEFRSTNAMENO"
Coll.Add Left(Splited(1), 50), "SURNAMEFRSTNAMENO"
'by Abhi on 13-Jul-2015 for caseid 5398 worldspan reissue tickets
Coll.Add vSAC2_String, "SETTMNTCDENO2"
Coll.Add vSAC3_String, "SETTMNTCDENO3"
'by Abhi on 13-Jul-2015 for caseid 5398 worldspan reissue tickets

Set RecA_NameDocumentNumbers = Coll
End Function

Private Function RecD_OriginalIssueData(Data)
Dim Splited
Dim SubSplited
Dim Coll As New Collection
Splited = SplitForcePlus(Data, "\", 9)

SubSplited = SplitWithLengthsPlus(Splited(0), 1)
    Coll.Add SubSplited(0), "Document type indicator"
    Coll.Add SubSplited(1), "Original issue document number"
Coll.Add Splited(1), "Place of issue "
Coll.Add Splited(2), "Original issue date"
Coll.Add Splited(3), "Original issue agency IATA"
Coll.Add Splited(4), "FREEDATA"
Set RecD_OriginalIssueData = Coll
End Function



Private Function RecG_AirFair(Data) As Collection
Dim Splited
Dim Coll As New Collection
Dim SubSplit
Splited = SplitForce(Data, "\", 7)

SubSplit = SplitWithLengthsPlus(Splited(0), 3)
    Coll.Add SubSplit(0), "Currency code of fare "
    Coll.Add SubSplit(1), "Fare amount"
SubSplit = SplitWithLengthsPlus(Splited(1), 3)
    Coll.Add SubSplit(0), "Tax 1 code identifier "
    Coll.Add SubSplit(1), "Tax 1 amount"

SubSplit = SplitWithLengthsPlus(Splited(2), 3)
    Coll.Add SubSplit(0), "Tax 2 code identifier "
    Coll.Add SubSplit(1), "Tax 2 amount"

SubSplit = SplitWithLengthsPlus(Splited(3), 3)
    Coll.Add SubSplit(0), "Tax 3 code identifier "
    Coll.Add SubSplit(1), "Tax 3 amount"

SubSplit = SplitWithLengthsPlus(Splited(4), 3)
    Coll.Add SubSplit(0), "Currency code of Total Fare"
    Coll.Add SubSplit(1), "Total Fare amount"

SubSplit = SplitWithLengthsPlus(Splited(5), 3)
    Coll.Add SubSplit(0), "Currency code of equivalent fare"
    Coll.Add SubSplit(1), "Equivalent fare amount "

SubSplit = SplitWithLengthsPlus(Splited(6), 2)
    Coll.Add SubSplit(1), "Invoice Air Amount"

If PassengerToAirFairFlag = True Then
    PassengerToAirFairFlag = False
End If
    Coll.Add PassengerToAirFairID, "PassengerID"
    
    WorldSpan_AirFareSLNO = WorldSpan_AirFareSLNO + 1
    Coll.Add WorldSpan_AirFareSLNO, "SLNO"

Set RecG_AirFair = Coll
End Function



Private Function Rec4_4_AdditionalPTCs(Data) As Collection
Dim Temp1
Dim Splited
Dim SubSplited
Dim Coll As New Collection
'by Abhi on 19-Oct-2009 for International GDS Subscript out of range
'splited = SplitForcePlus(Data, "\", 5)
Splited = SplitForcePlus(Data, "\", 6)
Coll.Add "4", "PTCCODE"
Coll.Add Splited(1), "FBASISCODE"
Coll.Add Splited(2), "BDATE"
Coll.Add Splited(3), "ADATE"
Coll.Add Splited(4), "STATUS"

Temp1 = Splited(5)
SubSplited = ExtractBetween(Temp1, "BG-", "\")
    Coll.Add SubSplited, "BAGGAGE"
SubSplited = ExtractBetween(Temp1, "E-", "\")
    Coll.Add SubSplited, "ENDRSMNT"
SubSplited = ExtractBetween(Temp1, "BP-", "\")
    Coll.Add SubSplited, "AMT"
    
Set Rec4_4_AdditionalPTCs = Coll
End Function


Private Function Rec1_HotelSegment(Data)
Dim Splited
Dim SubSplited, SubSplited2, SubSplited3
Dim Temp1, temp2
Dim Coll As New Collection
Temp1 = Data

'--------Starting Details are below---------




temp2 = ExtractBetween(Temp1, "NP-", "\")
    Coll.Add temp2, "NOPER"
temp2 = ExtractBetween(Temp1, "R-", "\")
    Coll.Add temp2, "TYPE"
    
temp2 = ExtractBetween(Temp1, "RS-", "\")
    SubSplited = SplitForce(temp2, "-", 2)
        SubSplited2 = SplitFirstTwo(SubSplited(0), 3)
            Coll.Add SubSplited2(0), "CURRCODE"
            Coll.Add SubSplited2(1), "SYSRAMT"
        Coll.Add SubSplited(1), "SYSRPLAN"
temp2 = ExtractBetween(Temp1, "RG-", "\")
    SubSplited = SplitForce(temp2, "-", 2)
        SubSplited2 = SplitTwo(SubSplited(0), 3)
            Coll.Add SubSplited2(0), "RGCURRCODE"
            Coll.Add SubSplited2(1), "RGAMT"
temp2 = ExtractBetween(Temp1, "ROE-", "\")
    SubSplited = SplitForce(temp2, "-", 2)
        SubSplited2 = SplitForce(SubSplited(0), "/", 3)
            Coll.Add CStr(ToPenDate(SubSplited2(0), Date)), "EXRATEDTE"
            Coll.Add SubSplited2(1), "CURRCODELOC"
            SubSplited3 = SplitTwo(SubSplited2(2), 2)
                Coll.Add SubSplited3(0), "CNTRYCODEH"
                Coll.Add SubSplited3(1), "EXCHRATE"
temp2 = ExtractBetween(Temp1, "RD-", "\")
    Coll.Add temp2, "ROOMDESC"
temp2 = ExtractBetween(Temp1, "RTD-", "\")
    Coll.Add temp2, "RATEDESC"
temp2 = ExtractBetween(Temp1, "RL-", "\")
    Coll.Add temp2, "ROOMLOC"
temp2 = ExtractBetween(Temp1, "BS-", "\")
    Coll.Add temp2, "BSOURCE"
temp2 = ExtractBetween(Temp1, "NM-", "\")
    Coll.Add temp2, "GUESTNAME"
temp2 = ExtractBetween(Temp1, "CD-", "\")
    Coll.Add temp2, "DISCOUNTNO"
temp2 = ExtractBetween(Temp1, "FT-", "\")
    Coll.Add temp2, "FTNO"
temp2 = ExtractBetween(Temp1, "FG-", "\")
    Coll.Add temp2, "FGUESTNO"
temp2 = ExtractBetween(Temp1, "TTL-", "\")
    Coll.Add temp2, "TOTALAMT"
temp2 = ExtractBetween(Temp1, "BAS-", "\")
    Coll.Add temp2, "BASEAMT"
temp2 = ExtractBetween(Temp1, "SVC-", "\")
    Coll.Add temp2, "SCHRGEAMT"
    
temp2 = ExtractBetween(Temp1, "SUR-", "\")
    Coll.Add temp2, "SURCHARGE"
temp2 = ExtractBetween(Temp1, "TTD-", "\")
    Coll.Add temp2, "TTD"
temp2 = ExtractBetween(Temp1, "CM-", "\")
    Coll.Add temp2, "COMAMT"
temp2 = ExtractBetween(Temp1, "CV-", "\")
    Coll.Add temp2, "VCOM"
temp2 = ExtractBetween(Temp1, "CF-", "\")
    Coll.Add temp2, "CONFNO"

temp2 = ExtractBetween(Temp1, "CX-", "\")
    Coll.Add temp2, "CNCLNNO"
temp2 = ExtractBetween(Temp1, "TX-", "\")
    Coll.Add temp2, "TAXRATE"
temp2 = ExtractBetween(Temp1, "HA1-", "\")
    Coll.Add temp2, "HOTADD1"
temp2 = ExtractBetween(Temp1, "HA2-", "\")
    Coll.Add temp2, "HOTADD2"
temp2 = ExtractBetween(Temp1, "SCC-", "\")
    Coll.Add temp2, "CNTRYCODE"

temp2 = ExtractBetween(Temp1, "ZIP-", "\")
    Coll.Add temp2, "POSTELCODE"
temp2 = ExtractBetween(Temp1, "PH-", "\")
    Coll.Add temp2, "TELENO"
temp2 = ExtractBetween(Temp1, "FAX-", "\")
    Coll.Add temp2, "FAXNO"
temp2 = ExtractBetween(Temp1, "CI-", "\")
    Coll.Add temp2, "CHKINTIME"
temp2 = ExtractBetween(Temp1, "CO-", "\")
    Coll.Add temp2, "CHKOUTTIME"


'------Starting Details
Splited = SplitForcePlus(Temp1, "\", 8)
Coll.Add Splited(1), "SEGNO"
Coll.Add Splited(2), "CHAINCODE"
Coll.Add Splited(3), "CHAINNAME"
SubSplited = SplitTwo(Splited(4), 2) 'SplitFirstTwo(splited(4), 2)
    Coll.Add SubSplited(0), "STATUSCODE"
    Coll.Add SubSplited(1), "ROOMS"
SubSplited = SplitWithLengths(Splited(5), 3, 5, 5)
    Coll.Add SubSplited(0), "CTYCODE"
    Coll.Add SubSplited(1), "INDATE"
    Coll.Add SubSplited(2), "OUTDATE"
Coll.Add Splited(6), "PROPCODE"
Coll.Add Splited(7), "PROPNAME"

Coll.Add "RT-", "RT"
temp2 = ExtractBetween(Temp1, "RT-", "\")
    SubSplited = SplitWithLengths(temp2, 3, 10)
    Coll.Add SubSplited(0), "RTCURRCODE"
    Coll.Add SubSplited(1), "RTAMT"

Coll.Add "RQ-", "RQ"
temp2 = ExtractBetween(Temp1, "RQ-", "\")
    SubSplited = SplitWithLengths(temp2, 3, 10)
    Coll.Add SubSplited(0), "RQCURRCODE"
    Coll.Add SubSplited(1), "RQAMT"

Coll.Add "RQ1-", "RQ1"
temp2 = ExtractBetween(Temp1, "RQ1-", "\")
    SubSplited = SplitWithLengths(temp2, 3, 10)
    Coll.Add SubSplited(0), "RQ1CURRCODE"
    Coll.Add SubSplited(1), "RQ1AMT"

Coll.Add "RG1-", "RG1"
temp2 = ExtractBetween(Temp1, "RG1-", "\")
    SubSplited = SplitWithLengths(temp2, 3, 10)
    Coll.Add SubSplited(0), "RG1CURRCODE"
    Coll.Add SubSplited(1), "RG1AMT"

Coll.Add "RT1", "RT1"
temp2 = ExtractBetween(Temp1, "RT1", "\")
    SubSplited = SplitWithLengths(temp2, 3, 10)
    Coll.Add SubSplited(0), "RT1CURRCODE"
    Coll.Add SubSplited(1), "RT1AMT"

'----------------------

Set Rec1_HotelSegment = Coll

End Function

Private Function Rec_BranchAgentSines(Data)
Dim Splited
Dim SubSplited, SubSplited2
Dim Temp1, temp2
Dim Coll As New Collection
Temp1 = Data

temp2 = ExtractBetween(Temp1, "BL-", "\")
    SubSplited = SplitForce(temp2, "/", 2)
        SubSplited2 = SplitWithLengths(SubSplited(0), 3, 8)
            Coll.Add SubSplited2(0), "BSID"
            Coll.Add SubSplited2(1), "BIATA"
        SubSplited2 = SplitWithLengths(SubSplited(1), 7, 4, 2)
            Coll.Add SubSplited2(0), "BDATE"
            Coll.Add SubSplited2(1), "BTIME"
            Coll.Add SubSplited2(2), "BAGENT"
            
            
temp2 = ExtractBetween(Temp1, "TL-", "\")
    SubSplited = SplitForce(temp2, "/", 2)
        SubSplited2 = SplitWithLengths(SubSplited(0), 3, 8)
            Coll.Add SubSplited2(1), "TIATA"
        SubSplited2 = SplitWithLengths(SubSplited(1), 7, 4, 2)
            Coll.Add CStr(ToPenDateX(SubSplited2(0))), "TDATE"
            Coll.Add (SubSplited2(2)), "TAGENT"
            
Set Rec_BranchAgentSines = Coll
End Function

Private Sub ClearTable(TableName As String)
dbCompany.Execute "Delete  from " & TableName
End Sub

Public Function ClearAll()
    ClearTable "WSPAIRCOM"
    ClearTable "WSPAIRFARE"
    ClearTable "WSPARUNK"
    ClearTable "WSPTKTSEG"
    ClearTable "WSPNTKTSEG"
    ClearTable "WSPORGISSDATA"
    ClearTable "WSPPNAME"
    ClearTable "WSPPNRADD"
    ClearTable "WSPTCARRIER"
    ClearTable "WSPADDPTC"
    ClearTable "WSPCLNTACNO"
    
End Function

Private Function To24Hour(time)
    Dim MainTime, AMPM, HourT, MinuteT, retn
    If Len(time) = 0 Then Exit Function
    AMPM = Right(time, 1)
    If Trim(AMPM) = "N" Then
        AMPM = "P"
    End If
    MainTime = Left(time, Len(time) - 1)
    MinuteT = Right(MainTime, 2)
    HourT = Left(MainTime, Len(MainTime) - 2)
    'If UCase(AMPM) = "P" Then
    'HourT = Val(HourT) + 12
    'End If
    'retn = HourT & MinuteT
    retn = Format(HourT & ":" & MinuteT & " " & UCase(AMPM), "HHMM")
    To24Hour = retn
End Function

Private Function ToPenDate(ByVal TheDate As String, Optional default = Empty) As Date
On Error GoTo errPara
Dim Day, Month, Year, TESTDATE
Day = Left(TheDate, 2)
Month = Right(TheDate, 3)
Year = VBA.Year(Date)
TESTDATE = CDate(Day & "/" & Month & "/" & Year)
If (TESTDATE < CDate(Format(Date, "dd-MMM-yyyy"))) Then
    Year = Year + 1
    TESTDATE = CDate(Day & "/" & Month & "/" & Year)
End If
ToPenDate = TESTDATE
Exit Function
errPara:
ToPenDate = default
End Function

Private Function ToPenDateX(TheDate, Optional default = Empty) As Date
On Error GoTo errPara
Dim Day, Month, Year, TESTDATE
Day = Left(TheDate, 2)
Month = Mid(TheDate, 3, 3)
Year = Right(TheDate, 2)
TESTDATE = CDate(Day & "/" & Month & "/" & Year)
ToPenDateX = TESTDATE
Exit Function
errPara:
    ToPenDateX = default
End Function

'Public Function SplitTwoReversePlus(Data, delim)
'Dim aa, ub, ab, ac, AD(1)
'Dim temp
'aa = Split(Data, ".")
'ub = UBound(aa)
'ab = aa(0)
'temp = Len(Data) - (Len(ab) + 1)
'If temp > 0 Then
'    ac = Right(Data, temp)
'Else
'    ac = ""
'End If
'AD(0) = ab
'AD(1) = ac
'SplitTwoReversePlus = AD
'End Function
Private Function SplitFirstNameAndInitial(Data)
Dim aa, ub, ab, ac, AD(1)
Dim temp

'by Abhi on 28-Oct-2010 for caseid 1531 PenGDS should take *CHD in passenger name as Child
'aa = Split(Data, ".")
'ub = UBound(aa)
'If ub > 0 Then
'    ab = aa(ub)
'    temp = Len(Data) - (Len(ab) + 1)
'    If temp > 0 Then
'        ac = Left(Data, temp)
'    Else
'        ac = ""
'    End If
'Else
'temp = FindInitialAndName(aa(0))
temp = FindInitialAndName(Data)
    ac = temp(0)
    ab = temp(1)
'by Abhi on 28-Oct-2010 for caseid 1531 PenGDS should take *CHD in passenger name as Child
'End If
AD(0) = ac
AD(1) = ab
SplitFirstNameAndInitial = AD
End Function


Private Function ToMealsServiceCode(id)
Dim ret
'by Abhi on 07-Aug-2015 for caseid 5479 Additional Meal Service Codes from Worldspan
'Select Case Trim(id)
'Case "B"
'    ret = "Breakfast"
'Case "D"
'    ret = "Dinner"
'Case "L"
'    ret = "Lunch"
'Case "R"
'    ret = "Brunch"
'Case "S"
'    ret = "Snack"
'Case "*"
'    ret = "Miscellaneous service"
'Case ""
'    ret = "No meal service"
'End Select
Select Case Trim(id)
Case "B"
    ret = "BREAKFAST"
Case "D"
    ret = "DINNER"
Case "L"
    ret = "LUNCH"
Case "R"
    ret = "REFRESHMENTS"
Case "S"
    ret = "SNACK OR BRUNCH"
Case "*"
    ret = "Miscellaneous service"
Case ""
    'by Abhi on 04-Nov-2016 for caseid 6712 Meal service code picking for Worldspan
    'ret = "No meal service"
    ret = ""
    'by Abhi on 04-Nov-2016 for caseid 6712 Meal service code picking for Worldspan
Case "C"
    ret = "COMPLIMENTARY LIQUOR"
Case "H"
    ret = "HOT MEAL"
Case "M"
    ret = "MEAL"
Case "K"
    ret = "CONTINENTAL BREAKFAST"
Case "P"
    ret = "LIQUOR FOR PURCHASE"
Case "F"
    ret = "FOOD FOR PURCHASE"
Case "O"
    ret = "COLD MEAL"
Case "N"
    ret = "NO MEAL SERVICE"
Case "G"
    ret = "FOOD AND BEVERAGE FOR PURCHASE"
Case "V"
    ret = "REFRESHMENTS FOR PURCHASE"
End Select
'by Abhi on 07-Aug-2015 for caseid 5479 Additional Meal Service Codes from Worldspan
ToMealsServiceCode = ret
End Function



Private Function Rec_ClientAcNo(Data)
Dim Splited
Dim SubSplited
Dim Temp1, temp2
Dim Coll As New Collection
Temp1 = Data
Coll.Add Data, "CLNTACNO"
'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
If Trim(Data) <> "" Then
    GIT_CUSTOMERUSERCODE_String = Data
End If
'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
Set Rec_ClientAcNo = Coll
End Function

Private Function Rec_Endorsement(Data)
Dim Splited
Dim SubSplited
Dim Temp1, temp2
Dim Coll As New Collection
Temp1 = Data
Coll.Add Data, "INFORMATION"
Set Rec_Endorsement = Coll
End Function

Private Function Rec_SSRData(Data)
Dim Splited
Dim SubSplited
Dim Temp1, temp2
Dim Coll As New Collection
Temp1 = Data
SubSplited = SplitForce(Temp1, "\", 4)

Coll.Add SubSplited(0), "SSR1"
Coll.Add SubSplited(1), "SSR2"
Coll.Add SubSplited(2), "SSR3"
Coll.Add SubSplited(3), "SSR4"

Set Rec_SSRData = Coll
End Function


Private Function Rec_PhoneContact(Data)
Dim Splited
Dim SubSplited
Dim Temp1, temp2
Dim Coll As New Collection
Temp1 = Data
Coll.Add Data, "PHONE"
Set Rec_PhoneContact = Coll
End Function

Private Function Rec_UInput(Data)
Dim Splited
Dim SubSplited
Dim Temp1, temp2
Dim Coll As New Collection
Temp1 = Data
Temp1 = SplitWithLengthsPlus(Data, 1)
Coll.Add Temp1(1), "UINPUT"
Set Rec_UInput = Coll
End Function



Private Function Rec_FormOfPayment(Data) As Collection
Dim Temp1, temp2, temp3
Dim Splited
Dim SubSplited
Dim Coll As New Collection
    
Temp1 = Data
SubSplited = ExtractBetween(Temp1, "F1-", "\")
    Coll.Add SubSplited, "FFDATA1"
SubSplited = ExtractBetween(Temp1, "F2-", "\")
    Coll.Add SubSplited, "FFDATA2"
SubSplited = ExtractBetween(Temp1, "F3-", "\")
    Coll.Add SubSplited, "FFDATA3"
SubSplited = ExtractBetween(Temp1, "F4-", "\")
    Coll.Add SubSplited, "FFDATA4"
    
SubSplited = ExtractBetween(Temp1, "ES-", "\")
    Coll.Add SubSplited, "STAXCODE"
SubSplited = ExtractBetween(Temp1, "IT1-", "\")
    Coll.Add SubSplited, "ITAXCODE1"
SubSplited = ExtractBetween(Temp1, "IT2-", "\")
    Coll.Add SubSplited, "ITAXCODE2"
    
Splited = SplitForcePlus(Temp1, "\", 3)
temp2 = SplitWithLengthsPlus(Splited(1), 2)
temp3 = FormOfPayNotes(temp2(0))
    Coll.Add temp2(0), "PAYCODE"
    Coll.Add temp3, "PAYNAME"
    
    Coll.Add temp2(1), "CCDETAILS"
    
    
Set Rec_FormOfPayment = Coll
End Function


Private Function FormOfPayNotes(id) As String
Dim retn
Select Case id
    Case "AR"
        retn = "Accounts Receivable"
    Case "AN"
        retn = "Agent Non-refundable"
    Case "CA"
        retn = "Cash"
    Case "CK"
        retn = "Check"
    Case "GR"
        retn = "Gtr"
    Case "CC"
        retn = "Credit Card"
    Case "MS"
        retn = "Miscellaneous Cash"
    Case "MSC"
        retn = "Miscellaneous Charge"
End Select
FormOfPayNotes = retn
End Function



Private Function Rec_Remarks(Data) As Collection
Dim Temp1, temp
Dim Splited
Dim SubSplited
Dim Coll As New Collection


Dim CollGEN As New Collection
Dim CollSellRate As New Collection
Dim CollVL As New Collection
Dim CollRef As New Collection
Dim CollRemarks As New Collection
Dim CollPenline As New Collection
Dim CollPenSplited As New Collection
Dim CollPenPass As New Collection
'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
Dim CollPenPENFARE As New Collection

Dim TempPEN, TempPENSplited
Dim TempPENi As Long, TempPENCount As Long
Dim TempPENPNO As Long
Dim TempPENP1E1 As String
'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
Dim CollPenPENLINK As New Collection
'by Abhi on 10-Aug-2010 for caseid 1433 Penline PENAUTOOFF
Dim CollPenPENAUTOOFF As New Collection
'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
Dim CollPenPENO As New Collection
Dim vPENONO_Long As Long
'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
Dim CollPenPENATOL As New Collection
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Dim CollPenPENRT As New Collection
Dim CollPenPENPOL As New Collection
Dim CollPenPENPROJ As New Collection
Dim CollPenPENCC As New Collection
Dim CollPenPENEID As New Collection
Dim CollPenPENPO As New Collection
Dim CollPenPENHFRC As New Collection
Dim CollPenPENLFRC As New Collection
Dim CollPenPENHIGHF As New Collection
Dim CollPenPENLOWF As New Collection
Dim CollPenPENUC1 As New Collection
Dim CollPenPENUC2 As New Collection
Dim CollPenPENUC3 As New Collection
Dim CollPenPENBB As New Collection
'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
Dim CollPenPENAGROSS As New Collection
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
Dim CollPenPENAIRTKT As New Collection
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
Dim CollPenPENBILLCUR As New Collection
'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
Dim CollPenPENWAIT As New Collection
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
Dim CollPenPENVC As New Collection
'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
Dim CollPenPENCS As New Collection
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
Dim CollPenPENRC As New Collection
'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

Temp1 = Data
TempPEN = Data
TempPENP1E1 = Data

'TempPENCount = StringOccurs(TempPEN, "PEN/", False)
'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
'For TempPENPNO = 1 To 99
'    'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
'    'If InStr(1, TempPEN, "P" & TempPENPNO & "E-", vbTextCompare) > 0 Or InStr(1, TempPEN, "P" & TempPENPNO & "T-", vbTextCompare) > 0 Then
'    If InStr(1, TempPEN, PENLINEID_String & "E" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, TempPEN, PENLINEID_String & "T" & TempPENPNO & "-", vbTextCompare) > 0 Then
'        Set CollPenSplited = RecPEN_PENLINEPassenger(TempPEN, TempPENPNO)
'        CollPenPass.Add CollPenSplited, str(TempPENPNO)
'    End If
'    DoEvents
'Next

'TempPENSplited = Split(Data, "\")
TempPENSplited = Split(TempPEN, "\")
TempPENCount = UBound(TempPENSplited)
'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
vPENONO_Long = 1
For TempPENi = 0 To TempPENCount
    'If UCase(Left(TempPENSplited(TempPENi), 4)) = UCase("PEN/") Then
    'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
    'If InStr(1, UCase(TempPENSplited(TempPENi)), UCase("PEN/"), vbTextCompare) > 0 Then
    If InStr(1, UCase(TempPENSplited(TempPENi)), UCase(PENLINEID_String & "PEN/"), vbTextCompare) > 0 Then
        'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
        TempPENP1E1 = TempPENSplited(TempPENi)
        For TempPENPNO = 1 To 99
            'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
            'If InStr(1, TempPENP1E1, "E" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, TempPENP1E1, "T" & TempPENPNO & "-", vbTextCompare) > 0 Then
            If InStr(1, TempPENP1E1, "E" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, TempPENP1E1, "T" & TempPENPNO & "-", vbTextCompare) > 0 Or InStr(1, TempPENP1E1, "DOB" & TempPENPNO & "-", vbTextCompare) > 0 Then
            'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
                Set CollPenSplited = RecPEN_PENLINEPassenger(TempPENP1E1, TempPENPNO)
                'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
                'CollPenPass.Add CollPenSplited, str(TempPENPNO)
                CollPenPass.Add CollPenSplited, str(TempPENi) & str(TempPENPNO)
            End If
            DoEvents
        Next
        'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
        'Set CollPenSplited = RecPEN_PENLINE(TempPENSplited(TempPENi))
        Set CollPenSplited = RecPEN_PENLINE(TempPENP1E1)
        CollPenline.Add CollPenSplited, str(TempPENi)
        'by Abhi on 05-Jun-2010 for caseid 1380 VLocator for worldspan is picking wrongly due to incorrect linking
        Temp1 = Replace(Temp1, TempPENSplited(TempPENi), "")
    End If
    DoEvents
Next

'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
'If InStr(1, UCase(TempPEN), UCase("PENFARE/"), vbTextCompare) > 0 Then
'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
'Do While InStr(1, UCase(TempPEN), UCase("PENFARE/"), vbTextCompare) > 0
Do While InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENFARE/"), vbTextCompare) > 0
    Set CollPenSplited = RecPEN_PENLINEPENFARE(TempPEN)
    If CollPenSplited.Count > 0 Then
        'by Abhi on 06-Apr-2010 for caseid 1293 PENFARE modification for Worldspan and Amadeus
        'CollPenPENFARE.Add CollPenSplited, CStr(CollPenSplited("SURNAMEFRSTNAMENO"))
        'by Abhi on 28-Apr-2010 for caseid 1205 PENFARE old format error in Worldspan in PenGDS will not stuck
        'CollPenPENFARE.Add CollPenSplited, CStr(CollPenSplited("PENFAREPASSTYPE"))
        CollPenPENFARE.Add CollPenSplited, CStr(PENFAREPNO_Long)
        PENFAREPNO_Long = PENFAREPNO_Long + 1
    End If
    DoEvents
Loop
'End If
'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENLINK/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENLINK(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENLINK.Add CollPenSplited, "PENLINK"
    End If
End If

'by Abhi on 10-Aug-2010 for caseid 1433 Penline PENAUTOOFF
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENAUTOOFF"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENAUTOOFF(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENAUTOOFF.Add CollPenSplited, "PENAUTOOFF"
    End If
End If

'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
Do While InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENO/"), vbTextCompare) > 0
    Set CollPenSplited = RecPEN_PENLINEPENO(TempPEN)
    If CollPenSplited.Count > 0 Then
        CollPenPENO.Add CollPenSplited, CStr(vPENONO_Long)
        vPENONO_Long = vPENONO_Long + 1
    End If
    DoEvents
Loop
'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENATOL/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENATOL(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENATOL.Add CollPenSplited, "PENATOL"
    End If
End If

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENRT-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENRT(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENRT.Add CollPenSplited, "PENRT"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENPOL-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENPOL(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENPOL.Add CollPenSplited, "PENPOL"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENPROJ-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENPROJ(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENPROJ.Add CollPenSplited, "PENPROJ"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENCC-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENCC(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENCC.Add CollPenSplited, "PENCC"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENEID-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENEID(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENEID.Add CollPenSplited, "PENEID"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENPO-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENPO(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENPO.Add CollPenSplited, "PENPO"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENHFRC-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENHFRC(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENHFRC.Add CollPenSplited, "PENHFRC"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENLFRC-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENLFRC(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENLFRC.Add CollPenSplited, "PENLFRC"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENHIGHF-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENHIGHF(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENHIGHF.Add CollPenSplited, "PENHIGHF"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENLOWF-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENLOWF(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENLOWF.Add CollPenSplited, "PENLOWF"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENUC1-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENUC1(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENUC1.Add CollPenSplited, "PENUC1"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENUC2-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENUC2(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENUC2.Add CollPenSplited, "PENUC2"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENUC3-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENUC3(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENUC3.Add CollPenSplited, "PENUC3"
    End If
End If
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENBB-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENBB(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENBB.Add CollPenSplited, "PENBB"
    End If
End If

'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENAGROSS/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENAGROSS(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENAGROSS.Add CollPenSplited, "PENAGROSS"
    End If
End If

'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
Do While InStr(1, UCase(TempPEN), UCase(PENLINEID_String & "PENAIRTKT/"), vbTextCompare) > 0
'If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENAIRTKT/"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENAIRTKT(TempPEN)
    If CollPenSplited.Count > 0 Then
        CollPenPENAIRTKT.Add CollPenSplited, CStr(PENAIRTKTPNO_Long)
        PENAIRTKTPNO_Long = PENAIRTKTPNO_Long + 1
    End If
'End If
    DoEvents
Loop

'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS

'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENBILLCUR"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENBILLCUR(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENBILLCUR.Add CollPenSplited, "PENBILLCUR"
    End If
End If
'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENWAIT"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENWAIT(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENWAIT.Add CollPenSplited, "PENWAIT"
    End If
End If
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT

'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENVC"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENVC(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENVC.Add CollPenSplited, "PENVC"
    End If
End If
'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus

'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENCS"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENCS(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENCS.Add CollPenSplited, "PENCS"
    End If
End If
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field

'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
If InStr(1, UCase(Temp1), UCase(PENLINEID_String & "PENRC-"), vbTextCompare) > 0 Then
    Set CollPenSplited = RecPEN_PENLINEPENRC(Temp1)
    If CollPenSplited.Count > 0 Then
        CollPenPENRC.Add CollPenSplited, "PENRC"
    End If
End If
'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

Set CollVL = Rec_N_VLocator(Temp1)

'SubSplited = ExtractBetween(temp1, "ICC-FT-SP-", "\", True)
'    Set CollSellRate = Rec_N_SellRate(SubSplited)

'CollGEN.Add "", "REMARKS"
CollGEN.Add Data, "REMARKS"

SubSplited = ExtractBetween(Temp1, "ICC-BOOKED-", "\")
CollRef.Add SubSplited, "REF"

'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
'SubSplited = ExtractBetween(Temp1, "TK-", "\")
SubSplited = ExtractBetween(Temp1, "\TK-", "\")
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    CollGEN.Add SubSplited, "TKTNO"
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
'SubSplited = ExtractBetween(Temp1, "IN-", "\")
SubSplited = ExtractBetween(Temp1, "\IN-", "\")
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    CollGEN.Add SubSplited, "INVNO"
    
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
'SubSplited = ExtractBetween(Temp1, "PID-", "\")
SubSplited = ExtractBetween(Temp1, "\PID-", "\")
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    CollGEN.Add SubSplited, "PREMARKS"
    temp = SubSplited
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
'SubSplited = ExtractBetween(Temp1, "LNS-", "\")
SubSplited = ExtractBetween(Temp1, "\LNS-", "\")
'by Abhi on 02-Jul-2016 for caseid 6551 PenGDS Error Multiple-step operation generated errors due to data length of WSPGENREMARKS.INVNO
    CollGEN.Add SubSplited, "LREMARKS"
    temp = temp & SubSplited

CollRemarks.Add temp, "Remarks"

Coll.Add CollGEN, "GEN"
Coll.Add CollSellRate, "SEL"
Coll.Add CollVL, "VL"
Coll.Add CollRef, "REF"
Coll.Add CollRemarks, "REM"
Coll.Add CollPenline, "PEN"
Coll.Add CollPenPass, "PENPASSENGER"
'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
Coll.Add CollPenPENFARE, "PENFARE"
'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
Coll.Add CollPenPENLINK, "PENLINK"
'by Abhi on 10-Aug-2010 for caseid 1433 Penline PENAUTOOFF
Coll.Add CollPenPENAUTOOFF, "PENAUTOOFF"
'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
Coll.Add CollPenPENO, "PENO"
'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
Coll.Add CollPenPENATOL, "PENATOL"
'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Coll.Add CollPenPENRT, "PENRT"
Coll.Add CollPenPENPOL, "PENPOL"
Coll.Add CollPenPENPROJ, "PENPROJ"
Coll.Add CollPenPENCC, "PENCC"
Coll.Add CollPenPENEID, "PENEID"
Coll.Add CollPenPENPO, "PENPO"
Coll.Add CollPenPENHFRC, "PENHFRC"
Coll.Add CollPenPENLFRC, "PENLFRC"
Coll.Add CollPenPENHIGHF, "PENHIGHF"
Coll.Add CollPenPENLOWF, "PENLOWF"
Coll.Add CollPenPENUC1, "PENUC1"
Coll.Add CollPenPENUC2, "PENUC2"
Coll.Add CollPenPENUC3, "PENUC3"
Coll.Add CollPenPENBB, "PENBB"
'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
Coll.Add CollPenPENAGROSS, "PENAGROSS"
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
Coll.Add CollPenPENAIRTKT, "PENAIRTKT"
'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
Coll.Add CollPenPENBILLCUR, "PENBILLCUR"
'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
Coll.Add CollPenPENWAIT, "PENWAIT"
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
Coll.Add CollPenPENVC, "PENVC"
'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
Coll.Add CollPenPENCS, "PENCS"
'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
Coll.Add CollPenPENRC, "PENRC"
'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
Set Rec_Remarks = Coll
End Function



'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER
'Private Function FindInitialAndName(mData)
'On Error GoTo errPara
'Dim aa, Data
'Data = mData
'Dim retn(1) As String
''by Abhi on 20-Apr-2015 for caseid 5145 Worldspan -Flat file loading issue
'Dim rsSelect As New ADODB.Recordset
'Dim vRecordCount_Long As Long
'Dim vi_Long As Long
'Dim vTITLE_String As String
'Dim vTITLELen_Integer As Integer
''by Abhi on 20-Apr-2015 for caseid 5145 Worldspan -Flat file loading issue
'
'Data = Replace(Data, "*CHD", "")
''by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
'Data = Replace(Data, "*YTH", "")
''by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
''by Abhi on 20-Apr-2015 for caseid 5145 Worldspan -Flat file loading issue
''aa = UCase(Right(Data, 2))
''If (aa = "MR") Then
''    retn(0) = Mid(Data, 1, Len(Data) - 2)
''    retn(1) = "MR"
''End If
''
''aa = UCase(Right(Data, 3))
''If (aa = "MRS") Then
''    retn(0) = Mid(Data, 1, Len(Data) - 3)
''    retn(1) = "MRS"
''End If
''
''aa = UCase(Right(Data, 4))
''If (aa = "MISS") Then
''    retn(0) = Mid(Data, 1, Len(Data) - 4)
''    retn(1) = "MISS"
''End If
''
''aa = UCase(Right(Data, 4))
''If (aa = "MSTR") Then
''    retn(0) = Mid(Data, 1, Len(Data) - 4)
''    retn(1) = "MSTR"
''End If
''
''aa = UCase(Right(Data, 4))
''If (aa = "PROF") Then
''    retn(0) = Mid(Data, 1, Len(Data) - 4)
''    retn(1) = "PROF"
''End If
'rsSelect.Open "SELECT TITLE FROM TitleMaster WITH (NOLOCK)", dbCompany, adOpenForwardOnly, adLockReadOnly
'vRecordCount_Long = rsSelect.RecordCount
'For vi_Long = 1 To vRecordCount_Long
'    vTITLE_String = Trim(SkipNull(rsSelect.Fields("TITLE")))
'    vTITLELen_Integer = Len(vTITLE_String)
'    If UCase(vTITLE_String) = UCase(Right(Data, vTITLELen_Integer)) Then
'        'by Abhi on 27-Jul-2015 for caseid 5437 World span title separation issue
'        'retn(0) = Replace(Data, vTITLE_String, "")
'        retn(0) = Left(Data, Len(Data) - vTITLELen_Integer)
'        'by Abhi on 27-Jul-2015 for caseid 5437 World span title separation issue
'        retn(1) = vTITLE_String
'        Exit For
'    End If
'    DoEvents
'    rsSelect.MoveNext
'Next
'rsSelect.Close
'Set rsSelect = Nothing
''by Abhi on 20-Apr-2015 for caseid 5145 Worldspan -Flat file loading issue
'
'If retn(0) = "" Then
'    retn(0) = Data
'    retn(1) = ""
'End If
'
'FindInitialAndName = retn
'Exit Function
'errPara:
'
'If retn(0) = "" Then
'    retn(0) = Data
'    retn(1) = ""
'End If
'End Function
'by Abhi on 22-Nov-2016 for caseid 6804 Title master is missing the tag MASTER

Private Function Rec_N_SellRate(mData)
Dim Splited, SubSplited
Dim Coll As New Collection, coll2 As New Collection
Splited = SplitWithLengths(mData, 10, 10)


If Len(mData) = 10 Then
    SubSplited = SplitWithLengths(Splited(0), 3, 4, 3)
        Coll.Add SubSplited(0), "CURCODE"
        Coll.Add SubSplited(1), "AMOUNT"
        Coll.Add SubSplited(2), "TYPE"
    coll2.Add Coll
    SubSplited = SplitWithLengths(Splited(1), 3, 4, 3)
        Set Coll = New Collection
        Coll.Add SubSplited(0), "CURCODE"
        Coll.Add SubSplited(1), "AMOUNT"
        Coll.Add SubSplited(2), "TYPE"
    coll2.Add Coll
ElseIf (Val(mData) > 0 And CStr(Val(mData)) = mData) Then
        'by Abhi on 01-Jul-2014 for caseid 4222 Exchange value picking logic-Hardcoded currency "GBP" should be replaced with company currency for Galileo
        'Coll.Add "GBP", "CURCODE"
        Coll.Add COMCID_String, "CURCODE"
        'by Abhi on 01-Jul-2014 for caseid 4222 Exchange value picking logic-Hardcoded currency "GBP" should be replaced with company currency for Galileo
        Coll.Add mData, "AMOUNT"
        Coll.Add "ADT", "TYPE"
    coll2.Add Coll
    SubSplited = SplitWithLengths(Splited(1), 3, 4, 3)
        Set Coll = New Collection
        'by Abhi on 01-Jul-2014 for caseid 4222 Exchange value picking logic-Hardcoded currency "GBP" should be replaced with company currency for Galileo
        'Coll.Add "GBP", "CURCODE"
        Coll.Add COMCID_String, "CURCODE"
        'by Abhi on 01-Jul-2014 for caseid 4222 Exchange value picking logic-Hardcoded currency "GBP" should be replaced with company currency for Galileo
        Coll.Add "0", "AMOUNT"
        Coll.Add "CHD", "TYPE"
    coll2.Add Coll
Else
    

End If
End Function

Private Function Rec_N_VLocator(mData)
Dim Splited, SubSplited, temp, Res, Counter
Dim Coll As New Collection, coll2 As New Collection
    temp = mData
    Counter = 1
    Res = " "
'by Abhi on 10-Nov-2009 for VLOCATOR in AirSegDetails
'While Len(Res) > 0
For Counter = 1 To mITNRYSEGNO_Long
    Res = ExtractBetween(temp, "- " & Counter & " ", "\", True)
    If Len(Res) > 0 Then
        Set Coll = New Collection
        Coll.Add Left(Res, 2), "AIRCODE"
        Coll.Add Trim(Right(Res, 6)), "VLOCATOR"
        Coll.Add Counter, "SLNO"
        coll2.Add Coll
    End If
    'Counter = Counter + 1
'Wend
    DoEvents
Next
    Set Rec_N_VLocator = coll2
End Function
Private Function Rec_N_Referance(mData)
    Dim Splited, SubSplited
    Dim Coll As New Collection, coll2 As New Collection
    Coll.Add mData, "REF"
End Function

Private Function Rec_N_Remarks(mData)
    Dim Splited, SubSplited
    Dim Coll As New Collection
    Coll.Add mData, "Remarks"
End Function



Private Function InsertDataCollectionINDEX(Coll As Collection, TableNae As String, LineID, UploadNo As Long)
For j = 1 To Coll.Count
    InsertData Coll(j), TableNae, CStr(LineID), UploadNo
    DoEvents
Next
End Function
Private Function InsertDataCollectionKey(Coll As Collection, TableNae As String, LineID, UploadNo As Long)
For j = 1 To Coll.Count
    InsertDataByFieldName Coll(j), TableNae, CStr(LineID), UploadNo
    DoEvents
Next
End Function


Private Function TryParseSellRate1(mData, mColl As Collection) As Boolean
Dim Coll As New Collection
Dim coll2 As New Collection
Dim Splited, SubSplited
If Len(mData) <> 10 Then TryParseSellRate1 = False: Exit Function
    SubSplited = SplitWithLengths(Splited(0), 3, 4, 3)
        Coll.Add SubSplited(0), "CURCODE"
        Coll.Add SubSplited(1), "AMOUNT"
        If Val(SubSplited(1)) = 0 Then TryParseSellRate1 = False: Exit Function
        Coll.Add SubSplited(2), "TYPE"
    coll2.Add Coll
    SubSplited = SplitWithLengths(Splited(1), 3, 4, 3)
        Set Coll = New Collection
        Coll.Add SubSplited(0), "CURCODE"
        Coll.Add SubSplited(1), "AMOUNT"
        If Val(SubSplited(1)) = 0 Then TryParseSellRate1 = False: Exit Function
        Coll.Add SubSplited(2), "TYPE"
    coll2.Add Coll

End Function

Private Function RecX_EXCHVALUE(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
    Splited = SplitForce(Data, "\", 57)
    
    Coll.Add Splited(1), "FINDICATOR"
    
    Coll.Add "NTF-", "TFAREFID"
    SubSplited = ExtractBetween(Data, "NTF-", "\")
    Coll.Add SubSplited, "TFAREAMT"
    
    Coll.Add "OBF-", "OBASEFID"
    SubSplited = ExtractBetween(Data, "OBF-", "\")
    Coll.Add SubSplited, "BFARE"
    
    Coll.Add "OT1-", "OTAX1FID"
    SubSplited = ExtractBetween(Data, "OT1-", "\")
    Coll.Add Left(SubSplited, 2), "OTAX1CODE"
    Coll.Add Right(SubSplited, 8), "OTAX1AMT"
    
    Coll.Add "OT2-", "OTAX2FID"
    SubSplited = ExtractBetween(Data, "OT2-", "\")
    Coll.Add Left(SubSplited, 2), "OTAX2CODE"
    Coll.Add Right(SubSplited, 8), "OTAX2AMT"
    
    Coll.Add "OT3-", "OTAX3FID"
    SubSplited = ExtractBetween(Data, "OT3-", "\")
    Coll.Add Left(SubSplited, 2), "OTAX3CODE"
    Coll.Add Right(SubSplited, 8), "OTAX3AMT"
    
    Coll.Add "OTF-", "OTCOSTFID"
    SubSplited = ExtractBetween(Data, "OTF-", "\")
    Coll.Add SubSplited, "TCOST"
    
    Coll.Add "AF-", "ADMFEEID"
    SubSplited = ExtractBetween(Data, "AF-", "\")
    Coll.Add SubSplited, "ADMFEEAMT"
    
    Coll.Add "RFC-", "CURCODEFID"
    SubSplited = ExtractBetween(Data, "RFC-", "\")
    Coll.Add SubSplited, "CURCODERFARE"
    
    Coll.Add "RF-", "RISSUEFID"
    SubSplited = ExtractBetween(Data, "RF-", "\")
    Coll.Add SubSplited, "RISSUEFAMT"
    
    Coll.Add "BFC-", "CURBFARE"
    SubSplited = ExtractBetween(Data, "BFC-", "\")
    Coll.Add SubSplited, "CURCODE"
    
    Coll.Add "BF-", "BFAREEX"
    SubSplited = ExtractBetween(Data, "BF-", "\")
    Coll.Add SubSplited, "BFAREEXAMT"
    
    Coll.Add "PC-", "PCURCODEID"
    SubSplited = ExtractBetween(Data, "PC-", "\")
    Coll.Add SubSplited, "PCURCODE"
    
    Coll.Add "T1-", "TAX1FID"
    SubSplited = ExtractBetween(Data, "T1-", "\")
    Coll.Add Left(SubSplited, 2), "TAX1CODE"
    Coll.Add Mid(SubSplited, 3, 8), "TAX1EXAMT"
    Coll.Add Right(SubSplited, 2), "TAX1STSIND"
    
    Coll.Add "T2-", "TAX2FID"
    SubSplited = ExtractBetween(Data, "T2-", "\")
    Coll.Add Left(SubSplited, 2), "TAX2CODE"
    Coll.Add Mid(SubSplited, 3, 8), "TAX2EXAMT"
    Coll.Add Right(SubSplited, 2), "TAX2STSIND"
    
    Coll.Add "T3-", "TAX3FID"
    SubSplited = ExtractBetween(Data, "T3-", "\")
    Coll.Add Left(SubSplited, 2), "TAX3CODE"
    Coll.Add Mid(SubSplited, 3, 8), "TAX3EXAMT"
    Coll.Add Right(SubSplited, 2), "TAX3STSIND"
    
    Coll.Add "AP-", "APFEEID"
    SubSplited = ExtractBetween(Data, "AP-", "\")
    Coll.Add SubSplited, "APFEEAMT"
    
    Coll.Add "TFC-", "CURCODETFC"
    SubSplited = ExtractBetween(Data, "TFC-", "\")
    Coll.Add SubSplited, "CURCODETFARE"
    
    Coll.Add "TF-", "TEXCFAREID"
    SubSplited = ExtractBetween(Data, "TF-", "\")
    Coll.Add SubSplited, "TEXCAMT"
    
    Coll.Add "MA-", "MCOAID"
    SubSplited = ExtractBetween(Data, "MA-", "\")
    Coll.Add SubSplited, "MCOAMT"
    
    Coll.Add "P1-", "PFC1ID"
    SubSplited = ExtractBetween(Data, "P1-", "\")
    Coll.Add SubSplited, "PFC1"
    
    Coll.Add "P2-", "PFC2ID"
    SubSplited = ExtractBetween(Data, "P2-", "\")
    Coll.Add SubSplited, "PFC2"
    
    Coll.Add "P3-", "PFC3ID"
    SubSplited = ExtractBetween(Data, "P3-", "\")
    Coll.Add SubSplited, "PFC3"
    
    Coll.Add "P4-", "PFC4ID"
    SubSplited = ExtractBetween(Data, "P4-", "\")
    Coll.Add SubSplited, "PFC4"
    
    Set RecX_EXCHVALUE = Coll
End Function


Private Function RecT_WSPTVLSEG(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Temp1
Dim Coll As New Collection
    Splited = SplitForce(Data, "\", 103)
    
    Coll.Add Splited(1), "SEGNO"
    
    SubSplited = SplitWithLengths(Splited(2), 2, 2, 3, 3, 5, 5)
    Coll.Add SubSplited(0), "VCODE"
    Coll.Add SubSplited(1), "STSCODE"
    Coll.Add SubSplited(2), "QTY"
    Coll.Add SubSplited(3), "ATYPE"
    Coll.Add SubSplited(4), "SDATE"
    Coll.Add SubSplited(5), "EDATE"

    Coll.Add "Y1-", "BYRID"
    Temp1 = ExtractBetween(Data, "Y1-", "\")
    Coll.Add Temp1, "BYEAR"

    Coll.Add "Y2-", "EYRID"
    Temp1 = ExtractBetween(Data, "Y2-", "\")
    Coll.Add Temp1, "EYEAR"
    
    Coll.Add "AA1-", "AAI"
    Temp1 = ExtractBetween(Data, "AA1-", "\")
    Coll.Add Temp1, "AADD1"
    
    Coll.Add "AA2-", "AA2"
    Temp1 = ExtractBetween(Data, "AA2-", "\")
    Coll.Add Temp1, "AADD2"
    
    Coll.Add "AC1-", "AC1"
    Temp1 = ExtractBetween(Data, "AC1-", "\")
    Coll.Add Temp1, "ACINFO1"
    
    Coll.Add "AC2-", "AC2"
    Temp1 = ExtractBetween(Data, "AC2-", "\")
    Coll.Add Temp1, "ACINFO2"
    
    Coll.Add "AGT-", "AGT"
    Temp1 = ExtractBetween(Data, "AGT-", "\")
    Coll.Add Temp1, "AGTNAME"
    
    Coll.Add "AN-", "AN"
    Temp1 = ExtractBetween(Data, "AN-", "\")
    Coll.Add Temp1, "ANAME"
    
    Coll.Add "AP-", "AP"
    Temp1 = ExtractBetween(Data, "AP-", "\")
    Coll.Add Temp1, "APHONE"
    
    Coll.Add "ARR-", "ARR"
    Temp1 = ExtractBetween(Data, "ARR-", "\")
    Coll.Add Temp1, "ARRINFO"
    
    Coll.Add "BC-", "BC"
    Temp1 = ExtractBetween(Data, "BC-", "\")
    Coll.Add Temp1, "BCODE"
    
    Coll.Add "BD-", "BD"
    Temp1 = ExtractBetween(Data, "BD-", "\")
    Coll.Add Temp1, "BDATE"
    
    Coll.Add "BED-", "BED"
    Temp1 = ExtractBetween(Data, "BED-", "\")
    Coll.Add Temp1, "BEDCOFIG"
    
    Coll.Add "BN-", "BN"
    Temp1 = ExtractBetween(Data, "BN-", "\")
    Coll.Add Temp1, "BNAME"
    
    Coll.Add "BS-", "BS"
    Temp1 = ExtractBetween(Data, "BS-", "\")
    Coll.Add Temp1, "BSOURCE"
    
    Coll.Add "CAT-", "CAT"
    Temp1 = ExtractBetween(Data, "CAT-", "\")
    Coll.Add Temp1, "CATEGARY"
    
    Coll.Add "CBN-", "CBN"
    Temp1 = ExtractBetween(Data, "CBN-", "\")
    Coll.Add Temp1, "CBNDESC"
    
    Coll.Add "CBP-", "CBP"
    Temp1 = ExtractBetween(Data, "CBP-", "\")
    Coll.Add Temp1, "CBPOSTN"
    
    Coll.Add "CC1-", "CC1"
    Temp1 = ExtractBetween(Data, "CC1-", "\")
    Coll.Add Temp1, "CCODE1"
    
    Coll.Add "CC2-", "CC2"
    Temp1 = ExtractBetween(Data, "CC2-", "\")
    Coll.Add Temp1, "CCODE2"
    
    Coll.Add "CD-", "CD"
    Temp1 = ExtractBetween(Data, "CD-", "\")
    Coll.Add Temp1, "CID"
    
    Coll.Add "CF-", "CF"
    Temp1 = ExtractBetween(Data, "CF-", "\")
    Coll.Add Temp1, "CFNO"
    
    Coll.Add "CL-", "CL"
    Temp1 = ExtractBetween(Data, "CL-", "\")
    Coll.Add Temp1, "CLSERVICE"
    
    Coll.Add "CK-", "CK"
    Temp1 = ExtractBetween(Data, "CK-", "\")
    Coll.Add Temp1, "CKNO"
    
   
    Coll.Add "CKD-", "CKD"
    Temp1 = ExtractBetween(Data, "CKD-", "\")
    Coll.Add Temp1, "CKDATE"
    
    Coll.Add "CM-", "CM"
    Temp1 = ExtractBetween(Data, "CM-", "\")
    SubSplited = SplitWithLengths(Temp1, 3, 10)
    Coll.Add SubSplited(0), "CURRCODE"
    Coll.Add SubSplited(1), "COMAMT"
    
    Coll.Add "FLT-", "FLT"
    Temp1 = ExtractBetween(Data, "FLT-", "\")
    Coll.Add Temp1, "FLTNO"
    
    Coll.Add "FOP-", "FOP"
    Temp1 = ExtractBetween(Data, "FOP-", "\")
    Coll.Add Temp1, "FOPAY"
    
    Coll.Add "ST-", "ST"
    Temp1 = ExtractBetween(Data, "ST-", "\")
    Coll.Add Temp1, "STIME"
    
    Coll.Add "TA-", "TA"
    Temp1 = ExtractBetween(Data, "TA-", "\")
    Coll.Add Temp1, "ARRTIME"
    
    Coll.Add "TD-", "TD"
    Temp1 = ExtractBetween(Data, "TD-", "\")
    Coll.Add Temp1, "DEPTIME"
    
    Set RecT_WSPTVLSEG = Coll
End Function

Private Function RecQ_DELIVERYADD(Data1) As Collection
Dim Splited
Dim Data
Data = Data1
Dim Coll As New Collection
    Splited = SplitForce(Data, "\", 6)
    
    Coll.Add Splited(0), "ADD1"
    Coll.Add Splited(1), "ADD2"
    Coll.Add Splited(2), "ADD3"
    Coll.Add Splited(3), "ADD4"
    Coll.Add Splited(4), "ADD5"
    Coll.Add Splited(5), "ADD6"
    
    Set RecQ_DELIVERYADD = Coll
End Function


Private Function RecJ_REFUND(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Splited = SplitForce(Data, "\", 14)
    
    Coll.Add Splited(1), "RFID"

    Coll.Add "RFC-", "RFC"
    Temp1 = ExtractBetween(Data, "RFC-", "\")
    Coll.Add Temp1, "RFCCURRCODE"
    
    Coll.Add "RF-", "RF"
    Temp1 = ExtractBetween(Data, "RF-", "\")
    Coll.Add Temp1, "BFARE"
    
    Coll.Add "T1-", "T1"
    Temp1 = ExtractBetween(Data, "T1-", "\")
    SubSplited = SplitWithLengths(Temp1, 2, 8)
    Coll.Add SubSplited(0), "TAXCODE1"
    Coll.Add SubSplited(1), "RFTAX1"
    
    Coll.Add "T2-", "T2"
    Temp1 = ExtractBetween(Data, "T2-", "\")
    SubSplited = SplitWithLengths(Temp1, 2, 8)
    Coll.Add SubSplited(0), "TAXCODE2"
    Coll.Add SubSplited(1), "RFTAX2"
    
    Coll.Add "T3-", "T3"
    Temp1 = ExtractBetween(Data, "T3-", "\")
    SubSplited = SplitWithLengths(Temp1, 2, 8)
    Coll.Add SubSplited(0), "TAXCODE3"
    Coll.Add SubSplited(1), "RFTAX3"
    
    Coll.Add "TFC-", "TFC"
    Temp1 = ExtractBetween(Data, "TFC-", "\")
    Coll.Add Temp1, "TFCURRCODE"
    
    Coll.Add "TF-", "TF"
    Temp1 = ExtractBetween(Data, "TF-", "\")
    Coll.Add Temp1, "TFAMT"
    
    Coll.Add "AP-", "AP"
    Temp1 = ExtractBetween(Data, "AP-", "\")
    Coll.Add Temp1, "APAMT"
    
    Coll.Add "P1-", "P1"
    Temp1 = ExtractBetween(Data, "P1-", "\")
    SubSplited = SplitWithLengths(Temp1, 3, 5)
    Coll.Add SubSplited(0), "P1AIRCODE"
    Coll.Add SubSplited(1), "P1AMT"
    
    Coll.Add "P2-", "P2"
    Temp1 = ExtractBetween(Data, "P2-", "\")
    SubSplited = SplitWithLengths(Temp1, 3, 5)
    Coll.Add SubSplited(0), "P2AIRCODE"
    Coll.Add SubSplited(1), "P2AMT"
    
    Coll.Add "P3-", "P3"
    Temp1 = ExtractBetween(Data, "P3-", "\")
    SubSplited = SplitWithLengths(Temp1, 3, 5)
    Coll.Add SubSplited(0), "P3AIRCODE"
    Coll.Add SubSplited(1), "P3AMT"
    
    Coll.Add "P4-", "P4"
    Temp1 = ExtractBetween(Data, "P4-", "\")
    SubSplited = SplitWithLengths(Temp1, 3, 5)
    Coll.Add SubSplited(0), "P4AIRCODE"
    Coll.Add SubSplited(1), "P4AMT"
    
    Coll.Add "ET-", "ET"
    Temp1 = ExtractBetween(Data, "ET-", "\")
    SubSplited = SplitForce(Temp1, "-", 3)
    Coll.Add SubSplited(0), "ETDOCNO"
    Coll.Add SubSplited(1), "CPNOS"
    Coll.Add SubSplited(2), "LAST3DIGITS"
    
    Set RecJ_REFUND = Coll
End Function

Private Function RecPEN_PENLINE(ByVal Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
Dim vi As Long
    Splited = SplitForce(Data, "/", 9)
    
    Temp1 = ExtractBetween(Data, "AC-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "AC"
    Coll.Add Left(Temp1, 50), "AC"
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    If Trim(Left(Temp1, 50)) <> "" Then
        GIT_CUSTOMERUSERCODE_String = Left(Temp1, 50)
    End If
    'by Abhi on 16-May-2017 for caseid 7427 Customer code in gds intray
    
    Temp1 = ExtractBetween(Data, "DES-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "DES"
    Coll.Add Left(Temp1, 50), "DES"
    
    Temp1 = ExtractBetween(Data, "REFE-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "REFE"
    Coll.Add Left(Temp1, 50), "REFE"
    
    Temp1 = ExtractBetween(Data, "SELL-ADT-", "/")
    Coll.Add Temp1, "SELLADT"
    
    Temp1 = ExtractBetween(Data, "SELL-CHD-", "/")
    Coll.Add Temp1, "SELLCHD"
    
    Temp1 = ExtractBetween(Data, "SCHARGE-", "/")
    Coll.Add Temp1, "SCHARGE"
    
    Temp1 = ExtractBetween(Data, "DEPT-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "DEPT"
    Coll.Add Left(Temp1, 50), "DEPT"
    
    Temp1 = ExtractBetween(Data, "BRANCH-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "BRANCH"
    Coll.Add Left(Temp1, 50), "BRANCH"
    
    Temp1 = ExtractBetween(Data, "CC-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "CUSTCC"
    Coll.Add Left(Temp1, 100), "CUSTCC"
            
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    'by Abhi on 06-Apr-2010 for caseid 1293 PENFARE modification for Worldspan and Amadeus
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    'by Abhi on 15-Apr-2010 for caseid 1301 Worldspan Delivery Address
    Temp1 = ExtractBetween(Data, "ADD-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "DeliAdd"
    Coll.Add Left(Temp1, 500), "DeliAdd"
    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
    Temp1 = ExtractBetween(Data, "MC-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "MC"
    Coll.Add Left(Temp1, 50), "MC"
    Temp1 = ExtractBetween(Data, "BB-", "/")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Temp1, "BB"
    Coll.Add Left(Temp1, 50), "BB"
    'by Abhi on 18-Jun-2010 for caseid 1399 Supplier of ticket for PENFARE
    Coll.Add "", "PENFARESUPPID"
    'by Abhi on 27-Jul-2010 for caseid 1439 Penline PENFARE Airline
    Coll.Add "", "PENFAREAIRID"
    'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
    Coll.Add "", "PENLINKPNR"
    'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
    Coll.Add "", "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Temp1 = ExtractBetween(Data, "INETREF-", "/")
    Coll.Add Left(Temp1, 100), "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add "", "PENOPAYMETHODID"
    'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Temp1 = ExtractBetween(Data, "TKTDEADLINE-", "/")
    Coll.Add Left(Temp1, 9), "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Temp1 = ExtractBetween(Data, "DEPOSITAMT-", "/")
    Coll.Add Val(Temp1), "DEPOSITAMT"
    Temp1 = ExtractBetween(Data, "DEPOSITDUEDATE-", "/")
    Coll.Add Left(Temp1, 9), "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

Set RecPEN_PENLINE = Coll
End Function

Private Function RecPEN_PENLINEPassenger(Data1, ByVal PNO As Long) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
        Coll.Add "", "AC"
        Coll.Add "", "DES"
        Coll.Add "", "REFE"
        Coll.Add "", "SELLADT"
        Coll.Add "", "SELLCHD"
        Coll.Add "", "SCHARGE"
        Coll.Add "", "DEPT"
        Coll.Add "", "BRANCH"
        Coll.Add "", "CUSTCC"
        
    'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
    'Temp1 = ExtractBetween(Data, "P" & PNO & "E-", "\")
    'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
    'Temp1 = ExtractBetween(Data, PENLINEID_String & "E" & PNO & "-", "\")
    Temp1 = ExtractBetween(Data, "E" & PNO & "-", "\")
        'by Abhi on 07-OCt-2009 limited up to 50 characters
        'by Abhi on 15-Apr-2010 for caseid 1303 remove slash (/) from penline email
        'by Abhi on 08-Mar-2018 for caseid 8331 Amadeus,Worldspan and Sabre - Pengds penline email replacement
        'Temp1 = Replace(Temp1, "/", "")
        Temp1 = PenlineEmailReplacement(Temp1)
        Temp1 = Replace(Temp1, "/", "")
        'by Abhi on 08-Mar-2018 for caseid 8331 Amadeus,Worldspan and Sabre - Pengds penline email replacement
        'by Abhi on 11-Oct-2017 for caseid 7862 Passenger Email id validation in folder should also save multiple email id with coma separator
        'Coll.Add Left(Temp1, 50), "PEMAIL"
        Coll.Add Left(Temp1, 300), "PEMAIL"
        'by Abhi on 11-Oct-2017 for caseid 7862 Passenger Email id validation in folder should also save multiple email id with coma separator
    'by Abhi on 23-Jun-2010 for caseid 1405 Client wise Penlines
    'Temp1 = ExtractBetween(Data, "P" & PNO & "T-", "\")
    'by Abhi on 07-Jul-2010 for caseid 1405 Client wise Penlines issue with T1 and E1
    'Temp1 = ExtractBetween(Data, PENLINEID_String & "T" & PNO & "-", "\")
    Temp1 = ExtractBetween(Data, "T" & PNO & "-", "\")
        'by Abhi on 07-OCt-2009 limited up to 50 characters
        'by Abhi on 15-Apr-2010 for caseid 1303 remove slash (/) from penline email
        Temp1 = Replace(Temp1, "/", "")
        Coll.Add Left(Temp1, 50), "PTELE"
        Coll.Add PNO, "PNO"
        'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
        Coll.Add 0, "FARE"
        Coll.Add 0, "TAXES"
        Coll.Add 0, "MARKUP"
        Coll.Add "", "SURNAMEFRSTNAMENO"
        'by Abhi on 06-Apr-2010 for caseid 1293 PENFARE modification for Worldspan and Amadeus
        Coll.Add 0, "PENFARESELL"
        Coll.Add "", "PENFAREPASSTYPE"
        'by Abhi on 15-Apr-2010 for caseid 1301 Worldspan Delivery Address
        Coll.Add "", "DeliAdd"
        'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
        Coll.Add "", "MC"
        Coll.Add "", "BB"
        'by Abhi on 18-Jun-2010 for caseid 1399 Supplier of ticket for PENFARE
        Coll.Add "", "PENFARESUPPID"
        'by Abhi on 27-Jul-2010 for caseid 1439 Penline PENFARE Airline
        Coll.Add "", "PENFAREAIRID"
        'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
        Coll.Add "", "PENLINKPNR"
        'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
        Coll.Add "", "PENOPRDID"
        Coll.Add 0, "PENOQTY"
        Coll.Add 0, "PENORATE"
        Coll.Add 0, "PENOSELL"
        Coll.Add "", "PENOSUPPID"
        'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
        Coll.Add "", "PENATOLTYPE"
        'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
        Coll.Add "", "INETREF"
        'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
        Coll.Add "", "PENOPAYMETHODID"
        'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
        Coll.Add "", "PENRT"
        Coll.Add "", "PENPOL"
        Coll.Add "", "PENPROJ"
        Coll.Add "", "PENEID"
        Coll.Add "", "PENPO"
        Coll.Add "", "PENHFRC"
        Coll.Add "", "PENLFRC"
        Coll.Add 0, "PENHIGHF"
        Coll.Add 0, "PENLOWF"
        Coll.Add "", "PENUC1"
        Coll.Add "", "PENUC2"
        Coll.Add "", "PENUC3"
        'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
        Coll.Add 0, "PENAGROSSADULT"
        Coll.Add 0, "PENAGROSSCHILD"
        Coll.Add 0, "PENAGROSSINFANT"
        Coll.Add 0, "PENAGROSSPACKAGE"
        'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
        Coll.Add 0, "PENAGROSSYOUTH"
        'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
        'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
        Coll.Add "", "TKTDEADLINE"
        'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
        'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
        Coll.Add "", "AIRTKTPAX"
        Coll.Add "", "AIRTKTTKT"
        Coll.Add "", "AIRTKTDATE"
        'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Temp1 = ExtractBetween(Data, "DOB" & PNO & "-", "\")
        Temp1 = Replace(Temp1, "/", "")
        Temp1 = Trim(Temp1)
        Coll.Add Left(Temp1, 9), "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

Data1 = Data
Set RecPEN_PENLINEPassenger = Coll
End Function

'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
Private Function RecPEN_PENLINEPENFARE(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    'by Abhi on 22-Mar-2010 for caseid 1205 PENFARE for Worldspan
    'by Abhi on 10-Oct-2015 for caseid 5645 PenGDS for Worldspan stuck due to missing backslash in PENFARE penline tag
    'Data = ExtractBetweenFrom(Data1, "PENFARE/", "\", False)
    Data = ExtractBetweenFrom(Data1, "PENFARE/", "\", True)
    'by Abhi on 10-Oct-2015 for caseid 5645 PenGDS for Worldspan stuck due to missing backslash in PENFARE penline tag
    'Data = Replace(Data, "-", "/")
    'by Abhi on 16-Mar-2010 for caseid 1205 PENFARE for Worldspan
    'by Abhi on 18-Jun-2010 for caseid 1399 Supplier of ticket for PENFARE
    'splited = SplitForce(Data, "/", 5)
    'by Abhi on 27-Jul-2010 for caseid 1439 Penline PENFARE Airline
    'splited = SplitForce(Data, "/", 6)
    Splited = SplitForce(Data, "/", 7)
    If Trim(Splited(4)) = "" Then
        'by Abhi on 10-Oct-2015 for caseid 5645 PenGDS for Worldspan stuck due to missing backslash in PENFARE penline tag
        Data1 = Replace(Data1, Data1, "")
        'by Abhi on 10-Oct-2015 for caseid 5645 PenGDS for Worldspan stuck due to missing backslash in PENFARE penline tag
        Exit Function
    End If
    Coll.Add Val(Splited(1)), "FARE"
    Coll.Add Val(Splited(2)), "TAXES"
    'by Abhi on 06-Apr-2010 for caseid 1293 PENFARE modification for Worldspan and Amadeus
    'Coll.Add Val(splited(3)), "MARKUP"
    'Coll.Add splited(4), "SURNAMEFRSTNAMENO"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    'by Abhi on 06-Apr-2010 for caseid 1293 PENFARE modification for Worldspan and Amadeus
    Coll.Add Val(Splited(3)), "PENFARESELL"
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add splited(4), "PENFAREPASSTYPE"
    Coll.Add Left(Splited(4), 50), "PENFAREPASSTYPE"
    'by Abhi on 15-Apr-2010 for caseid 1301 Worldspan Delivery Address
    Coll.Add "", "DeliAdd"
    'by Abhi on 18-Jun-2010 for caseid 1394 Marketing code and booked by penline
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    'by Abhi on 18-Jun-2010 for caseid 1399 Supplier of ticket for PENFARE
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add splited(5), "PENFARESUPPID"
    Coll.Add Left(Splited(5), 50), "PENFARESUPPID"
    'by Abhi on 27-Jul-2010 for caseid 1439 Penline PENFARE Airline
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add splited(6), "PENFAREAIRID"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'Coll.Add Left(Splited(6), 50), "PENFAREAIRID"
    Coll.Add "", "PENFAREAIRID"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
    Coll.Add "", "PENLINKPNR"
    'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
    Coll.Add "", "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Coll.Add "", "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add "", "PENOPAYMETHODID"
    'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add Left(Splited(6), 50), "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENFARE = Coll
End Function

'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
Private Function RecPEN_PENLINEPENLINK(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    'by Abhi on 06-Aug-2010 for caseid 1447 Penline PENLINK
    Data = ExtractBetween(Data1, PENLINEID_String & "PENLINK/", "\", False)
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Data, "PENLINKPNR"
    Coll.Add Left(Data, 50), "PENLINKPNR"
    'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    'by Abhi on 09-Oct-2010 for caseid 1511 Penline PENATOL
    Coll.Add "", "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Coll.Add "", "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add "", "PENOPAYMETHODID"
    'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENLINK = Coll
End Function

'by Abhi on 10-Aug-2010 for caseid 1433 Penline PENAUTOOFF
Private Function RecPEN_PENLINEPENAUTOOFF(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
    Coll.Add "", "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Coll.Add "", "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add "", "PENOPAYMETHODID"
    'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENAUTOOFF = Coll
End Function

'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
Private Function RecPEN_PENLINEPENO(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add 0, "FARE"
    Coll.Add 0, "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add 0, "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    'by Abhi on 24-Aug-2010 for caseid 1473 Penline PENO
    Data = ExtractBetweenFrom(Data1, PENLINEID_String & "PENO/", "\", False)
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    'splited = SplitForce(Data, "/", 6)
    Splited = SplitForce(Data, "/", 7)
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add splited(1), "PENOPRDID"
    Coll.Add Left(Splited(1), 50), "PENOPRDID"
    Coll.Add Val(Splited(2)), "PENOQTY"
    Coll.Add Val(Splited(3)), "PENORATE"
    Coll.Add Val(Splited(4)), "PENOSELL"
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add splited(5), "PENOSUPPID"
    Coll.Add Left(Splited(5), 50), "PENOSUPPID"
    'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
    Coll.Add "", "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Coll.Add "", "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add Left(Splited(6), 50), "PENOPAYMETHODID"
    'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENO = Coll
End Function

'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
Private Function RecPEN_PENLINEPENATOL(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    'by Abhi on 02-Oct-2010 for caseid 1511 Penline PENATOL
    Data = ExtractBetween(Data1, PENLINEID_String & "PENATOL/", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    'by Abhi on 23-Jul-2011 for caseid 1826 Multiple-step operation generated error in pengds penline REFE length
    'Coll.Add Data, "PENATOLTYPE"
    Coll.Add Left(Data, 50), "PENATOLTYPE"
    'by Abhi on 28-Jan-2010 for caseid 1560 INETREF for receipt update from online populating from pnr file to folder
    Coll.Add "", "INETREF"
    'by Abhi on 20-Oct-2011 for caseid 1889 Payment method in Default Charges(SAFI) module
    Coll.Add "", "PENOPAYMETHODID"
    'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENATOL = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENRT(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENRT-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 50), "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENRT = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENPOL(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENPOL-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 50), "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENPOL = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENPROJ(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENPROJ-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 50), "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENPROJ = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENCC(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENCC-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 100), "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    Coll.Add "", "PDOB"
    'by Abhi on 21-Mar-2016 for caseid 6121 Passenger DOB on folder passenger list
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENCC = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENEID(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENEID-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 50), "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENEID = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENPO(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENPO-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 255), "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENPO = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENHFRC(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENHFRC-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 50), "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENHFRC = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENLFRC(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENLFRC-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 50), "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENLFRC = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENHIGHF(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1

    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENHIGHF-", "\", False)
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    'Data = Replace(Data, "/", "")
    'Data = Replace(Data, "\", "")
    'Coll.Add Val(Data), "PENHIGHF"
    Data = Replace(Data, "\", "")
    Data = Replace(Data, "/", "-", , , vbTextCompare)
    Splited = Split(Data, "-")
    ReDim Preserve Splited(2)
    Coll.Add Val(Splited(0)), "PENHIGHF"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add Left(Splited(2), 50), "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENHIGHF = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENLOWF(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENLOWF-", "\", False)
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    'Data = Replace(Data, "/", "")
    'Data = Replace(Data, "\", "")
    'Coll.Add Val(Data), "PENLOWF"
    Data = Replace(Data, "\", "")
    Data = Replace(Data, "/", "-", , , vbTextCompare)
    Splited = Split(Data, "-")
    ReDim Preserve Splited(2)
    Coll.Add Val(Splited(0)), "PENLOWF"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add Left(Splited(2), 50), "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENLOWF = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENUC1(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENUC1-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 50), "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENUC1 = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENUC2(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENUC2-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 50), "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENUC2 = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENUC3(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENUC3-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 50), "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENUC3 = Coll
End Function

'by Abhi on 14-Jun-2012 for caseid 2123 Worldspan Penline Corporate booking for airplus and barclays
Private Function RecPEN_PENLINEPENBB(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENBB-", "\", False)
    Data = Replace(Data, "/", "")
    Data = Replace(Data, "\", "")
    Coll.Add Left(Data, 50), "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENBB = Coll
End Function


Private Function RecV_TRIPVALUE(Data1) As Collection
Dim Splited
Dim SubSplited
Dim SubSubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Splited = SplitForce(Data, "\", 14)
    
    Coll.Add Splited(1), "FCI"

    Coll.Add Splited(2), "PTC"
    If Len(Splited(3)) = 5 Then
        Temp1 = Mid(Splited(3), 2, 3)
    Else
        Temp1 = ""
    End If
    Coll.Add Temp1, "ORGAIRCODE"

    Coll.Add "XT", "XT"
    Temp1 = ExtractBetweenFrom(Data, "YQ", "\", , 3)
    SubSplited = SplitWithLengths(Temp1, 3, 2, 6)
        Coll.Add SubSplited(0), "CURRCODE"
        Coll.Add SubSplited(1), "TAXCODE"
    Temp1 = ToCurrency(SubSplited(2))
        Coll.Add Temp1, "TAXVALUE"

    Coll.Add "ROE", "ROE"
    Temp1 = ExtractFrom(Data, "ROE")
    SubSplited = SplitForce(Temp1, "\", 5)
    SubSubSplited = SplitWithLengths(SubSplited(0), 9)
        Coll.Add SubSubSplited(0), "BLANK"
    SubSubSplited = SplitWithLengths(SubSplited(1), 1, 10, 1)
        Coll.Add SubSubSplited(0), "BLANK2"
        Coll.Add SubSubSplited(1), "ROE2"
        Coll.Add SubSubSplited(2), "BLANK3"
    SubSubSplited = SplitWithLengths(SubSplited(2), 5, 1, 6)
        Coll.Add SubSubSplited(0), "TRIPID"
        Coll.Add SubSubSplited(1), "NUC"
        Coll.Add SubSubSplited(2), "TTVALUE"
    SubSubSplited = SplitWithLengths(SubSplited(3), 1, 3, 2, 6)
        Coll.Add SubSubSplited(0), "NO1"
        Coll.Add SubSubSplited(1), "CCURRCODE"
        Coll.Add SubSubSplited(2), "SLASH"
        Coll.Add SubSubSplited(3), "BLANK4"
    SubSubSplited = SplitWithLengths(SubSplited(4), 8, 3)
        Coll.Add SubSubSplited(0), "TCURRRATIO"
        Coll.Add SubSubSplited(1), "TCURRCODE"
    

    Set RecV_TRIPVALUE = Coll
End Function

Private Function ToCurrency(ByVal vlsData As String) As String
Dim result As String
    If Trim(vlsData) <> "" Then
        If InStr(1, vlsData, ".", vbTextCompare) = 0 Then
            result = Left(vlsData, Len(vlsData) - 2) & "." & Right(vlsData, 2)
        Else
            result = vlsData
        End If
    Else
        result = vlsData
    End If
ToCurrency = result
End Function

'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
Private Function RecPEN_PENLINEPENAGROSS(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    'by Abhi on 30-Nov-2012 for caseid 2653 Penline Agent Gross Invoice
    Data = ExtractBetweenFrom(Data1, PENLINEID_String & "PENAGROSS/", "\", False)
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'splited = SplitForce(Data, "/", 5)
    Splited = SplitForce(Data, "/", 6)
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add Val(Splited(1)), "PENAGROSSADULT"
    Coll.Add Val(Splited(2)), "PENAGROSSCHILD"
    Coll.Add Val(Splited(3)), "PENAGROSSINFANT"
    Coll.Add Val(Splited(4)), "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add Val(Splited(5)), "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 06-Nov-2014 for caseid 4684 Ticketing Deadline in Airticket tab and create penline
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENAGROSS = Coll
End Function


Private Function RecPEN_PENLINEPENAIRTKT(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1, temp2
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add 0, "PENAGROSSYOUTH"
    'by Abhi on 08-Mar-2016 for caseid 6124 New field for Youth in penline PENAGROSS
    Coll.Add "", "TKTDEADLINE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    Data = ExtractBetweenFrom(Data1, PENLINEID_String & "PENAIRTKT/", "\", False)
    Temp1 = ExtractBetween(Data, "PAX-", "/")
        Coll.Add Left(Temp1, 150), "AIRTKTPAX"
    Temp1 = ExtractBetween(Data, "TKT-", "/")
        Coll.Add Left(Temp1, 50), "AIRTKTTKT"
    Temp1 = ExtractBetween(Data, "DATE-", "/")
        Coll.Add Left(Temp1, 9), "AIRTKTDATE"
    'by Abhi on 24-Jun-2015 for caseid 5344 New Penline PENAIRTKT for All GDS
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    Coll.Add "", "PDOB"
    'by Abhi on 06-Apr-2016 for caseid 6226 PenGDS Worldspan Invalid procedure call or argument due to new penline field DOB
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENAIRTKT = Coll
End Function

'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency
Private Function RecPEN_PENLINEPENBILLCUR(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    Coll.Add 0, "PENAGROSSYOUTH"
    Coll.Add "", "TKTDEADLINE"
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    Coll.Add "", "PDOB"
    Data = ExtractBetweenFrom(Data1, PENLINEID_String & "PENBILLCUR", "\", False)
    Splited = SplitForce(Data, "-", 2)
    Coll.Add Left(Splited(1), 4), "PENBILLCUR"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENBILLCUR = Coll
End Function
'by Abhi on 13-Mar-2017 for caseid 7219 New penline for billing currency

'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan
Private Function Rec1_TicketableSegment_NonTicketableSegment_Operating(Data, ByVal pUploadNo_Long As Long)
Dim Splited
Dim SubSplit
Dim vOperatingAIRNAME_String As String
Dim vOperatingAIRID_String As String
Dim vOperatingFlightNo_String As String
Dim vSQL_String As String
    
    'MsgBox "Data=" & Data & ", fOperatingITNRYSEGNO_Table_String=" & fOperatingITNRYSEGNO_Table_String & ", fOperatingITNRYSEGNO_Long=" & fOperatingITNRYSEGNO_Long
    
    Data = Replace(Data, "OPERATED BY", "", 1, , vbTextCompare)
    Data = Trim(Data)
    Splited = SplitForce(Data, "-", 2)
        vOperatingAIRNAME_String = Trim(Splited(0))
    
    SubSplit = SplitTwo(Trim(Splited(1)), 2)
        vOperatingAIRID_String = SubSplit(0)
        vOperatingFlightNo_String = SubSplit(1)

    vSQL_String = "" _
        & "UPDATE    dbo." & fOperatingITNRYSEGNO_Table_String & " " _
        & "SET              OperatingAIRNAME = '" & SkipChars(vOperatingAIRNAME_String) & "', OperatingAIRID = '" & SkipChars(vOperatingAIRID_String) & "', OperatingFlightNo = '" & SkipChars(vOperatingFlightNo_String) & "' " _
        & "WHERE     (ITNRYSEGNO = '" & fOperatingITNRYSEGNO_Long & "') AND (UpLoadNo = " & pUploadNo_Long & ")"
    dbCompany.Execute vSQL_String
End Function
'by Abhi on 02-Dec-2015 for caseid 5827 Operating flight details from Worldspan

'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT
Private Function RecPEN_PENLINEPENWAIT(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    Coll.Add 0, "PENAGROSSYOUTH"
    Coll.Add "", "TKTDEADLINE"
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    Coll.Add "", "PDOB"
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
        
    GIT_PENWAIT_String = "Y"
'Data1 = Data
Set RecPEN_PENLINEPENWAIT = Coll
End Function
'by Abhi on 21-Jun-2017 for caseid 7544 New penline for waiting the PNR in tray-PENWAIT

'by Abhi on 23-Jun-2017 for caseid 4527 OB ticketing fee mapping for Worldspan
Private Function RecY_OBFees(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1

    Splited = SplitForce(Data, "\", 4)
    
    Temp1 = GetPassengerTypefromShort(Splited(1), True)
    Coll.Add Temp1, "PassengerTypeCode"
    
    Coll.Add "CA-", "CarrierIdentifier"
    Temp1 = ExtractBetween(Splited(2), "CA-", "\")
    SubSplited = SplitForce(Temp1, "/", 2)
    Coll.Add SubSplited(0), "Alphacodecarrier"
    Coll.Add SubSplited(1), "Numericcodeofcarrier"
    
    Coll.Add "TOB-", "TotalOBAmountIdentifier"
    Temp1 = ExtractBetween(Splited(3), "TOB-", "\")
    Coll.Add Left(Temp1, 3), "CurrencyCode"
    Coll.Add Mid(Temp1, 4), "TotalamountofallOBfees"
    
    Coll.Add "OB1-", "FirstOBFee"
    Temp1 = ExtractBetween(Splited(4), "OB1-", "\")
    Coll.Add Left(Temp1, 3), "Subcode"
    Coll.Add Mid(Temp1, 4), "Amountofthisfee"
    
    Coll.Add "", "Taxidentifieroffirsttax"
    Coll.Add "", "Taxcodefirsttax"
    Coll.Add "", "Totalamountoffirsttax"
    
    Coll.Add "", "Taxidentifierofsecondtax"
    Coll.Add "", "Taxcodesecondtax"
    Coll.Add "", "Totalamountofsecondtax"
    
    Set RecY_OBFees = Coll
End Function
'by Abhi on 23-Jun-2017 for caseid 4527 OB ticketing fee mapping for Worldspan

'by Abhi on 31-Aug-2017 for caseid 7742 Worldspan EMD Ticktets from Record -E – For an Issued EMD
Private Function RecE_IssuedEMD(Data1) As Collection '"-E"
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1

    Temp1 = ExtractBetween(Data, "TS-", "\")
    Coll.Add "TS-", "TS-"
    Coll.Add Left(Temp1, 6), "TypeofService"
    
    Temp1 = ExtractBetween(Data, "BI-", "\")
    Coll.Add "BI-", "BI-"
    Coll.Add Left(Temp1, 2), "BookingIndicator"

    Temp1 = ExtractBetween(Data, "EID-", "\")
    Coll.Add "EID-", "EID-"
    Coll.Add Left(Temp1, 2), "EMDID"

    Temp1 = ExtractBetween(Data, "NN-", "\")
    Coll.Add "NN-", "NN-"
    Coll.Add Left(Temp1, 7), "SurnameFirstnameNumber"

    Temp1 = ExtractBetween(Data, "NM-", "\")
    Coll.Add "NM-", "NM-"
    Coll.Add Left(Temp1, 68), "PassengerName"

    Temp1 = ExtractBetween(Data, "RFI-", "\")
    Coll.Add "RFI-", "RFI-"
    Coll.Add Left(Temp1, 1), "ReasonforIssuanceCode"

    Temp1 = ExtractBetween(Data, "EMD-", "\")
    Coll.Add "EMD-", "EMD-"
    Coll.Add Mid(Trim(Temp1), 4, 14), "EMDNumber"

    Temp1 = ExtractBetween(Data, "VC-", "\")
    Coll.Add "VC-", "VC-"
    SubSplited = SplitForce(Temp1, "/", 2)
    Coll.Add Left(SubSplited(0), 3), "ValidatingCarrierCode"
    Coll.Add Left(SubSplited(1), 4), "ValidatingCarrierNumber"

    Temp1 = ExtractBetween(Data, "TL-", "\")
    Coll.Add "TL-", "TL-"
    SubSplited = SplitWithLengths(Temp1, 3, 12)
    Coll.Add Left(SubSplited(0), 3), "TotalCurrencyCode"
    Coll.Add Left(SubSplited(1), 12), "TotalAmount"

    Temp1 = ExtractBetween(Data, "FP1-", "\")
    Coll.Add "FP1-", "FP1-"
    SubSplited = SplitWithLengths(Temp1, 3, 45)
    Coll.Add Left(SubSplited(0), 3), "FormofPaymentCode"
    Coll.Add Left(SubSplited(1), 45), "FormofPaymentInfo"

    Set RecE_IssuedEMD = Coll
End Function
'by Abhi on 31-Aug-2017 for caseid 7742 Worldspan EMD Ticktets from Record -E – For an Issued EMD

'by Abhi on 15-Jan-2018 for caseid 8130 Company Card checking in upload files-Amadeus
Private Function RecPEN_PENLINEPENVC(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    Coll.Add 0, "PENAGROSSYOUTH"
    Coll.Add "", "TKTDEADLINE"
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    Coll.Add "", "PDOB"
    Coll.Add "", "PENBILLCUR"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    'by Abhi on 29-Jan-2019 for caseid 9799 Penline for Deposit due date /amount
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    Coll.Add "", "PENFARETICKETTYPE"
    'by Abhi on 12-Feb-2019 for caseid 9870 New penfare for reissue tickets for all GDS
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    'by Abhi on 16-Apr-2019 for caseid 10151 Create a new penline for CS Field
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENVC = Coll
End Function


Private Function RecPEN_PENLINEPENCS(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    Coll.Add 0, "PENAGROSSYOUTH"
    Coll.Add "", "TKTDEADLINE"
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    Coll.Add "", "PDOB"
    Coll.Add "", "PENBILLCUR"
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    Coll.Add "", "PENFARETICKETTYPE"
    Data = ExtractBetweenFrom(Data1, PENLINEID_String & "PENCS/", "\", False)
    Splited = SplitForce(Data, "/", 3)
    Coll.Add Left(Trim(Splited(1)), 20), "PENCSLABELID"
    Coll.Add Left(Trim(Splited(2)), 20), "PENCSLISTID"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Coll.Add "", "PENRC"
    Coll.Add "", "PENRCTKTNO"
    'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

'Data1 = Data
Set RecPEN_PENLINEPENCS = Coll
End Function

'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information
Private Function RecPEN_PENLINEPENRC(Data1) As Collection
Dim Splited
Dim SubSplited
Dim Data
Data = Data1
Dim Coll As New Collection
Dim Temp1
    Coll.Add "", "AC"
    Coll.Add "", "DES"
    Coll.Add "", "REFE"
    Coll.Add "", "SELLADT"
    Coll.Add "", "SELLCHD"
    Coll.Add "", "SCHARGE"
    Coll.Add "", "DEPT"
    Coll.Add "", "BRANCH"
    Coll.Add "", "CUSTCC"
    Coll.Add "", "PEMAIL"
    Coll.Add "", "PTELE"
    Coll.Add 0, "PNO"
    Coll.Add "0", "FARE"
    Coll.Add "0", "TAXES"
    Coll.Add 0, "MARKUP"
    Coll.Add "", "SURNAMEFRSTNAMENO"
    Coll.Add "0", "PENFARESELL"
    Coll.Add "", "PENFAREPASSTYPE"
    Coll.Add "", "DeliAdd"
    Coll.Add "", "MC"
    Coll.Add "", "BB"
    Coll.Add "", "PENFARESUPPID"
    Coll.Add "", "PENFAREAIRID"
    Coll.Add "", "PENLINKPNR"
    Coll.Add "", "PENOPRDID"
    Coll.Add 0, "PENOQTY"
    Coll.Add 0, "PENORATE"
    Coll.Add 0, "PENOSELL"
    Coll.Add "", "PENOSUPPID"
    Coll.Add "", "PENATOLTYPE"
    Coll.Add "", "INETREF"
    Coll.Add "", "PENOPAYMETHODID"
    Coll.Add "", "PENRT"
    Coll.Add "", "PENPOL"
    Coll.Add "", "PENPROJ"
    Coll.Add "", "PENEID"
    Coll.Add "", "PENPO"
    Coll.Add "", "PENHFRC"
    Coll.Add "", "PENLFRC"
    Coll.Add 0, "PENHIGHF"
    Coll.Add 0, "PENLOWF"
    Coll.Add "", "PENUC1"
    Coll.Add "", "PENUC2"
    Coll.Add "", "PENUC3"
    Coll.Add 0, "PENAGROSSADULT"
    Coll.Add 0, "PENAGROSSCHILD"
    Coll.Add 0, "PENAGROSSINFANT"
    Coll.Add 0, "PENAGROSSPACKAGE"
    Coll.Add 0, "PENAGROSSYOUTH"
    Coll.Add "", "TKTDEADLINE"
    Coll.Add "", "AIRTKTPAX"
    Coll.Add "", "AIRTKTTKT"
    Coll.Add "", "AIRTKTDATE"
    Coll.Add "", "PDOB"
    Coll.Add "", "PENBILLCUR"
    Coll.Add 0, "DEPOSITAMT"
    Coll.Add "", "DEPOSITDUEDATE"
    Coll.Add "", "PENFARETICKETTYPE"
    Coll.Add "", "PENCSLABELID"
    Coll.Add "", "PENCSLISTID"
    Coll.Add "", "PENHIGHFTKTNO"
    Coll.Add "", "PENLOWFTKTNO"
    Data = ExtractBetween(Data1, PENLINEID_String & "PENRC-", "\", False)
    Data = Replace(Data, "\", "")
    Data = Replace(Data, "/", "-", , , vbTextCompare)
    Splited = Split(Data, "-")
    ReDim Preserve Splited(2)
    Coll.Add Left(Splited(0), 50), "PENRC"
    Coll.Add Left(Splited(2), 50), "PENRCTKTNO"

'Data1 = Data
Set RecPEN_PENLINEPENRC = Coll
End Function
'by Abhi on 23-Jun-2019 for caseid 10409 Reason code for corporate ticket information

