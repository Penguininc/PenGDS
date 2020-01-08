Attribute VB_Name = "WorldSpan"
Dim PassengerToAirFairID As Long
Dim PassengerToAirFairFlag As Boolean
Public WorldSpan_AirFareSLNO As Long

Public Function SplitDelims(Data, ParamArray Delims())
Dim temp, mData, result(), mCount
mData = Data
mCount = UBound(Delims)
ReDim result(mCount + 1)
For j = 0 To mCount
    temp = SplitFirstTwo(mData, Delims(j))
    mData = temp(1)
    result(j) = temp(0)
Next
    result(j) = mData
SplitDelims = result
End Function



Public Function SplitTwo(Data, First As Long)
On Error Resume Next
Dim temp(1) As String
temp(0) = Left(Data, First)
temp(1) = Mid(Data, First + 1, Len(Data) - First)
SplitTwo = temp
End Function

Public Function SplitTwoReverse(Data, First As Long)
On Error Resume Next
Dim temp(1) As String
temp(1) = Right(Data, First)
temp(0) = Mid(Data, 1, Len(Data) - First)
SplitTwoReverse = temp
End Function

Public Function SplitFirstTwo(Data, find)
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

Public Function SplitField(Data, Start As String, Finish As String)
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

Public Function SplitWithLengths(Data, ParamArray Lengths())
Dim Nos As Integer
Dim pos As Integer, j As Integer
Nos = UBound(Lengths)
Dim temp() As String
ReDim temp(Nos) As String
pos = 1

For j = 0 To Nos
    temp(j) = Mid(Data, pos, Lengths(j))
    pos = pos + Lengths(j)
Next
SplitWithLengths = temp
End Function
Public Function SplitWithLengthsPlus(Data, ParamArray Lengths())
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
Next
temp(Nos + 1) = Mid(Data, pos, SkipNegative(Len(Data) - pos + 1))
SplitWithLengthsPlus = temp

End Function

Public Function SplitForce(Data, delimiter As String, minNos As Integer)
Dim temp() As String
Dim Nos As Integer
temp = Split(Data, delimiter)
Nos = UBound(temp) + 1
If Nos < minNos Then
ReDim Preserve temp(minNos)
End If
SplitForce = temp
End Function

Public Function SplitForcePlus(Data, delimiter As String, minNos As Integer)
Dim temp() As String
Dim Nos As Integer
temp = Split(Data, delimiter, minNos + 1)
Nos = UBound(temp) + 1
If Nos < minNos Then
ReDim Preserve temp(minNos)
End If
SplitForcePlus = temp
End Function
Public Function ExtractBetween(Data, startText, endText, Optional AllifNoValue As Boolean = True)
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


Public Function InsertData(Data As Collection, TableName As String, LineIdent As String, UploadNo As Long) As Boolean
Dim Rs As New ADODB.Recordset
Dim j As Integer
Rs.Open "Select * from " + TableName, dbCompany, adOpenDynamic, adLockPessimistic
Rs.AddNew
Rs.Fields(0) = UploadNo
Rs.Fields(1) = LineIdent
For j = 2 To Rs.Fields.Count - 1
    Rs.Fields(j) = Data(j - 1)
Next
Rs.Update
End Function
Public Function InsertDataByFieldName(Data As Collection, TableName As String, LineIdent As String, UploadNo As Long) As Boolean
Dim Rs As New ADODB.Recordset
Dim temp
Dim j As Integer
Rs.Open "Select * from " + TableName, dbCompany, adOpenDynamic, adLockPessimistic
Rs.AddNew
Rs.Fields("UpLoadNo") = UploadNo
Rs.Fields("RecID") = LineIdent
For j = 2 To Rs.Fields.Count - 1
    temp = Rs.Fields(j).Name
    Rs.Fields(temp) = Data(temp)
Next
Rs.Update
End Function


Public Function PostLine(Line As String, UploadNo As Long, Optional FNAME = "") As String
Dim temp As String, NewLine As String
Dim Coll As New Collection
temp = Left(Line, 1)
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
        Case "2"
            Set Coll = Rec1_NonTicketableSegment(NewLine)
            InsertData Coll, "WSPNTKTSEG", temp, UploadNo
        Case "3"
'            Set Coll = Rec33_TARUNKSegment(NewLine)
'            InsertData Coll, "WSPARUNK", temp, UploadNo
        Case "4"
            Set Coll = Rec4_4_AdditionalPTCs(NewLine)
            InsertData Coll, "WSPADDPTC", temp, UploadNo

        End Select
    '------------------------------------------------------------
    Case "A"
        Set Coll = RecA_NameDocumentNumbers(NewLine)
        InsertData Coll, "WSPPNAME", "1", UploadNo
    Case "D"
        Set Coll = RecD_OriginalIssueData(NewLine)
        InsertData Coll, "WSPORGISSDATA", "1", UploadNo
    Case "G"
        Set Coll = RecG_AirFair(NewLine)
        InsertData Coll, "WSPAIRFARE", "1", UploadNo
    Case "H"
        Set Coll = Rec1_HotelSegment(NewLine)
        InsertData Coll, "WSPHTLSEG", "1", UploadNo
    Case "E"
            Set Coll = Rec_Endorsement(NewLine)
            InsertDataByFieldName Coll, "WSPENDORSEMENT", "E", UploadNo
    Case "M"
            Set Coll = Rec_SSRData(NewLine)
            InsertDataByFieldName Coll, "WSPSSRDATA", "M", UploadNo
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


    
End Select

End Function


Public Function Rec1_TicketableSegment(Data) As Collection
Dim splited
Dim Coll As New Collection
Dim SubSplit, SubSplit2
splited = Split(Data, "\")

Coll.Add "1", "AIRSEGCODE"
Coll.Add splited(1), "ITNRYSEGNO"
SubSplit = SplitWithLengths(splited(2), 1, 1, 2, 4)
    Coll.Add SubSplit(0), "NOINTSTOP"
    Coll.Add SubSplit(1), "SHDESIGNINDI"
    Coll.Add SubSplit(2), "AIRCODE"
    Coll.Add SubSplit(3), "FLTNO"
Coll.Add splited(3), "CLASS"
SubSplit = SplitWithLengths(splited(4), 3, 5, 5)
    Coll.Add SubSplit(0), "ORGAIRCODE"
    Coll.Add SubSplit(1), "DEPDATE"
    Coll.Add SubSplit(2), "DEPTIME"
SubSplit = SplitWithLengths(splited(5), 3, 5, 5)
    Coll.Add SubSplit(0), "DESTAIRCODE"
    Coll.Add SubSplit(1), "ARRDATE"
    Coll.Add SubSplit(2), "ARRTIME"
Coll.Add splited(6), "BDATE"
Coll.Add splited(7), "ADATE"
Coll.Add splited(8), "BAGG"
SubSplit = SplitWithLengths(splited(9), 2, 1, 1) 'Some ModificationNeeded
    Coll.Add SubSplit(0), "STATUS"
    Coll.Add ToMealsServiceCode(SubSplit(1)), "MEALSERVCODE" 'Mealsdetails
    Coll.Add SubSplit(2), "SEGSTOPCODE"
Coll.Add splited(10), "EQUIPTYPE"
Coll.Add splited(11), "FBASISCODE"
Coll.Add splited(12), "SEGMILEAGE"
Coll.Add splited(13), "INTSTOP"
Coll.Add splited(14), "AFLTTIME"
Coll.Add splited(15), "SEGOVRRDIND"
Coll.Add splited(16), "DEPTRMNLCODE"
Coll.Add splited(17), "ARRTRMNLCODE"
SubSplit = SplitWithLengthsPlus(splited(9), 3)
Coll.Add SubSplit(0), "BPAYAMT"
    
Set Rec1_TicketableSegment = Coll
End Function

'Non-Ticketable Segment
Public Function Rec1_NonTicketableSegment(Data) As Collection

Dim splited
Dim Coll As New Collection
Dim SubSplit
splited = Split(Data, "\")

Coll.Add "2", "AIRSEGCODE"

Coll.Add splited(1), "ITNRYSEGNO"
SubSplit = SplitWithLengths(splited(2), 1, 1, 2, 4)
    Coll.Add SubSplit(0), "NOINTERSTOP"
    Coll.Add SubSplit(1), "SHDESIIND"
    Coll.Add SubSplit(2), "AIRCODE"
    Coll.Add SubSplit(3), "FLTNO"
Coll.Add splited(3), "Class of Service"
SubSplit = SplitWithLengths(splited(4), 3, 5, 5)
    Coll.Add SubSplit(0), "ORGAIRCODE"
    Coll.Add SubSplit(1), "DEPDATE"
    Coll.Add SubSplit(2), "DEPTIME"
SubSplit = SplitWithLengths(splited(5), 3, 5, 5)
    Coll.Add SubSplit(0), "DESTAIRCODE"
    Coll.Add SubSplit(1), "ARRDATE"
    Coll.Add SubSplit(2), "ARRTIME"

Coll.Add splited(6), "AS1"
Coll.Add splited(7), "AS2"
Coll.Add splited(8), "AS3"

SubSplit = SplitWithLengths(splited(9), 2, 42, 1)
    Coll.Add SubSplit(0), "STATUS"
    Coll.Add SubSplit(1), "MEALSERVCODE"
    Coll.Add SubSplit(2), "SEGSTOPCODE"


Coll.Add splited(10), "Equipment type"
Coll.Add splited(11), "AS4*"
Coll.Add splited(12), "Segment Mileage"
Coll.Add splited(13), "Intermediate Stops"
Coll.Add splited(14), "Accumulated Elapsed Flight Time"
Coll.Add splited(15), "Segment Override Indicator"
Coll.Add splited(16), "Departure airport terminal code"
Coll.Add splited(17), "Arrival airport terminal code"

Set Rec1_NonTicketableSegment = Coll
End Function


Public Function Rec1_PNRFileAddress(Data, Optional FNAME = "")
Dim splited
Dim SubSplited
Dim temp1, temp2
Dim Coll As New Collection

splited = SplitForcePlus(Data, "\", 2)
Coll.Add splited(1), "INTLVL"
temp1 = splited(2)

temp2 = ExtractBetween(temp1, "FA-", "\")
    Coll.Add temp2, "PNRADD"
temp2 = ExtractBetween(temp1, "IN-", "\")
    SubSplited = SplitForce(temp2, "-", 2)
        Coll.Add SubSplited(0), "FINVNO"
        Coll.Add SubSplited(1), "LINVNO"
temp2 = ExtractBetween(temp1, "NC-", "\")
    Coll.Add temp2, "ITNRYCHNGE"
temp2 = ExtractBetween(temp1, "LC-", "\")
    SubSplited = SplitForce(temp2, "/", 2)
        Coll.Add SubSplited(0), "DLCI"
        Coll.Add SubSplited(1), "TLCI"
        
        Coll.Add FNAME, "FNAME"
        Coll.Add Now, "LUPDATE"

Set Rec1_PNRFileAddress = Coll
End Function

Public Function Rec8_AirCommission(Data)
'Some Problems with MCP- & CP-

Dim splited
Dim SubSplited, SubSplitedPreserve
Dim temp1, temp2
Dim Coll As New Collection

temp1 = Data

temp2 = ExtractBetween(temp1, "MCP-", "\")
    Coll.Add temp2, "MNYCOLAMT"


temp2 = ExtractBetween(temp1, "OCA-", "\")
    Coll.Add temp2, "OCOMAMT"
temp2 = ExtractBetween(temp1, "ACA-", "\")
    Coll.Add temp2, "ADMCOMAMT"
temp2 = ExtractBetween(temp1, "NCA-", "\")
    Coll.Add temp2, "NCOMAMT"
temp2 = ExtractBetween(temp1, "CP-", "\")
    Coll.Add temp2, "COMPERC"
temp2 = ExtractBetween(temp1, "CA-", "\")
    SubSplited = SplitWithLengthsPlus(temp2, 1)
        Coll.Add SubSplited(0), "PNORBANK"
        Coll.Add SubSplited(1), "PCOM"
        Coll.Add PassengerToAirFairID, "PassengerID"
        
        
'--------Begin Details

splited = SplitForcePlus(temp1, "\", 3)
Coll.Add splited(1), "PTYPECODE"
Coll.Add splited(2), "EXCCOMID"
'-----------------
Set Rec8_AirCommission = Coll
End Function


Public Function Rec9_TicketingCarrier(Data)
Dim splited
Dim SubSplited
Dim temp1, temp2
Dim Coll As New Collection

temp1 = Data
temp2 = ExtractBetween(temp1, "V-", "\")
    SubSplited = SplitForce(temp2, "/", 2)
        Coll.Add SubSplited(0), "VAIRCODE"
        Coll.Add SubSplited(1), "VAIRNO"
temp2 = ExtractBetween(temp1, "S-", "\")
    Coll.Add temp2, "TKTIND"
temp2 = ExtractBetween(temp1, "I-", "\")
    Coll.Add temp2, "INTIND"
temp2 = ExtractBetween(temp1, "DP-", "\")
    SubSplited = SplitForce(temp2, "/", 2)
        Coll.Add SubSplited(0), "DESTCDE"
        Coll.Add SubSplited(1), "PTCODE"
        
Set Rec9_TicketingCarrier = Coll
End Function


Public Function Rec33_TARUNKSegment(Data)
Dim splited
Dim SubSplited
Dim Coll As New Collection
splited = SplitForce(Data, "\", 3)

Coll.Add "3", "ARUNK segment code"
Coll.Add splited(1), "Itinerary segment number"
Coll.Add splited(2), "Arrival Unknown"

Set Rec33_TARUNKSegment = Coll
End Function

Public Function RecA_NameDocumentNumbers(Data)
Dim splited
Dim SubSplited, SubSplited2
Dim temp
Dim temp1, temp2
Dim Coll As New Collection
temp1 = Data

temp2 = ExtractBetween(temp1, "E-", "\")
temp2 = ExtractBetween(temp1, "D-", "\")
temp2 = ExtractBetween(temp1, "SAC-", "\")

splited = SplitForcePlus(temp1, "\", 9)

SubSplited = SplitForce(splited(2), "/", 2)
    Coll.Add SubSplited(0), "SURNAME"
    SubSplited2 = SplitFirstNameAndInitial(SubSplited(1))
    Coll.Add Replace(SubSplited2(0), "*CHD", ""), "FRSTNAME" 'Some Records Found unwanted '*CHD' which is skipped
    Coll.Add Replace(SubSplited2(1), "*CHD", ""), "PTitle"
    If UCase(splited(3)) = "ADT" Then
        temp = "Adult"
    ElseIf UCase(splited(3)) = "CNN" Then
        temp = "Child"
    Else
        temp = ""
    End If
Coll.Add temp, "PTYPE"
Coll.Add splited(4), "CUSTNO"
Coll.Add splited(5), "CUSTCMNTS"
Coll.Add splited(6), "DOCNO"
Coll.Add splited(7), "ISSUEDATE"
Coll.Add splited(8), "INVNO"
SubSplited = SplitForce(splited(9), "D-\E-\SAC-", 2)
Coll.Add SubSplited(1), "SETTMNTCDENO"

If PassengerToAirFairFlag = False Then
    PassengerToAirFairID = PassengerToAirFairID + 1
    PassengerToAirFairFlag = True
End If
Coll.Add PassengerToAirFairID, "PassengerID"

Set RecA_NameDocumentNumbers = Coll
End Function

Public Function RecD_OriginalIssueData(Data)
Dim splited
Dim SubSplited
Dim Coll As New Collection
splited = SplitForcePlus(Data, "\", 9)

SubSplited = SplitWithLengthsPlus(splited(0), 1)
    Coll.Add SubSplited(0), "Document type indicator"
    Coll.Add SubSplited(1), "Original issue document number"
Coll.Add splited(1), "Place of issue "
Coll.Add splited(2), "Original issue date"
Coll.Add splited(3), "Original issue agency IATA"
Coll.Add splited(4), "FREEDATA"
Set RecD_OriginalIssueData = Coll
End Function



Public Function RecG_AirFair(Data) As Collection
Dim splited
Dim Coll As New Collection
Dim SubSplit
splited = SplitForce(Data, "\", 7)

SubSplit = SplitWithLengthsPlus(splited(0), 3)
    Coll.Add SubSplit(0), "Currency code of fare "
    Coll.Add SubSplit(1), "Fare amount"
SubSplit = SplitWithLengthsPlus(splited(1), 3)
    Coll.Add SubSplit(0), "Tax 1 code identifier "
    Coll.Add SubSplit(1), "Tax 1 amount"

SubSplit = SplitWithLengthsPlus(splited(2), 3)
    Coll.Add SubSplit(0), "Tax 2 code identifier "
    Coll.Add SubSplit(1), "Tax 2 amount"

SubSplit = SplitWithLengthsPlus(splited(3), 3)
    Coll.Add SubSplit(0), "Tax 3 code identifier "
    Coll.Add SubSplit(1), "Tax 3 amount"

SubSplit = SplitWithLengthsPlus(splited(4), 3)
    Coll.Add SubSplit(0), "Currency code of Total Fare"
    Coll.Add SubSplit(1), "Total Fare amount"

SubSplit = SplitWithLengthsPlus(splited(5), 3)
    Coll.Add SubSplit(0), "Currency code of equivalent fare"
    Coll.Add SubSplit(1), "Equivalent fare amount "

SubSplit = SplitWithLengthsPlus(splited(6), 2)
    Coll.Add SubSplit(1), "Invoice Air Amount"

If PassengerToAirFairFlag = True Then
    PassengerToAirFairFlag = False
End If
    Coll.Add PassengerToAirFairID, "PassengerID"
    
    WorldSpan_AirFareSLNO = WorldSpan_AirFareSLNO + 1
    Coll.Add WorldSpan_AirFareSLNO, "SLNO"

Set RecG_AirFair = Coll
End Function



Public Function Rec4_4_AdditionalPTCs(Data) As Collection
Dim temp1
Dim splited
Dim SubSplited
Dim Coll As New Collection
splited = SplitForcePlus(Data, "\", 5)
Coll.Add "4", "PTCCODE"
Coll.Add splited(1), "FBASISCODE"
Coll.Add splited(2), "BDATE"
Coll.Add splited(3), "ADATE"
Coll.Add splited(4), "STATUS"

temp1 = splited(5)
SubSplited = ExtractBetween(temp1, "BG-", "\")
    Coll.Add SubSplited, "BAGGAGE"
SubSplited = ExtractBetween(temp1, "E-", "\")
    Coll.Add SubSplited, "ENDRSMNT"
SubSplited = ExtractBetween(temp1, "BP-", "\")
    Coll.Add SubSplited, "AMT"
    
Set Rec4_4_AdditionalPTCs = Coll
End Function


Public Function Rec1_HotelSegment(Data)
Dim splited
Dim SubSplited, SubSplited2
Dim temp1, temp2
Dim Coll As New Collection
temp1 = Data

'--------Starting Details are below---------




temp2 = ExtractBetween(temp1, "NP-", "\")
    Coll.Add temp2, "NOPER"
temp2 = ExtractBetween(temp1, "R-", "\")
    Coll.Add temp2, "TYPE"
    
temp2 = ExtractBetween(temp1, "RS-", "\")
    SubSplited = SplitForce(temp2, "-", 2)
        SubSplited2 = SplitFirstTwo(SubSplited(0), 3)
            Coll.Add SubSplited2(0), "CURRCODE"
            Coll.Add SubSplited2(1), "SYSRAMT"
        Coll.Add SubSplited(1), "SYSRPLAN"
temp2 = ExtractBetween(temp1, "RD-", "\")
    Coll.Add temp2, "ROOMDESC"
temp2 = ExtractBetween(temp1, "RTD-", "\")
    Coll.Add temp2, "RATEDESC"
temp2 = ExtractBetween(temp1, "RL-", "\")
    Coll.Add temp2, "ROOMLOC"
temp2 = ExtractBetween(temp1, "BS-", "\")
    Coll.Add temp2, "BSOURCE"
temp2 = ExtractBetween(temp1, "NM-", "\")
    Coll.Add temp2, "GUESTNAME"
temp2 = ExtractBetween(temp1, "CD-", "\")
    Coll.Add temp2, "DISCOUNTNO"
temp2 = ExtractBetween(temp1, "FT-", "\")
    Coll.Add temp2, "FTNO"
temp2 = ExtractBetween(temp1, "FG-", "\")
    Coll.Add temp2, "FGUESTNO"
temp2 = ExtractBetween(temp1, "TTL-", "\")
    Coll.Add temp2, "TOTALAMT"
temp2 = ExtractBetween(temp1, "BAS-", "\")
    Coll.Add temp2, "BASEAMT"
temp2 = ExtractBetween(temp1, "SVC-", "\")
    Coll.Add temp2, "SCHRGEAMT"
    
temp2 = ExtractBetween(temp1, "SUR-", "\")
    Coll.Add temp2, "SURCHARGE"
temp2 = ExtractBetween(temp1, "TTD-", "\")
    Coll.Add temp2, "TTD"
temp2 = ExtractBetween(temp1, "CM-", "\")
    Coll.Add temp2, "COMAMT"
temp2 = ExtractBetween(temp1, "CV-", "\")
    Coll.Add temp2, "VCOM"
temp2 = ExtractBetween(temp1, "CF-", "\")
    Coll.Add temp2, "CONFNO"

temp2 = ExtractBetween(temp1, "CX-", "\")
    Coll.Add temp2, "CNCLNNO"
temp2 = ExtractBetween(temp1, "TX-", "\")
    Coll.Add temp2, "TAXRATE"
temp2 = ExtractBetween(temp1, "HA1-", "\")
    Coll.Add temp2, "HOTADD1"
temp2 = ExtractBetween(temp1, "HA2-", "\")
    Coll.Add temp2, "HOTADD2"
temp2 = ExtractBetween(temp1, "SCC-", "\")
    Coll.Add temp2, "CNTRYCODE"

temp2 = ExtractBetween(temp1, "ZIP-", "\")
    Coll.Add temp2, "POSTELCODE"
temp2 = ExtractBetween(temp1, "PH-", "\")
    Coll.Add temp2, "TELENO"
temp2 = ExtractBetween(temp1, "FAX-", "\")
    Coll.Add temp2, "FAXNO"
temp2 = ExtractBetween(temp1, "CI-", "\")
    Coll.Add temp2, "CHKINTIME"
temp2 = ExtractBetween(temp1, "CO-", "\")
    Coll.Add temp2, "CHKOUTTIME"


'------Starting Details
splited = SplitForcePlus(temp2, "\", 8)
Coll.Add splited(1), "SEGNO"
Coll.Add splited(2), "CHAINCODE"
Coll.Add splited(3), "CHAINNAME"
SubSplited = SplitFirstTwo(splited(4), 2)
    Coll.Add SubSplited(0), "STATUSCODE"
    Coll.Add SubSplited(1), "ROOMS"
SubSplited = SplitWithLengths(splited(5), 3, 5, 5)
    Coll.Add SubSplited(0), "CTYCODE"
    Coll.Add SubSplited(1), "INDATE"
    Coll.Add SubSplited(2), "OUTDATE"
Coll.Add splited(6), "PROPCODE"
Coll.Add splited(7), "PROPNAME"
'----------------------

Set Rec1_HotelSegment = Coll

End Function

Public Function Rec_BranchAgentSines(Data)
Dim splited
Dim SubSplited, SubSplited2
Dim temp1, temp2
Dim Coll As New Collection
temp1 = Data

temp2 = ExtractBetween(temp1, "BL-", "\")
    SubSplited = SplitForce(temp2, "/", 2)
        SubSplited2 = SplitWithLengths(SubSplited(0), 3, 8)
            Coll.Add SubSplited2(0), "BSID"
            Coll.Add SubSplited2(1), "BIATA"
        SubSplited2 = SplitWithLengths(SubSplited(1), 7, 4, 2)
            Coll.Add SubSplited2(0), "BDATE"
            Coll.Add SubSplited2(1), "BTIME"
            Coll.Add SubSplited2(2), "BAGENT"
            
            
temp2 = ExtractBetween(temp1, "TL-", "\")
    SubSplited = SplitForce(temp2, "/", 2)
        SubSplited2 = SplitWithLengths(SubSplited(0), 3, 8)
            Coll.Add SubSplited2(1), "TIATA"
        SubSplited2 = SplitWithLengths(SubSplited(1), 7, 4, 2)
            Coll.Add ToPenDateX(SubSplited2(0)), "TDATE"
            Coll.Add (SubSplited2(2)), "TAGENT"
            
Set Rec_BranchAgentSines = Coll
End Function










Public Sub ClearTable(TableName As String)
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

Public Function SkipNegative(value, Optional default = 0)
    SkipNegative = IIf(value < 0, default, value)
End Function


Public Function To24Hour(time)
    Dim MainTime, AMPM, HourT, MinuteT, retn
    If Len(time) = 0 Then Exit Function
    AMPM = Right(time, 1)
    MainTime = Left(time, Len(time) - 1)
    MinuteT = Right(MainTime, 2)
    HourT = Left(MainTime, Len(MainTime) - 2)
    If UCase(AMPM) = "P" Then
    HourT = Val(HourT) + 12
    End If
    retn = HourT & MinuteT
    To24Hour = retn
End Function
Private Function ToPenDate(TheDate As String, Optional default = Empty) As Date
On Error GoTo errPara
Dim Day, Month, Year, TESTDATE
Day = Left(TheDate, 2)
Month = Right(TheDate, 3)
Year = mYearID
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
Public Function SplitFirstNameAndInitial(Data)
Dim aa, ub, ab, ac, AD(1)
Dim temp
aa = Split(Data, ".")
ub = UBound(aa)
If ub > 0 Then
    ab = aa(ub)
    temp = Len(Data) - (Len(ab) + 1)
    If temp > 0 Then
        ac = Left(Data, temp)
    Else
        ac = ""
    End If
Else
temp = FindInitialAndName(aa(0))
    ac = temp(0)
    ab = temp(1)
End If
AD(0) = ac
AD(1) = ab
SplitFirstNameAndInitial = AD
End Function


Public Function ToMealsServiceCode(id)
Dim ret
Select Case Trim(id)
Case "B"
    ret = "Breakfast"
Case "D"
    ret = "Dinner"
Case "L"
    ret = "Lunch"
Case "R"
    ret = "Brunch"
Case "S"
    ret = "Snack"
Case "*"
    ret = "Miscellaneous service"
Case ""
    ret = "No meal service"
End Select
ToMealsServiceCode = ret
End Function



Public Function Rec_ClientAcNo(Data)
Dim splited
Dim SubSplited
Dim temp1, temp2
Dim Coll As New Collection
temp1 = Data
Coll.Add Data, "CLNTACNO"
Set Rec_ClientAcNo = Coll
End Function

Public Function Rec_Endorsement(Data)
Dim splited
Dim SubSplited
Dim temp1, temp2
Dim Coll As New Collection
temp1 = Data
Coll.Add Data, "INFORMATION"
Set Rec_Endorsement = Coll
End Function

Public Function Rec_SSRData(Data)
Dim splited
Dim SubSplited
Dim temp1, temp2
Dim Coll As New Collection
temp1 = Data
SubSplited = SplitForce(temp1, "\", 4)

Coll.Add SubSplited(0), "SSR1"
Coll.Add SubSplited(1), "SSR2"
Coll.Add SubSplited(2), "SSR3"
Coll.Add SubSplited(3), "SSR4"

Set Rec_SSRData = Coll
End Function


Public Function Rec_PhoneContact(Data)
Dim splited
Dim SubSplited
Dim temp1, temp2
Dim Coll As New Collection
temp1 = Data
Coll.Add Data, "PHONE"
Set Rec_PhoneContact = Coll
End Function

Public Function Rec_UInput(Data)
Dim splited
Dim SubSplited
Dim temp1, temp2
Dim Coll As New Collection
temp1 = Data
temp1 = SplitWithLengthsPlus(Data, 1)
Coll.Add temp1(1), "UINPUT"
Set Rec_UInput = Coll
End Function



Public Function Rec_FormOfPayment(Data) As Collection
Dim temp1, temp2, temp3
Dim splited
Dim SubSplited
Dim Coll As New Collection
    
temp1 = Data
SubSplited = ExtractBetween(temp1, "F1-", "\")
    Coll.Add SubSplited, "FFDATA1"
SubSplited = ExtractBetween(temp1, "F2-", "\")
    Coll.Add SubSplited, "FFDATA2"
SubSplited = ExtractBetween(temp1, "F3-", "\")
    Coll.Add SubSplited, "FFDATA3"
SubSplited = ExtractBetween(temp1, "F4-", "\")
    Coll.Add SubSplited, "FFDATA4"
    
SubSplited = ExtractBetween(temp1, "ES-", "\")
    Coll.Add SubSplited, "STAXCODE"
SubSplited = ExtractBetween(temp1, "IT1-", "\")
    Coll.Add SubSplited, "ITAXCODE1"
SubSplited = ExtractBetween(temp1, "IT2-", "\")
    Coll.Add SubSplited, "ITAXCODE2"
    
splited = SplitForcePlus(temp1, "\", 3)
temp2 = SplitWithLengthsPlus(splited(1), 2)
temp3 = FormOfPayNotes(temp2(0))
    Coll.Add temp2(0), "PAYCODE"
    Coll.Add temp3, "PAYNAME"
    
    Coll.Add temp2(1), "CCDETAILS"
    
    
Set Rec_FormOfPayment = Coll
End Function


Public Function FormOfPayNotes(id) As String
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



Public Function Rec_Remarks(Data) As Collection
Dim temp1, temp
Dim splited
Dim SubSplited
Dim Coll As New Collection


Dim CollGEN As New Collection
Dim CollSellRate As New Collection
Dim CollVL As New Collection
Dim CollRef As New Collection
Dim CollRemarks As New Collection


temp1 = Data

Set CollVL = Rec_N_VLocator(temp1)

'SubSplited = ExtractBetween(temp1, "ICC-FT-SP-", "\", True)
'    Set CollSellRate = Rec_N_SellRate(SubSplited)

CollGEN.Add "", "REMARKS"

SubSplited = ExtractBetween(temp1, "ICC-BOOKED-", "\")
CollRef.Add SubSplited, "REF"

SubSplited = ExtractBetween(temp1, "TK-", "\")
    CollGEN.Add SubSplited, "TKTNO"
SubSplited = ExtractBetween(temp1, "IN-", "\")
    CollGEN.Add SubSplited, "INVNO"
    
SubSplited = ExtractBetween(temp1, "PID-", "\")
    CollGEN.Add SubSplited, "PREMARKS"
    temp = SubSplited
SubSplited = ExtractBetween(temp1, "LNS-", "\")
    CollGEN.Add SubSplited, "LREMARKS"
    temp = temp & SubSplited

CollRemarks.Add temp, "Remarks"

Coll.Add CollGEN, "GEN"
Coll.Add CollSellRate, "SEL"
Coll.Add CollVL, "VL"
Coll.Add CollRef, "REF"
Coll.Add CollRemarks, "REM"


    
Set Rec_Remarks = Coll
End Function



Public Function FindInitialAndName(mData)
On Error GoTo errPara
Dim aa, Data
Data = mData
Dim retn(1) As String

Data = Replace(Data, "*CHD", "")

aa = UCase(Right(Data, 2))
If (aa = "MR") Then
    retn(0) = Mid(Data, 1, Len(Data) - 2)
    retn(1) = "MR"
End If

aa = UCase(Right(Data, 3))
If (aa = "MRS") Then
    retn(0) = Mid(Data, 1, Len(Data) - 3)
    retn(1) = "MRS"
End If

aa = UCase(Right(Data, 4))
If (aa = "MISS") Then
    retn(0) = Mid(Data, 1, Len(Data) - 4)
    retn(1) = "MISS"
End If

aa = UCase(Right(Data, 4))
If (aa = "MSTR") Then
    retn(0) = Mid(Data, 1, Len(Data) - 4)
    retn(1) = "MSTR"
End If

aa = UCase(Right(Data, 4))
If (aa = "PROF") Then
    retn(0) = Mid(Data, 1, Len(Data) - 4)
    retn(1) = "PROF"
End If



If retn(0) = "" Then
    retn(0) = Data
    retn(1) = ""
End If

FindInitialAndName = retn
Exit Function
errPara:

If retn(0) = "" Then
    retn(0) = Data
    retn(1) = ""
End If
End Function

Public Function Rec_N_SellRate(mData)
Dim splited, SubSplited
Dim Coll As New Collection, coll2 As New Collection
splited = SplitWithLengths(mData, 10, 10)


If Len(mData) = 10 Then
    SubSplited = SplitWithLengths(splited(0), 3, 4, 3)
        Coll.Add SubSplited(0), "CURCODE"
        Coll.Add SubSplited(1), "AMOUNT"
        Coll.Add SubSplited(2), "TYPE"
    coll2.Add Coll
    SubSplited = SplitWithLengths(splited(1), 3, 4, 3)
        Set Coll = New Collection
        Coll.Add SubSplited(0), "CURCODE"
        Coll.Add SubSplited(1), "AMOUNT"
        Coll.Add SubSplited(2), "TYPE"
    coll2.Add Coll
ElseIf (Val(mData) > 0 And CStr(Val(mData)) = mData) Then
        Coll.Add "GBP", "CURCODE"
        Coll.Add mData, "AMOUNT"
        Coll.Add "ADT", "TYPE"
    coll2.Add Coll
    SubSplited = SplitWithLengths(splited(1), 3, 4, 3)
        Set Coll = New Collection
        Coll.Add "GBP", "CURCODE"
        Coll.Add "0", "AMOUNT"
        Coll.Add "CHD", "TYPE"
    coll2.Add Coll
Else
    

End If
End Function

Public Function Rec_N_VLocator(mData)
    Dim splited, SubSplited, temp, Res, Counter
    Dim Coll As New Collection, coll2 As New Collection
    temp = mData
    Counter = 1
    Res = " "
While Len(Res) > 0
    Res = ExtractBetween(temp, "- " & Counter & " ", "\", True)
    If Len(Res) > 0 Then
        Set Coll = New Collection
        Coll.Add Left(Res, 2), "AIRCODE"
        Coll.Add Right(Res, 6), "VLOCATOR"
        Coll.Add Counter, "SLNO"
        coll2.Add Coll
    End If
    Counter = Counter + 1
Wend
    Set Rec_N_VLocator = coll2
End Function
Public Function Rec_N_Referance(mData)
    Dim splited, SubSplited
    Dim Coll As New Collection, coll2 As New Collection
    Coll.Add mData, "REF"
End Function

Public Function Rec_N_Remarks(mData)
    Dim splited, SubSplited
    Dim Coll As New Collection
    Coll.Add mData, "Remarks"
End Function



Public Function InsertDataCollectionINDEX(Coll As Collection, TableNae As String, LineID, UploadNo As Long)
For j = 1 To Coll.Count
    InsertData Coll(j), TableNae, CStr(LineID), UploadNo
Next
End Function
Public Function InsertDataCollectionKey(Coll As Collection, TableNae As String, LineID, UploadNo As Long)
For j = 1 To Coll.Count
    InsertDataByFieldName Coll(j), TableNae, CStr(LineID), UploadNo
Next
End Function


Public Function TryParseSellRate1(mData, mColl As Collection) As Boolean
Dim Coll As New Collection
Dim coll2 As New Collection
Dim splited, SubSplited
If Len(mData) <> 10 Then TryParseSellRate1 = False: Exit Function
    SubSplited = SplitWithLengths(splited(0), 3, 4, 3)
        Coll.Add SubSplited(0), "CURCODE"
        Coll.Add SubSplited(1), "AMOUNT"
        If Val(SubSplited(1)) = 0 Then TryParseSellRate1 = False: Exit Function
        Coll.Add SubSplited(2), "TYPE"
    coll2.Add Coll
    SubSplited = SplitWithLengths(splited(1), 3, 4, 3)
        Set Coll = New Collection
        Coll.Add SubSplited(0), "CURCODE"
        Coll.Add SubSplited(1), "AMOUNT"
        If Val(SubSplited(1)) = 0 Then TryParseSellRate1 = False: Exit Function
        Coll.Add SubSplited(2), "TYPE"
    coll2.Add Coll

End Function
