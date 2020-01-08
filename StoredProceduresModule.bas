Attribute VB_Name = "StoredProceduresModule"
Option Explicit

'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date
Public Enum MonthEndClosingLedgerType
    MonthEndClosingLedgerTypeGeneral = 1
    MonthEndClosingLedgerTypeCustomer = 2
    MonthEndClosingLedgerTypeSupplier = 3
End Enum
'by Abhi on 01-Nov-2018 for caseid 9385 Amadeus -GDS File Upload Out Of Bookings Date

'by Abhi on 12-Feb-2012 for caseid 1606 Folder PNR mixing first free number
Public Function FirstFreeNumber(ByVal pType_String As String) As Long
Dim vFirstFreeNumber_Long As Long
Dim vCommand As New ADODB.Command

    vCommand.ActiveConnection = dbCompany
    vCommand.CommandTimeout = dbCompany.CommandTimeout
    vCommand.CommandType = adCmdStoredProc
    vCommand.CommandText = "dbo.FirstFreeNumberStoredProcedure"
    vCommand.Parameters.Append vCommand.CreateParameter("pFFN_TYPE_nvarchar", adVarChar, adParamInput, 100, pType_String)
    vCommand.Parameters.Append vCommand.CreateParameter("vFirstFreeNumber_bigint", adBigInt, adParamOutput)
    vCommand.Execute
    vFirstFreeNumber_Long = vCommand("vFirstFreeNumber_bigint")
    Set vCommand.ActiveConnection = Nothing

FirstFreeNumber = vFirstFreeNumber_Long
End Function

