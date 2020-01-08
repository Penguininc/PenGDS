Attribute VB_Name = "GDSINTRAYModule"
'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
Option Explicit

'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
Public Type GDSINTRAYTABLE_Type
    GIT_AUTONO_Long As Long
    GIT_ID_String As String
    GIT_GDS_String As String
    GIT_INTERFACE_String As String
    GIT_UPLOADNO_Long As Long
    GIT_PNRDATE_String As String
    GIT_PNR_String As String
    GIT_LASTNAME_String As String
    GIT_FIRSTNAME_String As String
    GIT_BOOKINGAGENT_String As String
    GIT_TICKETNUMBER_String As String
    GIT_FILENAME_String As String
    GIT_GDSAUTOFAILED_Integer As Integer
    GIT_PCC_String As String
End Type

'by Abhi on 24-Sep-2013 for caseid 3394 GDS In Tray optimisation header table
Public Function GDSINTRAYTABLE_Add(ByRef pGDSINTRAYTABLE_Type As GDSINTRAYTABLE_Type) As Boolean
Dim vSQL_String As String
    
'    vSQL_String = "" _
'        & "INSERT INTO VATCASHTABLE (" _
'            & "  GIT_ID          , GIT_GDS     , GIT_INTERFACE    , GIT_UPLOADNO  " _
'            & ", GIT_PNRDATE     , GIT_PNR     , GIT_LASTNAME     , GIT_FIRSTNAME " _
'            & ", GIT_TICKETNUMBER, GIT_FILENAME, GIT_GDSAUTOFAILED                " _
'            & ""
'        vSQL_String = vSQL_String & ") VALUES ("
'        vSQL_String = vSQL_String & "" _
'            & "  '" & SkipChars(pGDSINTRAYTABLE_Type.GIT_ID_String) & "', '" & SkipChars(pGDSINTRAYTABLE_Type.GIT_GDS_String) & "', '" & SkipChars(pGDSINTRAYTABLE_Type.GIT_INTERFACE_String) & "', '" & SkipChars(pGDSINTRAYTABLE_Type.GIT_UPLOADNO_Long) & "' " _
'            & ", '" & SkipChars(DateFormatBlankto1900(pGDSINTRAYTABLE_Type.GIT_PNRDATE_String)) & "', '" & SkipChars(pGDSINTRAYTABLE_Type.GIT_PNR_String) & "', '" & SkipChars(pGDSINTRAYTABLE_Type.GIT_LASTNAME_String) & "', '" & SkipChars(pGDSINTRAYTABLE_Type.GIT_FIRSTNAME_String) & "' " _
'            & ", '" & SkipChars(pGDSINTRAYTABLE_Type.GIT_TICKETNUMBER_String) & "', '" & SkipChars(pGDSINTRAYTABLE_Type.GIT_FILENAME_String) & "', '" & SkipChars(pGDSINTRAYTABLE_Type.GIT_GDSAUTOFAILED_Integer) & "' " _
'            & ")"
    dbCompany.Execute vSQL_String
GDSINTRAYTABLE_Add = True
End Function


