Attribute VB_Name = "LegacyModule"
'by Abhi on 21-Jul-2012 for caseid 2401 Travcom data transfer
Option Explicit

Public Type WSPPNRADD_Type
    UpLoadNo_Long As Long
    RecID_String As String
    INTLVL_String As String
    PNRADD_String As String
    FINVNO_String As String
    LINVNO_String As String
    ITNRYCHNGE_String As String
    DLCI_String As String
    TLCI_String As String
    FNAME_String As String
    LUPDATE_String As String
    GDSAutoFailed_Integer As Integer
End Type

Public Type WSPAGNTSINE_Type
    UpLoadNo_Long As Long
    RecID_String As String
    BSID_String As String
    BIATA_String As String
    BDATE_String As String
    BTIME_String As String
    BAGENT_String As String
    TIATA_String As String
    TDATE_String As String
    TAGENT_String As String
End Type

Public Type WSPCLNTACNO_Type
    UpLoadNo_Long As Long
    RecID_String As String
    CLNTACNO_String As String
End Type

Public Type WSPFORMOFPYMNT_Type
    UpLoadNo_Long As Long
    RecID_String As String
    PAYCODE_String As String
    PAYNAME_String As String
    FFDATA1_String As String
    FFDATA2_String As String
    FFDATA3_String As String
    FFDATA4_String As String
    STAXCODE_String As String
    ITAXCODE1_String As String
    ITAXCODE2_String As String
    CCDETAILS_String As String
End Type

Public Type WSPTCARRIER_Type
    UpLoadNo_Long As Long
    RecID_String As String
    VAIRCODE_String As String
    VAIRNO_String As String
    TKTIND_String As String
    INTIND_String As String
    DESTCDE_String As String
    PTCODE_String As String
End Type

Public Type WSPTKTSEG_Type
    UpLoadNo_Long As Long
    RecID_String As String
    AIRSEGCODE_String As String
    ITNRYSEGNO_String As String
    NOINTSTOP_String As String
    SHDESIGNINDI_String As String
    AIRCODE_String As String
    FLTNO_String As String
    Class_String As String
    ORGAIRCODE_String As String
    DEPDATE_String As String
    DEPTIME_String As String
    DESTAIRCODE_String As String
    ARRDATE_String As String
    ARRTIME_String As String
    BDATE_String As String
    ADATE_String As String
    BAGG_String As String
    STATUS_String As String
    MEALSERVCODE_String As String
    SEGSTOPCODE_String As String
    EQUIPTYPE_String As String
    FBASISCODE_String As String
    SEGMILEAGE_String As String
    INTSTOP_String As String
    FLTTIME_String As String
    SEGOVRRDIND_String As String
    DEPTRMNLCODE_String As String
    ARRTRMNLCODE_String As String
    BPAYAMT_String As String
End Type

Public Type WSPPNAME_Type
    UpLoadNo_Long As Long
    RecID_String As String
    SURNAME_String As String
    FRSTNAME_String As String
    PTITLE_String As String
    pType_String As String
    CUSTNO_String As String
    CUSTCMNTS_String As String
    DOCNO_String As String
    ISSUEDATE_String As String
    INVNO_String As String
    SETTMNTCDENO_String As String
    PassengerID_Integer As Integer
    ETKTIND_String As String
    SURNAMEFRSTNAMENO_String As String
End Type

Public Type WSPAIRFARE_Type
    UpLoadNo_Long As Long
    RecID_String As String
    CURCODEFARE_String As String
    FARE_String As String
    TAX1CODEID_String As String
    TAX1AMT_String As String
    TAX2CODEID_String As String
    TAX2AMT_String As String
    TAX3CODEID_String As String
    TAX3AMT_String As String
    CURRCDETFARE_String As String
    TATFAMT_String As String
    CURRCDEEFARE_String As String
    EFAMT_String As String
    INVAMT_String As String
    PassengerID_Integer As Integer
    SLNO_Integer As Integer
End Type

Public Type WSPN_REF_Type
    UpLoadNo_Long As Long
    RecID_String As String
    REF_String As String
End Type

Public Type WSPPENLINE_Type
    UpLoadNo_Long As Long
    RecID_String As String
    AC_String As String
    DES_String As String
    REFE_String As String
    SELLADT_String As String
    SELLCHD_String As String
    SCHARGE_String As String
    DEPT_String As String
    BRANCH_String As String
    CUSTCC_String As String
    PEMAIL_String As String
    PTELE_String As String
    PNO_Long As Long
    FARE_Currency As Currency
    TAXES_Currency As Currency
    MARKUP_Currency As Currency
    SURNAMEFRSTNAMENO_String As String
    PENFARESELL_Currency As Currency
    PENFAREPASSTYPE_String As String
    DeliAdd_String As String
    MC_String As String
    BB_String As String
    PENFARESUPPID_String As String
    PENFAREAIRID_String As String
    PENLINKPNR_String As String
    PENOPRDID_String As String
    PENOQTY_Integer As Integer
    PENORATE_Currency As Currency
    PENOSELL_Currency As Currency
    PENOSUPPID_String As String
    PENATOLTYPE_String As String
    INETREF_String As String
    PENOPAYMETHODID_String As String
    PENRT_String As String
    PENPOL_String As String
    PENPROJ_String As String
    PENEID_String As String
    PENPO_String As String
    PENHFRC_String As String
    PENLFRC_String As String
    PENHIGHF_Currency As Currency
    PENLOWF_Currency As Currency
    PENUC1_String As String
    PENUC2_String As String
    PENUC3_String As String
End Type

Public Type WSPHTLSEG_Type
    UpLoadNo_Long As Long
    RecID_String As String
    SEGNO_String As String
    CHAINCODE_String As String
    CHAINNAME_String As String
    STATUSCODE_String As String
    ROOMS_String As String
    CTYCODE_String As String
    INDATE_String As String
    OUTDATE_String As String
    PROPCODE_String As String
    PROPNAME_String As String
    NOPER_String As String
    TYPE_String As String
    CURRCODE_String As String
    SYSRAMT_String As String
    SYSRPLAN_String As String
    ROOMDESC_String As String
    RATEDESC_String As String
    ROOMLOC_String As String
    BSOURCE_String As String
    GUESTNAME_String As String
    DISCOUNTNO_String As String
    FTNO_String As String
    FGUESTNO_String As String
    TOTALAMT_String As String
    BASEAMT_String As String
    SCHRGEAMT_String As String
    SURCHARGE_String As String
    TTD_String As String
    COMAMT_String As String
    VCOM_String As String
    CONFNO_String As String
    CNCLNNO_String As String
    TAXRATE_String As String
    HOTADD1_String As String
    HOTADD2_String As String
    CNTRYCODE_String As String
    POSTELCODE_String As String
    TELENO_String As String
    FAXNO_String As String
    CHKINTIME_String As String
    CHKOUTTIME_String As String
    RGCURRCODE_String As String
    RGAMT_String As String
    EXRATEDTE_String As String
    CURRCODELOC_String As String
    CNTRYCODEH_String As String
    EXCHRATE_String As String
    RT_String As String
    RTCURRCODE_String As String
    RTAMT_String As String
    RQ_String As String
    RQCURRCODE_String As String
    RQAMT_String As String
    RQ1_String As String
    RQ1CURRCODE_String As String
    RQ1AMT_String As String
    RG1_String As String
    RG1CURRCODE_String As String
    RG1AMT_String As String
    RT1_String As String
    RT1CURRCODE_String As String
    RT1AMT_String As String
End Type

Public Function WSPPNRADD_Add(ByRef pWSPPNRADD_Type As WSPPNRADD_Type) As Boolean
Dim vSQL_String As String
    
    vSQL_String = "" _
        & "INSERT INTO WSPPNRADD (" _
            & "UpLoadNo,RecID,INTLVL,PNRADD" _
            & ",FINVNO,LINVNO,ITNRYCHNGE,DLCI" _
            & ",TLCI,FNAME,LUPDATE,GDSAutoFailed"
        vSQL_String = vSQL_String & ") VALUES ("
        vSQL_String = vSQL_String & "" _
            & "" & pWSPPNRADD_Type.UpLoadNo_Long & ",'" & SkipChars(pWSPPNRADD_Type.RecID_String) & "','" & SkipChars(pWSPPNRADD_Type.INTLVL_String) & "','" & SkipChars(pWSPPNRADD_Type.PNRADD_String) & "'" _
            & ",'" & SkipChars(pWSPPNRADD_Type.FINVNO_String) & "','" & SkipChars(pWSPPNRADD_Type.LINVNO_String) & "','" & SkipChars(pWSPPNRADD_Type.ITNRYCHNGE_String) & "','" & SkipChars(pWSPPNRADD_Type.DLCI_String) & "'" _
            & ",'" & SkipChars(pWSPPNRADD_Type.TLCI_String) & "','" & SkipChars(pWSPPNRADD_Type.FNAME_String) & "','" & SkipChars(pWSPPNRADD_Type.LUPDATE_String) & "'," & SkipChars(pWSPPNRADD_Type.GDSAutoFailed_Integer) & "" _
            & ")"
    dbCompany.Execute vSQL_String

WSPPNRADD_Add = True
End Function

Public Function WSPAGNTSINE_Add(ByRef pWSPAGNTSINE_Type As WSPAGNTSINE_Type) As Boolean
Dim vSQL_String As String
    
    vSQL_String = "" _
        & "INSERT INTO WSPAGNTSINE (" _
            & "UpLoadNo,RecID,BSID,BIATA" _
            & ",BDATE,BTIME,BAGENT,TIATA" _
            & ",TDATE,TAGENT"
        vSQL_String = vSQL_String & ") VALUES ("
        vSQL_String = vSQL_String & "" _
            & "" & pWSPAGNTSINE_Type.UpLoadNo_Long & ",'" & SkipChars(pWSPAGNTSINE_Type.RecID_String) & "','" & SkipChars(pWSPAGNTSINE_Type.BSID_String) & "','" & SkipChars(pWSPAGNTSINE_Type.BIATA_String) & "'" _
            & ",'" & SkipChars(pWSPAGNTSINE_Type.BDATE_String) & "','" & SkipChars(pWSPAGNTSINE_Type.BTIME_String) & "','" & SkipChars(pWSPAGNTSINE_Type.BAGENT_String) & "','" & SkipChars(pWSPAGNTSINE_Type.TIATA_String) & "'" _
            & ",'" & SkipChars(pWSPAGNTSINE_Type.TDATE_String) & "','" & SkipChars(pWSPAGNTSINE_Type.TAGENT_String) & "'" _
            & ")"
    dbCompany.Execute vSQL_String

WSPAGNTSINE_Add = True
End Function

Public Function WSPCLNTACNO_Add(ByRef pWSPCLNTACNO_Type As WSPCLNTACNO_Type) As Boolean
Dim vSQL_String As String
    
    vSQL_String = "" _
        & "INSERT INTO WSPCLNTACNO (" _
            & "UploadNo,RecID,CLNTACNO"
        vSQL_String = vSQL_String & ") VALUES ("
        vSQL_String = vSQL_String & "" _
            & "" & pWSPCLNTACNO_Type.UpLoadNo_Long & ",'" & SkipChars(pWSPCLNTACNO_Type.RecID_String) & "','" & SkipChars(pWSPCLNTACNO_Type.CLNTACNO_String) & "'" _
            & ")"
    dbCompany.Execute vSQL_String

WSPCLNTACNO_Add = True
End Function

Public Function WSPFORMOFPYMNT_Add(ByRef pWSPFORMOFPYMNT_Type As WSPFORMOFPYMNT_Type) As Boolean
Dim vSQL_String As String
    
    vSQL_String = "" _
        & "INSERT INTO WSPFORMOFPYMNT (" _
            & "UpLoadNo,RecID,PAYCODE,PAYNAME" _
            & ",FFDATA1,FFDATA2,FFDATA3,FFDATA4" _
            & ",STAXCODE,ITAXCODE1,ITAXCODE2,CCDETAILS"
        vSQL_String = vSQL_String & ") VALUES ("
        vSQL_String = vSQL_String & "" _
            & "" & pWSPFORMOFPYMNT_Type.UpLoadNo_Long & ",'" & SkipChars(pWSPFORMOFPYMNT_Type.RecID_String) & "','" & SkipChars(pWSPFORMOFPYMNT_Type.PAYCODE_String) & "','" & SkipChars(pWSPFORMOFPYMNT_Type.PAYNAME_String) & "'" _
            & ",'" & pWSPFORMOFPYMNT_Type.FFDATA1_String & "','" & SkipChars(pWSPFORMOFPYMNT_Type.FFDATA2_String) & "','" & SkipChars(pWSPFORMOFPYMNT_Type.FFDATA3_String) & "','" & SkipChars(pWSPFORMOFPYMNT_Type.FFDATA4_String) & "'" _
            & ",'" & SkipChars(pWSPFORMOFPYMNT_Type.STAXCODE_String) & "','" & SkipChars(pWSPFORMOFPYMNT_Type.ITAXCODE1_String) & "','" & SkipChars(pWSPFORMOFPYMNT_Type.ITAXCODE2_String) & "','" & SkipChars(pWSPFORMOFPYMNT_Type.CCDETAILS_String) & "'" _
            & ")"
    dbCompany.Execute vSQL_String

WSPFORMOFPYMNT_Add = True
End Function

Public Function WSPTCARRIER_Add(ByRef pWSPTCARRIER_Type As WSPTCARRIER_Type) As Boolean
Dim vSQL_String As String
    
    vSQL_String = "" _
        & "INSERT INTO WSPTCARRIER (" _
            & "UpLoadNo,RecID,VAIRCODE,VAIRNO" _
            & ",TKTIND,INTIND,DESTCDE,PTCODE"
        vSQL_String = vSQL_String & ") VALUES ("
        vSQL_String = vSQL_String & "" _
            & "" & pWSPTCARRIER_Type.UpLoadNo_Long & ",'" & SkipChars(pWSPTCARRIER_Type.RecID_String) & "','" & SkipChars(pWSPTCARRIER_Type.VAIRCODE_String) & "','" & SkipChars(pWSPTCARRIER_Type.VAIRNO_String) & "'" _
            & ",'" & SkipChars(pWSPTCARRIER_Type.TKTIND_String) & "','" & SkipChars(pWSPTCARRIER_Type.INTIND_String) & "','" & SkipChars(pWSPTCARRIER_Type.DESTCDE_String) & "','" & SkipChars(pWSPTCARRIER_Type.PTCODE_String) & "'" _
            & ")"
    dbCompany.Execute vSQL_String

WSPTCARRIER_Add = True
End Function

Public Function WSPTKTSEG_Add(ByRef pWSPTKTSEG_Type As WSPTKTSEG_Type) As Boolean
Dim vSQL_String As String
    
    vSQL_String = "" _
        & "INSERT INTO WSPTKTSEG (" _
            & "UpLoadNo,RecID,AIRSEGCODE,ITNRYSEGNO" _
            & ",NOINTSTOP,SHDESIGNINDI,AIRCODE,FLTNO" _
            & ",CLASS,ORGAIRCODE,DEPDATE,DEPTIME" _
            & ",DESTAIRCODE,ARRDATE,ARRTIME,BDATE" _
            & ",ADATE,BAGG,STATUS,MEALSERVCODE" _
            & ",SEGSTOPCODE,EQUIPTYPE,FBASISCODE,SEGMILEAGE" _
            & ",INTSTOP,FLTTIME,SEGOVRRDIND,DEPTRMNLCODE" _
            & ",ARRTRMNLCODE,BPAYAMT"
        vSQL_String = vSQL_String & ") VALUES ("
        vSQL_String = vSQL_String & "" _
            & "" & pWSPTKTSEG_Type.UpLoadNo_Long & ",'" & SkipChars(pWSPTKTSEG_Type.RecID_String) & "','" & SkipChars(pWSPTKTSEG_Type.AIRSEGCODE_String) & "','" & SkipChars(pWSPTKTSEG_Type.ITNRYSEGNO_String) & "'" _
            & ",'" & SkipChars(pWSPTKTSEG_Type.NOINTSTOP_String) & "','" & SkipChars(pWSPTKTSEG_Type.SHDESIGNINDI_String) & "','" & SkipChars(pWSPTKTSEG_Type.AIRCODE_String) & "','" & SkipChars(pWSPTKTSEG_Type.FLTNO_String) & "'" _
            & ",'" & SkipChars(pWSPTKTSEG_Type.Class_String) & "','" & SkipChars(pWSPTKTSEG_Type.ORGAIRCODE_String) & "','" & SkipChars(pWSPTKTSEG_Type.DEPDATE_String) & "','" & SkipChars(pWSPTKTSEG_Type.DEPTIME_String) & "'" _
            & ",'" & SkipChars(pWSPTKTSEG_Type.DESTAIRCODE_String) & "','" & SkipChars(pWSPTKTSEG_Type.ARRDATE_String) & "','" & SkipChars(pWSPTKTSEG_Type.ARRTIME_String) & "','" & SkipChars(pWSPTKTSEG_Type.BDATE_String) & "'" _
            & ",'" & SkipChars(pWSPTKTSEG_Type.ADATE_String) & "','" & SkipChars(pWSPTKTSEG_Type.BAGG_String) & "','" & SkipChars(pWSPTKTSEG_Type.STATUS_String) & "','" & SkipChars(pWSPTKTSEG_Type.MEALSERVCODE_String) & "'" _
            & ",'" & SkipChars(pWSPTKTSEG_Type.SEGSTOPCODE_String) & "','" & SkipChars(pWSPTKTSEG_Type.EQUIPTYPE_String) & "','" & SkipChars(pWSPTKTSEG_Type.FBASISCODE_String) & "','" & SkipChars(pWSPTKTSEG_Type.SEGMILEAGE_String) & "'" _
            & ",'" & SkipChars(pWSPTKTSEG_Type.INTSTOP_String) & "','" & SkipChars(pWSPTKTSEG_Type.FLTTIME_String) & "','" & SkipChars(pWSPTKTSEG_Type.SEGOVRRDIND_String) & "','" & SkipChars(pWSPTKTSEG_Type.DEPTRMNLCODE_String) & "'" _
            & ",'" & SkipChars(pWSPTKTSEG_Type.ARRTRMNLCODE_String) & "','" & SkipChars(pWSPTKTSEG_Type.BPAYAMT_String) & "'" _
            & ")"
    dbCompany.Execute vSQL_String

WSPTKTSEG_Add = True
End Function

Public Function WSPPNAME_Add(ByRef pWSPPNAME_Type As WSPPNAME_Type) As Boolean
Dim vSQL_String As String
    
    vSQL_String = "" _
        & "INSERT INTO WSPPNAME (" _
            & "UpLoadNo,RecID,SURNAME,FRSTNAME" _
            & ",PTITLE,PTYPE,CUSTNO,CUSTCMNTS" _
            & ",DOCNO,ISSUEDATE,INVNO,SETTMNTCDENO" _
            & ",PassengerID,ETKTIND,SURNAMEFRSTNAMENO"
        vSQL_String = vSQL_String & ") VALUES ("
        vSQL_String = vSQL_String & "" _
            & "" & pWSPPNAME_Type.UpLoadNo_Long & ",'" & SkipChars(pWSPPNAME_Type.RecID_String) & "','" & SkipChars(pWSPPNAME_Type.SURNAME_String) & "','" & SkipChars(pWSPPNAME_Type.FRSTNAME_String) & "'" _
            & ",'" & SkipChars(pWSPPNAME_Type.PTITLE_String) & "','" & SkipChars(pWSPPNAME_Type.pType_String) & "','" & SkipChars(pWSPPNAME_Type.CUSTNO_String) & "','" & SkipChars(pWSPPNAME_Type.CUSTCMNTS_String) & "'" _
            & ",'" & SkipChars(pWSPPNAME_Type.DOCNO_String) & "','" & SkipChars(pWSPPNAME_Type.ISSUEDATE_String) & "','" & SkipChars(pWSPPNAME_Type.INVNO_String) & "','" & SkipChars(pWSPPNAME_Type.SETTMNTCDENO_String) & "'" _
            & "," & SkipChars(pWSPPNAME_Type.PassengerID_Integer) & ",'" & SkipChars(pWSPPNAME_Type.ETKTIND_String) & "','" & SkipChars(pWSPPNAME_Type.SURNAMEFRSTNAMENO_String) & "'" _
            & ")"
    dbCompany.Execute vSQL_String

WSPPNAME_Add = True
End Function

Public Function WSPAIRFARE_Add(ByRef pWSPAIRFARE_Type As WSPAIRFARE_Type) As Boolean
Dim vSQL_String As String
    
    vSQL_String = "" _
        & "INSERT INTO WSPAIRFARE (" _
            & "UpLoadNo,RecID,CURCODEFARE,FARE" _
            & ",TAX1CODEID,TAX1AMT,TAX2CODEID,TAX2AMT" _
            & ",TAX3CODEID,TAX3AMT,CURRCDETFARE,TATFAMT" _
            & ",CURRCDEEFARE,EFAMT,INVAMT,PassengerID" _
            & ",SLNO"
        vSQL_String = vSQL_String & ") VALUES ("
        vSQL_String = vSQL_String & "" _
            & "" & pWSPAIRFARE_Type.UpLoadNo_Long & ",'" & SkipChars(pWSPAIRFARE_Type.RecID_String) & "','" & SkipChars(pWSPAIRFARE_Type.CURCODEFARE_String) & "','" & SkipChars(pWSPAIRFARE_Type.FARE_String) & "'" _
            & ",'" & SkipChars(pWSPAIRFARE_Type.TAX1CODEID_String) & "','" & SkipChars(pWSPAIRFARE_Type.TAX1AMT_String) & "','" & SkipChars(pWSPAIRFARE_Type.TAX2CODEID_String) & "','" & SkipChars(pWSPAIRFARE_Type.TAX2AMT_String) & "'" _
            & ",'" & SkipChars(pWSPAIRFARE_Type.TAX3CODEID_String) & "','" & SkipChars(pWSPAIRFARE_Type.TAX3AMT_String) & "','" & SkipChars(pWSPAIRFARE_Type.CURRCDETFARE_String) & "','" & SkipChars(pWSPAIRFARE_Type.TATFAMT_String) & "'" _
            & ",'" & SkipChars(pWSPAIRFARE_Type.CURRCDEEFARE_String) & "','" & SkipChars(pWSPAIRFARE_Type.EFAMT_String) & "','" & SkipChars(pWSPAIRFARE_Type.INVAMT_String) & "'," & pWSPAIRFARE_Type.PassengerID_Integer & "" _
            & "," & pWSPAIRFARE_Type.SLNO_Integer & "" _
            & ")"
    dbCompany.Execute vSQL_String

WSPAIRFARE_Add = True
End Function

Public Function WSPN_REF_Add(ByRef pWSPN_REF_Type As WSPN_REF_Type) As Boolean
Dim vSQL_String As String
    
    vSQL_String = "" _
        & "INSERT INTO WSPN_REF (" _
            & "UpLoadNO,RecID,REF"
        vSQL_String = vSQL_String & ") VALUES ("
        vSQL_String = vSQL_String & "" _
            & "" & pWSPN_REF_Type.UpLoadNo_Long & ",'" & SkipChars(pWSPN_REF_Type.RecID_String) & "','" & SkipChars(pWSPN_REF_Type.REF_String) & "'" _
            & ")"
    dbCompany.Execute vSQL_String

WSPN_REF_Add = True
End Function

Public Function WSPPENLINE_Add(ByRef pWSPPENLINE_Type As WSPPENLINE_Type) As Boolean
Dim vSQL_String As String
    
    vSQL_String = "" _
        & "INSERT INTO WSPPENLINE (" _
            & "UpLoadNo,RecID,AC,DES" _
            & ",REFE,SELLADT,SELLCHD,SCHARGE" _
            & ",DEPT,BRANCH,CUSTCC,PEMAIL" _
            & ",PTELE,PNO,FARE,TAXES" _
            & ",MARKUP,SURNAMEFRSTNAMENO,PENFARESELL,PENFAREPASSTYPE" _
            & ",DeliAdd,MC,BB,PENFARESUPPID" _
            & ",PENFAREAIRID,PENLINKPNR,PENOPRDID,PENOQTY" _
            & ",PENORATE,PENOSELL,PENOSUPPID,PENATOLTYPE" _
            & ",INETREF,PENOPAYMETHODID,PENRT,PENPOL" _
            & ",PENPROJ,PENEID,PENPO,PENHFRC" _
            & ",PENLFRC,PENHIGHF,PENLOWF,PENUC1" _
            & ",PENUC2,PENUC3"
        vSQL_String = vSQL_String & ") VALUES ("
        vSQL_String = vSQL_String & "" _
            & "" & pWSPPENLINE_Type.UpLoadNo_Long & ",'" & SkipChars(pWSPPENLINE_Type.RecID_String) & "','" & SkipChars(pWSPPENLINE_Type.AC_String) & "','" & SkipChars(pWSPPENLINE_Type.DES_String) & "'" _
            & ",'" & SkipChars(pWSPPENLINE_Type.REFE_String) & "','" & SkipChars(pWSPPENLINE_Type.SELLADT_String) & "','" & SkipChars(pWSPPENLINE_Type.SELLCHD_String) & "','" & SkipChars(pWSPPENLINE_Type.SCHARGE_String) & "'" _
            & ",'" & SkipChars(pWSPPENLINE_Type.DEPT_String) & "','" & SkipChars(pWSPPENLINE_Type.BRANCH_String) & "','" & SkipChars(pWSPPENLINE_Type.CUSTCC_String) & "','" & SkipChars(pWSPPENLINE_Type.PEMAIL_String) & "'" _
            & ",'" & SkipChars(pWSPPENLINE_Type.PTELE_String) & "'," & pWSPPENLINE_Type.PNO_Long & "," & pWSPPENLINE_Type.FARE_Currency & "," & pWSPPENLINE_Type.TAXES_Currency & "" _
            & "," & pWSPPENLINE_Type.MARKUP_Currency & ",'" & SkipChars(pWSPPENLINE_Type.SURNAMEFRSTNAMENO_String) & "'," & pWSPPENLINE_Type.PENFARESELL_Currency & ",'" & SkipChars(pWSPPENLINE_Type.PENFAREPASSTYPE_String) & "'" _
            & ",'" & SkipChars(pWSPPENLINE_Type.DeliAdd_String) & "','" & SkipChars(pWSPPENLINE_Type.MC_String) & "','" & SkipChars(pWSPPENLINE_Type.BB_String) & "','" & SkipChars(pWSPPENLINE_Type.PENFARESUPPID_String) & "'" _
            & ",'" & SkipChars(pWSPPENLINE_Type.PENFAREAIRID_String) & "','" & SkipChars(pWSPPENLINE_Type.PENLINKPNR_String) & "','" & SkipChars(pWSPPENLINE_Type.PENOPRDID_String) & "'," & pWSPPENLINE_Type.PENOQTY_Integer & "" _
            & "," & pWSPPENLINE_Type.PENORATE_Currency & "," & pWSPPENLINE_Type.PENOSELL_Currency & ",'" & SkipChars(pWSPPENLINE_Type.PENOSUPPID_String) & "','" & SkipChars(pWSPPENLINE_Type.PENATOLTYPE_String) & "'" _
            & ",'" & SkipChars(pWSPPENLINE_Type.INETREF_String) & "','" & SkipChars(pWSPPENLINE_Type.PENOPAYMETHODID_String) & "','" & SkipChars(pWSPPENLINE_Type.PENRT_String) & "','" & SkipChars(pWSPPENLINE_Type.PENPOL_String) & "'" _
            & ",'" & SkipChars(pWSPPENLINE_Type.PENPROJ_String) & "','" & SkipChars(pWSPPENLINE_Type.PENEID_String) & "','" & SkipChars(pWSPPENLINE_Type.PENPO_String) & "','" & SkipChars(pWSPPENLINE_Type.PENHFRC_String) & "'" _
            & ",'" & SkipChars(pWSPPENLINE_Type.PENLFRC_String) & "'," & pWSPPENLINE_Type.PENHIGHF_Currency & "," & pWSPPENLINE_Type.PENLOWF_Currency & ",'" & SkipChars(pWSPPENLINE_Type.PENUC1_String) & "'" _
            & ",'" & SkipChars(pWSPPENLINE_Type.PENUC2_String) & "','" & SkipChars(pWSPPENLINE_Type.PENUC3_String) & "'" _
            & ")"
    dbCompany.Execute vSQL_String

WSPPENLINE_Add = True
End Function

Public Function WSPHTLSEG_Add(ByRef pWSPHTLSEG_Type As WSPHTLSEG_Type) As Boolean
Dim vSQL_String As String
    
    vSQL_String = "" _
        & "INSERT INTO WSPHTLSEG (" _
            & "UpLoadNo,RecID,SEGNO,CHAINCODE" _
            & ",CHAINNAME,STATUSCODE,ROOMS,CTYCODE" _
            & ",INDATE,OUTDATE,PROPCODE,PROPNAME" _
            & ",NOPER,TYPE,CURRCODE,SYSRAMT" _
            & ",SYSRPLAN,ROOMDESC,RATEDESC,ROOMLOC" _
            & ",BSOURCE,GUESTNAME,DISCOUNTNO,FTNO" _
            & ",FGUESTNO,TOTALAMT,BASEAMT,SCHRGEAMT" _
            & ",SURCHARGE,TTD,COMAMT,VCOM" _
            & ",CONFNO,CNCLNNO,TAXRATE,HOTADD1" _
            & ",HOTADD2,CNTRYCODE,POSTELCODE,TELENO" _
            & ",FAXNO,CHKINTIME,CHKOUTTIME,RGCURRCODE" _
            & ",RGAMT,EXRATEDTE,CURRCODELOC,CNTRYCODEH" _
            & ",EXCHRATE,RT,RTCURRCODE,RTAMT" _
            & ",RQ,RQCURRCODE,RQAMT,RQ1" _
            & ",RQ1CURRCODE,RQ1AMT,RG1,RG1CURRCODE" _
            & ",RG1AMT,RT1,RT1CURRCODE,RT1AMT"
        vSQL_String = vSQL_String & ") VALUES ("
        vSQL_String = vSQL_String & "" _
            & "" & pWSPHTLSEG_Type.UpLoadNo_Long & ",'" & SkipChars(pWSPHTLSEG_Type.RecID_String) & "','" & SkipChars(pWSPHTLSEG_Type.SEGNO_String) & "','" & SkipChars(pWSPHTLSEG_Type.CHAINCODE_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.CHAINNAME_String) & "','" & SkipChars(pWSPHTLSEG_Type.STATUSCODE_String) & "','" & SkipChars(pWSPHTLSEG_Type.ROOMS_String) & "','" & SkipChars(pWSPHTLSEG_Type.CTYCODE_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.INDATE_String) & "','" & SkipChars(pWSPHTLSEG_Type.OUTDATE_String) & "','" & SkipChars(pWSPHTLSEG_Type.PROPCODE_String) & "','" & SkipChars(pWSPHTLSEG_Type.PROPNAME_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.NOPER_String) & "','" & SkipChars(pWSPHTLSEG_Type.TYPE_String) & "','" & SkipChars(pWSPHTLSEG_Type.CURRCODE_String) & "','" & SkipChars(pWSPHTLSEG_Type.SYSRAMT_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.SYSRPLAN_String) & "','" & SkipChars(pWSPHTLSEG_Type.ROOMDESC_String) & "','" & SkipChars(pWSPHTLSEG_Type.RATEDESC_String) & "','" & SkipChars(pWSPHTLSEG_Type.ROOMLOC_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.BSOURCE_String) & "','" & SkipChars(pWSPHTLSEG_Type.GUESTNAME_String) & "','" & SkipChars(pWSPHTLSEG_Type.DISCOUNTNO_String) & "','" & SkipChars(pWSPHTLSEG_Type.FTNO_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.FGUESTNO_String) & "','" & SkipChars(pWSPHTLSEG_Type.TOTALAMT_String) & "','" & SkipChars(pWSPHTLSEG_Type.BASEAMT_String) & "','" & SkipChars(pWSPHTLSEG_Type.SCHRGEAMT_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.SURCHARGE_String) & "','" & SkipChars(pWSPHTLSEG_Type.TTD_String) & "','" & SkipChars(pWSPHTLSEG_Type.COMAMT_String) & "','" & SkipChars(pWSPHTLSEG_Type.VCOM_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.CONFNO_String) & "','" & SkipChars(pWSPHTLSEG_Type.CNCLNNO_String) & "','" & SkipChars(pWSPHTLSEG_Type.TAXRATE_String) & "','" & SkipChars(pWSPHTLSEG_Type.HOTADD1_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.HOTADD2_String) & "','" & SkipChars(pWSPHTLSEG_Type.CNTRYCODE_String) & "','" & SkipChars(pWSPHTLSEG_Type.POSTELCODE_String) & "','" & SkipChars(pWSPHTLSEG_Type.TELENO_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.FAXNO_String) & "','" & SkipChars(pWSPHTLSEG_Type.CHKINTIME_String) & "','" & SkipChars(pWSPHTLSEG_Type.CHKOUTTIME_String) & "','" & SkipChars(pWSPHTLSEG_Type.RGCURRCODE_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.RGAMT_String) & "','" & SkipChars(pWSPHTLSEG_Type.EXRATEDTE_String) & "','" & SkipChars(pWSPHTLSEG_Type.CURRCODELOC_String) & "','" & SkipChars(pWSPHTLSEG_Type.CNTRYCODEH_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.EXCHRATE_String) & "','" & SkipChars(pWSPHTLSEG_Type.RT_String) & "','" & SkipChars(pWSPHTLSEG_Type.RTCURRCODE_String) & "','" & SkipChars(pWSPHTLSEG_Type.RTAMT_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.RQ_String) & "','" & SkipChars(pWSPHTLSEG_Type.RQCURRCODE_String) & "','" & SkipChars(pWSPHTLSEG_Type.RQAMT_String) & "','" & SkipChars(pWSPHTLSEG_Type.RQ1_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.RQ1CURRCODE_String) & "','" & SkipChars(pWSPHTLSEG_Type.RQ1AMT_String) & "','" & SkipChars(pWSPHTLSEG_Type.RG1_String) & "','" & SkipChars(pWSPHTLSEG_Type.RG1CURRCODE_String) & "'" _
            & ",'" & SkipChars(pWSPHTLSEG_Type.RG1AMT_String) & "','" & SkipChars(pWSPHTLSEG_Type.RT1_String) & "','" & SkipChars(pWSPHTLSEG_Type.RT1CURRCODE_String) & "','" & SkipChars(pWSPHTLSEG_Type.RT1AMT_String) & "'" _
            & ")"
    dbCompany.Execute vSQL_String

WSPHTLSEG_Add = True
End Function


