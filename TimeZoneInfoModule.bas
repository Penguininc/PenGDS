Attribute VB_Name = "TimeZoneInfoModule"
Public Function DateAppTimeZone() As Date
'Dim d As Date
'Dim vDateAppTimeZone_Date As Date
'    d = UTCToTimeZoneDate(UTCDate, AppTimeZoneTZI_TZI)
'    vDateAppTimeZone_Date = Format(d, "DD/MMM/YYYY")
'DateAppTimeZone = vDateAppTimeZone_Date
DateAppTimeZone = Date
End Function

Public Function TimeAppTimeZone() As Date
'Dim d As Date
'Dim vTimeAppTimeZone_Date As Date
'    d = UTCToTimeZoneDate(UTCDate, AppTimeZoneTZI_TZI)
'    vTimeAppTimeZone_Date = Format(d, "H:MM:SS AMPM")
'TimeAppTimeZone = vTimeAppTimeZone_Date
TimeAppTimeZone = Format(Now, "H:MM:SS AMPM")
End Function

Public Function NowAppTimeZone() As Date
'Dim d As Date
'Dim vNowAppTimeZone_Date As Date
'    d = UTCToTimeZoneDate(UTCDate, AppTimeZoneTZI_TZI)
'    vNowAppTimeZone_Date = Format(d, "dd/mmm/yyyy") & " " & Format(d, "h:mm:ss AMPM")
'NowAppTimeZone = vNowAppTimeZone_Date
NowAppTimeZone = Now
End Function

