Attribute VB_Name = "Convensions"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Private Declare Function GetTimeZoneInformation Lib "kernel32" _
                          (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Type SYSTEMTIME
    wYear                       As Integer
    wMonth                      As Integer
    wDayOfWeek                  As Integer
    wDay                        As Integer
    wHour                       As Integer
    wMinute                     As Integer
    wSecond                     As Integer
    wMilliseconds               As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias                        As Long
    StandardName(63)            As Byte
    StandardDate                As SYSTEMTIME
    StandardBias                As Long
    DaylightName(63)            As Byte
    DaylightDate                As SYSTEMTIME
    DaylightBias                As Long
End Type
Public Function ContentType(Extens As String) As String
'References
'http://www.iana.org/assignments/media-types/
'http://www.wapforum.org/wina/wsp-content-type.htm
'http://www.utoronto.ca/webdocs/HTMLdocs/Book/Book-3ed/appb/mimetype.html
'http://www.utoronto.ca/ian/books/xhtml1/mime/mimetype.html#audi
'HKEY_CLASSES_ROOT\MIME\Database\Content Type
  Select Case LCase(Extens)
    Case "jad": ContentType = "text/vnd.sun.j2me.app-descriptor"
    Case "jar": ContentType = "application/java-archive"
    Case "wml", "xml": ContentType = "text/vnd.wap.wml"
    Case "mid": ContentType = "audio/midi"
    Case "jpg": ContentType = "image/jpeg"
    Case "gif": ContentType = "image/gif"
    Case "wmlc": ContentType = "application/vnd.wap.wmlc"
    Case "wbxml": ContentType = "application/vnd.wap.wbxml"
    Case "wmlsc": ContentType = "application/vnd.wap.wmlscriptc"
    Case "sic": ContentType = "application/vnd.wap.sic"
    Case "mmf": ContentType = "application/vnd.smaf" 'audio
    Case "wmls": ContentType = "text/vnd.wap.wmlscript"
    Case "wbmp": ContentType = "image/vnd.wap.wbmp"
    Case "wav": ContentType = "audio/x-wav"
    Case "mmid": ContentType = "x-music/x-midi"
    Case "amr": ContentType = "audio/amr"
    Case "ico": ContentType = "image/x-icon"
  Case Else
      Dim RegVal$
      RegVal$ = GetKeyValue(HKEY_CLASSES_ROOT, "." & LCase(Extens), "Content Type")
      If RegVal = "" Then ErrorLog "Unknown Content type: ." & Extens
      ContentType = RegVal
  End Select
End Function

Public Function dIcon(Extens As String) As Integer
  Select Case LCase(Extens)
    Case "wml", "xml", "wmlc", "wbxml", "wmlsc", "sic", "wmls", "wbmp": dIcon = 0
    Case "mid", "mmid", "mmf": dIcon = 1
    Case "wav", "amr", "mp3": dIcon = 2
    Case "jpg", "gif", "bmp", "ico": dIcon = 3
    Case "jad", "jar": dIcon = 4
  End Select
End Function
Public Sub ErrorLog(ErrDesc As String)
  Debug.Print ErrDesc
  Open App.Path & "\Logs\ErrorLog.txt" For Append As #1
     Print #1, ErrDesc
  Close #1
End Sub

Public Function ddd(ByVal day As Integer) As String
  'because format(date,"ddd") may return day in computer native language
  Select Case day
    Case 1: ddd = "Sun"
    Case 2: ddd = "Mon"
    Case 3: ddd = "Tue"
    Case 4: ddd = "Wed"
    Case 5: ddd = "Thu"
    Case 6: ddd = "Fri"
    Case 7: ddd = "Sat"
  End Select
End Function

Public Function mmm(ByVal month As Integer) As String
  'because format(date,"mmm") may return day in computer native language
  Select Case month
    Case 1: mmm = "Jan"
    Case 2: mmm = "Feb"
    Case 3: mmm = "Mar"
    Case 4: mmm = "Apr"
    Case 5: mmm = "May"
    Case 6: mmm = "Jun"
    Case 7: mmm = "Jul"
    Case 8: mmm = "Aug"
    Case 9: mmm = "Sep"
    Case 10: mmm = "Oct"
    Case 11: mmm = "Nov"
    Case 12: mmm = "Dec"
  End Select
End Function

Function GetGMTDateTime() As String
    'Wild function to get the GMT Date/Time
    Dim utTZ As TIME_ZONE_INFORMATION
    Dim h&, m&, hh&, mm&, dy&, mo&, yy&
    Select Case GetTimeZoneInformation(utTZ)
      Case TIME_ZONE_ID_DAYLIGHT
        dwBias = utTZ.Bias + utTZ.DaylightBias
      Case Else
        dwBias = utTZ.Bias + utTZ.StandardBias
    End Select
    h = dwBias \ 60
    m = dwBias - (dwBias \ 60) * 60
    hh = Hour(Time) + h
    mm = Minute(Time) + m
    dy = day(Date)
    mo = month(Date)
    yy = Year(Date)
    If mm < 0 Then
       mm = mm + 60
       hh = hh - 1
    End If
    If mm > 60 Then
       mm = mm - 60
       hh = hh + 1
    End If
    If hh < 0 Then
       hh = hh + 24
       dy = dy - 1
    End If
    If hh > 24 Then
       hh = hh - 24
       dy = dy + 1
    End If
    If dy <= 0 Then
       mo = mo - 1
       dy = MonthDays(mo, yy)
    End If
    If dy > MonthDays(mo, yy) Then
       mo = mo + 1
       dy = 1
    End If
    If mo < 0 Then
       mo = mo + 12
       yy = yy - 1
    End If
    If mo > 12 Then
       mo = mo - 12
       yy = yy + 1
    End If
    
    GetGMTDateTime = ddd(Weekday(DateSerial(yy, mo, dy))) & ", " & Format(dy, "00") & " " & mmm(mo) & " " & yy & " " & Format(hh, "00") & ":" & Format(mm, "00") & ":" & Format(Second(Time), "00") & " GMT"
End Function

Function MonthDays(ByVal month As Integer, ByVal inYear As Integer) As Integer
'Todo add leapyear support
Select Case month
  Case 1, 3, 5, 7, 8, 10, 12: MonthDays = 31
  Case 4, 6, 9, 11: MonthDays = 30
  Case 2:
    If ((inYear Mod 4 = 0) And (inYear Mod 100 <> 0) Or (inYear Mod 400 = 0)) Then
      MonthDays = 29
    Else
      MonthDays = 28
    End If
End Select
End Function
