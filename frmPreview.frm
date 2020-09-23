VERSION 5.00
Begin VB.Form frmPreview 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preview of ...."
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2820
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   173
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   188
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   2565
      Left            =   0
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   182
      TabIndex        =   1
      Top             =   0
      Width           =   2790
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2280
         Left            =   0
         MouseIcon       =   "frmPreview.frx":0000
         ScaleHeight     =   152
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   183
         TabIndex        =   2
         Top             =   225
         Width           =   2745
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   -15
         TabIndex        =   3
         Top             =   15
         Width           =   2760
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   1005
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   0
      Top             =   3540
      Visible         =   0   'False
      Width           =   2235
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A WML Debugger/Browser (quite successful)

Private Type Link
  x As Integer
  y As Integer
  x2 As Integer
  url As String
End Type

Dim Links(20) As Link
Dim lcnt As Integer

Sub LoadFile(FileName As String)
Dim wml$
Select Case UCase(Right(FileName, 3))
 Case "WML"
   If Form1.GetFilePath(FileName) = "" Or Dir(Form1.GetFilePath(FileName)) = "" Then
     MsgBox "Error: File not on server", vbExclamation
     Exit Sub
   End If
   
   Open Form1.GetFilePath(FileName) For Input As #1
   Do While Not EOF(1)
     Line Input #1, t
     wml = wml & t & vbCrLf
   Loop
   Close #1
   Render wml
   Me.Caption = "Preview of " & FileName
 Case Else
   ShellExecute Me.hwnd, vbNullString, Form1.GetFilePath(FileName), vbNullString, "C:\", SW_SHOWNORMAL
End Select
End Sub

Public Sub Render(ByVal wml$)
Dim CurrentMode(20) As String
Dim Depth As Integer
Dim i&, CurrCmd$
Dim LineCount&, CharCountInLine& 'For Debbuger Messages
Dim DisplayX&, DisplayY&, PrevWasSpace As Boolean
Erase Links
lcnt = 0
Picture3.Picture = Me.Picture
Picture3.Cls
i = 1
PrevWasSpace = True
Do
If Mid(wml, i, 1) = "<" Then
  CurrCmd = ""
  Do
    i = i + 1
    CharCountInLine = CharCountInLine + 1
    If Mid(wml, i, 1) <> ">" Then CurrCmd = CurrCmd & Mid(wml, i, 1)
  Loop Until (Mid(wml, i, 1) = ">") Or (i >= Len(wml))
  If Right(CurrCmd, 1) = "/" Or Mid(CurrCmd, 1, 1) = "!" Or Mid(CurrCmd, 1, 1) = "?" Then
     If Right(CurrCmd, 1) = "/" Then
       Select Case SAS(Mid(CurrCmd, 1, Len(CurrCmd) - 1))
          Case "br"
             DisplayX = 0
             DisplayY = DisplayY + Picture3.TextHeight("|")
          Case "img"
             Picture2.Picture = LoadPicture(GetSrc(Mid(CurrCmd, 1, Len(CurrCmd) - 1)))
             Picture3.PaintPicture Picture2.Picture, DisplayX, DisplayY
             DisplayX = DisplayX + Picture2.ScaleWidth
             DisplayY = DisplayY + Picture2.ScaleHeight - Picture3.TextHeight("|")
          Case "input"
             Picture3.Line (DisplayX, DisplayY)-(DisplayX + 70, DisplayY + Picture3.TextHeight("|")), , B
       End Select
     End If
     If Mid(CurrCmd, 1, 1) = "?" Then
        If Right(CurrCmd, 1) <> "?" Then MsgBox "Syntax Error" & vbCrLf & "Expected ?" & "Line: " & LineCount + 1 & " Char: " & CharCountInLine
     End If
     i = i + 1
  Else
     If Mid(CurrCmd, 1, 1) <> "/" Then
       Depth = Depth + 1
       CurrentMode(Depth) = CurrCmd
       i = i + 1
       ''''''''''''''''''''''''''''''
       If SAS(CurrCmd) = "a" Then
          Links(lcnt).x = DisplayX
          Links(lcnt).y = DisplayY
          Links(lcnt).url = GetParam(CurrCmd, "href")
       End If
       ''''''''''''''''''''''''''''''
       If SAS(CurrCmd) = "card" Then Label1.Caption = GetTitle(CurrCmd)
     Else
       If Mid(CurrCmd, 2) <> SAS(CurrentMode(Depth)) Then
          MsgBox "Syntax Error" & vbCrLf & "Expected </" & SAS(CurrentMode(Depth)) & ">" & vbCrLf & "Line: " & LineCount + 1 & " Char: " & CharCountInLine - Len(CurrCmd) - 1, vbExclamation, "XML Debugger"
          Exit Sub
       End If
       Depth = Depth - 1
       i = i + 1
       If Mid(CurrCmd, 2) = "a" Then
          Links(lcnt).x2 = DisplayX
          lcnt = lcnt + 1
       End If
     End If
  End If
Else
If Mid(wml, i, 2) = vbCrLf Then
   LineCount = LineCount + 1
   CharCountInLine = 0
End If

If Mid(wml, i, 1) <> Chr(10) And Mid(wml, i, 1) <> Chr(13) Then

If Mid(wml, i, 1) <> " " Then
  With Picture3
  .ForeColor = vbBlack
  .FontSize = 10
  .FontBold = False
  .FontItalic = False
  .FontUnderline = False
  For j = 1 To Depth
    Select Case SAS(CurrentMode(j))
       Case "a":      .ForeColor = vbBlue
       Case "small":  .FontSize = 8
       Case "big":    .FontSize = 12
       Case "strong": .ForeColor = &H44&
       Case "em":     .FontSize = 12: .ForeColor = &H44&
       Case "u":      .FontUnderline = True
       Case "i":      .FontItalic = True
       Case "b":      .FontBold = True
    End Select
  Next
  End With
  Picture3.CurrentX = DisplayX
  DisplayX = DisplayX + Picture3.TextWidth(Mid(wml, i, 1))
  If Mid(wml, i + 1, 1) = " " Then DisplayX = DisplayX + 8
  
  Picture3.CurrentY = DisplayY
  Picture3.Print Mid(wml, i, 1)
  If DisplayX >= Picture3.Width Then
     DisplayX = 0
     DisplayY = DisplayY + Picture3.TextHeight("|")
  End If
End If

End If
'Do stuff
i = i + 1
CharCountInLine = CharCountInLine + 1
End If
Loop Until i >= Len(wml)
If Depth > 0 Then
   MsgBox "Syntax Error" & vbCrLf & "Expected </" & SAS(CurrentMode(Depth)) & ">" & vbCrLf & "Line: " & LineCount + 1 & " Char: " & CharCountInLine, vbExclamation, "XML Debugger"
   Exit Sub
End If
Picture3.Refresh
Picture3.Picture = Picture3.Image

End Sub

Function SAS(text As String) As String
  If InStr(1, text, " ") = 0 Then
    SAS = text
  Else
    SAS = Mid(text, 1, InStr(1, text, " ") - 1)
  End If
End Function

Function GetSrc(text As String) As String
Dim s&, e&, r$
s = InStr(1, text, "src=") + 5
e = InStr(s, text, """")
If s > 0 And e - s > 0 Then
  r = Mid(text, s, e - s)
  If Form1.GetFilePath(r) = "" Then
    If Mid(r, 2, 1) <> ":" Then
      GetSrc = App.Path & "\" & r
    Else
      GetSrc = r
    End If
  Else
    GetSrc = Form1.GetFilePath(r)
  End If
End If
End Function

Function GetTitle(text As String) As String
Dim s&, e&, r$
s = InStr(1, text, "title=") + 7
e = InStr(s, text, """")
If s > 0 And e - s > 0 Then
  GetTitle = Mid(text, s, e - s)
End If
End Function

Function GetParam(text As String, param As String) As String
Dim s&, e&, r$
s = InStr(1, text, param & "=") + Len(param) + 2
e = InStr(s, text, """")
If s > 0 And e - s > 0 Then
  GetParam = Mid(text, s, e - s)
End If
End Function


Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    For i = 0 To 20
      If Links(i).x < x And Links(i).x2 > x And Links(i).y < y And Links(i).y + Picture3.TextHeight("|") > y Then
        WaitUntilMouseUp = True
        LoadFile Links(i).url
        Exit For
      End If
    Next
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Picture3.Cls
  Picture3.MousePointer = 0
  Picture3.ToolTipText = ""
  For i = 0 To 20
  If Links(i).x < x And Links(i).x2 > x And Links(i).y < y And Links(i).y + Picture3.TextHeight("|") > y Then
     Picture3.MousePointer = 99
     Picture3.Line (Links(i).x, Links(i).y + Picture3.TextHeight("|") - 2)-(Links(i).x2, Links(i).y + Picture3.TextHeight("|") - 2), vbBlue
     Picture3.ToolTipText = Links(i).url
  End If
  Next
End Sub

