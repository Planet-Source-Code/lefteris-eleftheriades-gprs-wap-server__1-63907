VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Wap/GPRS Server"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3360
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Get URL"
      Height          =   270
      Left            =   2625
      TabIndex        =   13
      Top             =   45
      Width           =   1005
   End
   Begin wapserver.Socket Client 
      Index           =   0
      Left            =   4365
      Top             =   705
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin wapserver.Socket Socket1 
      Left            =   1440
      Top             =   -30
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   495
      TabIndex        =   11
      Text            =   "(obdaining)"
      Top             =   75
      Width           =   2115
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   315
      Left            =   7425
      TabIndex        =   9
      Top             =   5640
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   "Connections list"
      Height          =   3285
      Left            =   6825
      TabIndex        =   5
      Top             =   315
      Width           =   1830
      Begin VB.CommandButton Command1 
         Caption         =   "Block Client"
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   1605
      End
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   90
         TabIndex        =   6
         Top             =   195
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Index"
      Height          =   5325
      Left            =   0
      TabIndex        =   0
      Top             =   315
      Width           =   6765
      Begin VB.CommandButton Command5 
         Caption         =   "Save Conf"
         Height          =   315
         Left            =   4215
         TabIndex        =   16
         Top             =   4920
         Width           =   1035
      End
      Begin VB.PictureBox FileList 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4590
         Left            =   90
         OLEDropMode     =   1  'Manual
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   302
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   436
         TabIndex        =   4
         Top             =   225
         Width           =   6600
         Begin VB.PictureBox Picture2 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00CC3300&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   1575
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   15
            Top             =   2460
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   1575
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   14
            Top             =   1965
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.PictureBox iconlist 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1440
            Left            =   3165
            Picture         =   "Form1.frx":1EF8
            ScaleHeight     =   96
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   128
            TabIndex        =   12
            Top             =   2370
            Visible         =   0   'False
            Width           =   1920
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add file"
         Height          =   315
         Index           =   0
         Left            =   450
         TabIndex        =   3
         Top             =   4920
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove file"
         Height          =   315
         Index           =   1
         Left            =   1695
         TabIndex        =   2
         Top             =   4920
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Index           =   2
         Left            =   2940
         Picture         =   "Form1.frx":AF3A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4920
         Width           =   1035
      End
   End
   Begin VB.Label Label2 
      Caption         =   "URL:"
      Height          =   195
      Left            =   15
      TabIndex        =   10
      Top             =   75
      Width           =   465
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status: Active"
      Height          =   270
      Left            =   15
      TabIndex        =   8
      Top             =   5670
      Width           =   7230
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu mnuPreviewItm 
         Caption         =   "Preview"
      End
      Begin VB.Menu mnuContextItm 
         Caption         =   "Remove"
         Index           =   0
      End
      Begin VB.Menu mnuContextItm 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuContextItm 
         Caption         =   "Set as Index"
         Index           =   3
      End
      Begin VB.Menu mnuContextItm 
         Caption         =   "Set as 404"
         Index           =   4
      End
      Begin VB.Menu mnuContextItm 
         Caption         =   "Set as 500"
         Index           =   5
      End
   End
   Begin VB.Menu mnuSetAs 
      Caption         =   "Set As"
      Visible         =   0   'False
      Begin VB.Menu mnuSetAsItm 
         Caption         =   "Index"
         Index           =   0
      End
      Begin VB.Menu mnuSetAsItm 
         Caption         =   "Not found page"
         Index           =   1
      End
      Begin VB.Menu mnuSetAsItm 
         Caption         =   "Server error page"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Address already in use means that some instance of the program is
'Listening to port 80 try reopening VB
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCAND As Long = &H8800C6
Private Const SRCINVERT As Long = &H660046
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, _
                                                      ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type FileItems
 x As Integer
 y As Integer
 Icon As Byte
 text As String
 Path As String
 Special As Byte
End Type
   Const IconSpacingX& = 58 + 20
   Const IconSpacingY& = 58 + 20

Dim FileItem() As FileItems
Private Type ClientDataArrivals
    cData As String
    cCompleted As Boolean
End Type
Dim ClientDataArrival(100) As ClientDataArrivals
Dim CRC&
Dim SelectedFileIndex&
Dim bSent(100) As Boolean

Public Function GetFilePath(file As String) As String
  Dim i&
  On Error Resume Next
  For i = 0 To UBound(FileItem)
    If FileItem(i).text = file Then GetFilePath = FileItem(i).Path
  Next
End Function


Sub RenderFileList(Optional ExcludeX% = -1, Optional ExcludeY% = -1)
   Dim i&, txCapt$
   For i = 1 To UBound(FileItem)
     
       BitBlt Picture1.hdc, 0, 0, 32, 32, iconlist.hdc, FileItem(i).Icon * 32, 0, SRCCOPY
       AlphaBlend Picture1.hdc, 0, 0, 32, 32, Picture2.hdc, 0, 0, 32, 32, &H10000 * 128
       BitBlt Picture1.hdc, 0, 0, 32, 32, iconlist.hdc, FileItem(i).Icon * 32, 64, SRCAND
       
       
       BitBlt FileList.hdc, FileItem(i).x * IconSpacingX + 10 + IconSpacingX& / 2 - 16, FileItem(i).y * IconSpacingY + 10, 32, 32, iconlist.hdc, FileItem(i).Icon * 32, 32, SRCAND
       If ExcludeX = FileItem(i).x And ExcludeY = FileItem(i).y Then
         BitBlt FileList.hdc, FileItem(i).x * IconSpacingX + 10 + IconSpacingX& / 2 - 16, FileItem(i).y * IconSpacingY + 10, 32, 32, Picture1.hdc, 0, 0, SRCINVERT
       Else
         BitBlt FileList.hdc, FileItem(i).x * IconSpacingX + 10 + IconSpacingX& / 2 - 16, FileItem(i).y * IconSpacingY + 10, 32, 32, iconlist.hdc, FileItem(i).Icon * 32, 0, SRCINVERT
       End If
       txCapt = Mid(FileItem(i).text, 1, Len(FileItem(i).text) - 4)
       FileList.CurrentX = FileItem(i).x * IconSpacingX + 10 + (IconSpacingX - FileList.TextWidth(txCapt)) / 2
       If ExcludeX = FileItem(i).x And ExcludeY = FileItem(i).y Then AlphaBlend FileList.hdc, FileItem(i).x * IconSpacingX + 10 + (IconSpacingX - FileList.TextWidth(txCapt)) / 2 - 2, FileItem(i).y * IconSpacingY + 10 + 34, FileList.TextWidth(txCapt) + 4 + 2, FileList.TextHeight("|") + 2, Picture2.hdc, 0, 0, 32, 32, &H10000 * 128
       FileList.CurrentY = FileItem(i).y * IconSpacingY + 10 + 34
       FileList.Print txCapt
       
   Next
End Sub

Private Sub Client_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  'THIS CODE REFFERS TO CLIENT(0) THAT IS ALWAYS LISTENING
  Dim DontCreateNewSocks As Boolean, AvailiableSockIndex As Integer
  'Sees if a loaded socks control is free. if so assigns request to that one, else it creates a new one
  For i = 1 To Client.Count - 1
     If Client(i).State = sckClosed Then
        DontCreateNewSocks = True
        AvailiableSockIndex = i
        Exit For
     End If
  Next i
  If DontCreateNewSocks = False Then
     AvailiableSockIndex = Client.Count
     Load Client(AvailiableSockIndex)
  End If
  Client(AvailiableSockIndex).Accept requestID
  
End Sub

Private Sub Client_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim a$, L&, Request$, DCrLfPos&, ContentLengthPos&, ContentLength&
Client(Index).GetData a
  
'''HERE WE CHECK IF THE REQUEST IS COMPLETE OR IS MISSING ANY PIECE'''
  If InStr(1, a, " ") <> 0 Then Request = Mid(a, 1, InStr(1, a, " ") - 1)
  
  If Request = "GET" Or Request = "POST" Or Request = "OPTION" Then
  'If it's a fresh request
     'Store the data-arrived to cData
     ClientDataArrival(Index).cData = a
     'Mark as incomplete (check will be preformed later to see if complete or not)
     ClientDataArrival(Index).cCompleted = False
  ElseIf ClientDataArrival(Index).cCompleted = False Then
  'If last data-arrived was incomplete then add to cData the current data-arrived
     ClientDataArrival(Index).cData = ClientDataArrival(Index).cData & a
  Else
     MsgBox "Bad request: " & Request
  End If
  'Check if a double CrLF is found. (this indicates end of header)
  DCrLfPos = InStr(1, ClientDataArrival(Index).cData, vbCrLf & vbCrLf)
  If DCrLfPos <> 0 Then
  'Found Double CrLF
      'Check if a content-length tag exists (means there are more data to go after the header)
      ContentLengthPos = InStr(1, ClientDataArrival(Index).cData, "Content-Length:")
      
      If ContentLengthPos <> 0 Then
      'Found content length tag
         'Get it's value
         ContentLength = Val(Mid(ClientDataArrival(Index).cData, ContentLengthPos + 15, InStr(ContentLengthPos + 15, ClientDataArrival(Index).cData, vbCrLf) - ContentLengthPos - 15))
         'If the length of the data we have agrees with the header+content length data
         If DCrLfPos + 4 + ContentLength - 1 = Len(ClientDataArrival(Index).cData) Then
            'Request completed and not missing any part
            ClientDataArrival(Index).cCompleted = True
         End If
      Else
      'Content length tag not found => only header in request
         'Request completed and not missing any part
         ClientDataArrival(Index).cCompleted = True
      End If
  End If
'''END OF CHECK IF THE REQUEST IS COMPLETE OR IS MISSING ANY PIECE'''


If ClientDataArrival(Index).cCompleted Then
    'Store the request for later reference
    Open App.Path & "\Logs\CompleteReq" & CRC & ".txt" For Binary As #1
      Put #1, , ClientDataArrival(Index).cData
    Close #1
    CRC = CRC + 1
    
    Dim UserAgentPos&, UserAgentEndPos&
    UserAgentPos = InStr(1, ClientDataArrival(Index).cData, "User-Agent:") + 12
    UserAgentEndPos = InStr(UserAgentPos, ClientDataArrival(Index).cData, vbCrLf)
    List1.AddItem Mid(ClientDataArrival(Index).cData, UserAgentPos, UserAgentEndPos - UserAgentPos)
End If

'PROCESS THE COMPLETED REQUEST
If ClientDataArrival(Index).cCompleted Then
  Request = Mid(ClientDataArrival(Index).cData, 1, InStr(1, ClientDataArrival(Index).cData, " ") - 1)
  Select Case Request
  Case "GET"
    'This is a request for sending a file to the client
    
    'Track the file requested
    Dim bFlag As Integer
    bFlag = -1
    Dim ReqFileName$
    ReqFileName = UrlDecode(Mid(ClientDataArrival(Index).cData, 6, InStr(5, ClientDataArrival(Index).cData, " ") - 6))
    
SendFile:
    For i = 1 To UBound(FileItem)
      If LCase(ReqFileName) = LCase(FileItem(i).text) Then bFlag = i
      If ReqFileName = "" And FileItem(i).Special = 1 Then bFlag = i
    Next
    If bFlag = -1 Then
      For i = 1 To UBound(FileItem)
        If FileItem(i).Special = 2 Then bFlag = i
      Next
    End If
      
      Dim data() As Byte
      'Open the requested file and extract it's data
      If Dir(FileItem(bFlag).Path) = "" Then
         ErrorLog "File not exists: " & FileItem(bFlag).Path
         Client(Index).CloseSck
         Exit Sub
      End If
      
      Open FileItem(bFlag).Path For Binary As #1
        L = LOF(1)
        ReDim data(LOF(1) - 1)
        Get #1, , data
      Close #1
      'Send header and data to client
      Client(Index).SendData _
      "HTTP/1.0 200 OK" & vbCrLf & _
      "Date: " & GetGMTDateTime & vbCrLf & _
      "P3P: policyref=""http://p3p.yahoo.com/w3c/p3p.xml"", CP=""CAO DSP COR CUR ADM DEV TAI PSA PSD IVAi IVDi CONi TELo OTPi OUR DELi SAMi OTRi UNRi PUBi IND PHY ONL UNI PUR FIN COM NAV INT DEM CNT STA POL HEA PRE GOV""" & vbCrLf & _
      "Last-Modified: Tue, 20 May 2003 21:08:48 GMT" & vbCrLf & _
      "ETag: ""da68c-42d-3eca9960""" & vbCrLf & _
      "Accept-Ranges: bytes" & vbCrLf & _
      "Content-Length: " & L & vbCrLf & _
      "Content-Type: " & ContentType(Right(FileItem(bFlag).Path, 3)) & vbCrLf & _
      "Age: 2606" & vbCrLf & _
      "Connection: close" & vbCrLf & vbCrLf & StrConv(data(), vbUnicode)
      bSent(Index) = False
      Do
        DoEvents
      Loop Until bSent(Index)
      Client(Index).CloseSck
  Case "POST"
      ContentLengthPos = InStr(1, ClientDataArrival(Index).cData, "Content-Length:")
      DCrLfPos = InStr(1, ClientDataArrival(Index).cData, vbCrLf & vbCrLf)
      If ContentLengthPos <> 0 Then
         Dim PostResponce As String
         ContentLength = Val(Mid(ClientDataArrival(Index).cData, ContentLengthPos + 15, InStr(ContentLengthPos + 15, ClientDataArrival(Index).cData, vbCrLf) - ContentLengthPos - 15))
         PostResponce = Mid(ClientDataArrival(Index).cData, DCrLfPos + 4, ContentLength)
         'Process post responce here
         SavePostResponce PostResponce
         SavePostForm UrlDecode(Mid(PostResponce, InStr(1, PostResponce, "=") + 1))
         ReqFileName = UrlDecode(Mid(ClientDataArrival(Index).cData, 7, InStr(6, ClientDataArrival(Index).cData, " ") - 7))
         GoTo SendFile
      End If
  End Select
  ClientDataArrival(Index).cData = "" 'CLEAR THE REQUEST
  ClientDataArrival(Index).cCompleted = False
End If
End Sub

Private Sub Client_SendComplete(Index As Integer)
  bSent(Index) = True
End Sub

Private Sub Command2_Click(Index As Integer)
Dim C&
Select Case Index
Case 0
  cd1.Filter = "Web files|*.wml;*.xml;*.wmlc;*.wbxml;*.wmlsc;*.sic;*.wmls;*.wbmp;*.mid;*.mmid;*.mmf;*.wav;*.amr;*.mp3;*.jpg;*.gif;*.bmp;*.ico;*.jad;*.jar|Web page files|*.wml;*.xml;*.wmlc;*.wbxml;*.wmlsc;*.sic;*.wmls|Audio|*.mid;*.mmid;*.mmf;*.wav;*.amr;*.mp3|Images|*.jpg;*.gif;*.bmp;*.ico;*.wbmp|Applications|*.jad;*.jar|(All Files)|*.*"
  cd1.ShowOpen
  If cd1.FileName = "" Then Exit Sub
  C = UBound(FileItem) + 1
  ReDim Preserve FileItem(C)
  FileItem(C).x = (C - 1) Mod 5
  FileItem(C).y = (C - 1) \ 5
  FileItem(C).text = FileNameFromPath(cd1.FileName)
  FileItem(C).Path = cd1.FileName
  FileItem(C).Icon = dIcon(Right(cd1.FileName, 3))
  cd1.FileName = ""
Case 1
  For i = SelectedFileIndex To UBound(FileItem) - 1
     FileItem(i).Icon = FileItem(i + 1).Icon
     FileItem(i).Path = FileItem(i + 1).Path
     FileItem(i).text = FileItem(i + 1).text
  Next
  ReDim Preserve FileItem(UBound(FileItem) - 1)
  Case 2
  If SelectedFileIndex <> 0 And SelectedFileIndex <= UBound(FileItem) Then Me.PopupMenu mnuSetAs, , Command2(Index).Left, Command2(Index).Top + 600
End Select
  FileList.Cls
  RenderFileList
End Sub

Private Sub Command3_Click()
If Client(0).State = sckListening Then
  Client(0).CloseSck
  Command3.Caption = "Start"
  Label1.Caption = "Status: Inactive"
Else
  Client(0).Listen
  Command3.Caption = "Stop"
  Label1.Caption = "Status: Active"
End If
End Sub

Private Sub Command4_Click()
   Socket1.Connect "whatismyip.org", 80
End Sub

Private Sub Command5_Click()
If Dir(App.Path & "\Directory.txt") <> "" Then Kill App.Path & "\Directory.txt"
Open App.Path & "\Directory.txt" For Output As #1
For i = 1 To UBound(FileItem)
    Print #1, FileItem(i).Special & Replace(FileItem(i).Path, App.Path, "$Path", , , vbTextCompare)
Next i
Close #1
End Sub

Private Sub FileList_DblClick()
mnuPreviewItm_Click
End Sub

Private Sub FileList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FileList.Cls

'xx = ((x - 10 - IconSpacingX& / 2 + 16) \ IconSpacingX)
'xx = ((x - 10 - IconSpacingX& / 2 + 16) \ IconSpacingX)
xx = ((x - 10) \ IconSpacingX)
yy = (y \ IconSpacingY)
SelectedFileIndex = xx + yy * 5 + 1
RenderFileList (xx), (yy)
FileList.Refresh
  'makeit only popup when over a file
  If (Button And 2) = 2 And SelectedFileIndex <> 0 And SelectedFileIndex <= UBound(FileItem) Then PopupMenu mnuContext
End Sub

Sub LoadFiles()
  Dim File1$, C&
  
  Open App.Path & "\Directory.txt" For Input As #1
     Do While Not EOF(1)
       ReDim Preserve FileItem(UBound(FileItem) + 1)
       Line Input #1, File1
       C = C + 1
       FileItem(C).x = (C - 1) Mod 5
       FileItem(C).y = (C - 1) \ 5
       FileItem(C).Special = Mid(File1, 1, 1)
       File1 = Replace(Mid(File1, 2), "$Path", App.Path, , , vbTextCompare)
       FileItem(C).text = FileNameFromPath(File1)
       FileItem(C).Path = File1
       FileItem(C).Icon = dIcon(Right(File1, 3))
     Loop
  Close #1
End Sub

Private Sub FileList_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

  If data.Files(1) = "" Then Exit Sub
  C = UBound(FileItem) + 1
  ReDim Preserve FileItem(C)
  FileItem(C).x = (C - 1) Mod 5
  FileItem(C).y = (C - 1) \ 5
  FileItem(C).text = FileNameFromPath(data.Files(1))
  FileItem(C).Path = data.Files(1)
  FileItem(C).Icon = dIcon(Right(data.Files(1), 3))
  FileList.Cls
  RenderFileList
End Sub

Private Sub Form_Load()
   Text1.text = "http://" & Socket1.LocalIP
   ReDim FileItem(0)
   LoadFiles
   
   RenderFileList
   Client(0).LocalPort = 80
   Client(0).Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Client(0).State = sckListening Then
  Client(0).CloseSck
End If
DoEvents
End
End Sub

Private Sub mnuContextItm_Click(Index As Integer)
If Index = 0 Then Command2_Click 1
If Index > 2 Then mnuSetAsItm_Click Index - 3
End Sub

Private Sub mnuPreviewItm_Click()
Dim wml$, t$
If SelectedFileIndex > UBound(FileItem) Then Exit Sub
Select Case UCase(Right(FileItem(SelectedFileIndex).Path, 3))
 Case "WML"
   Open FileItem(SelectedFileIndex).Path For Input As #1
   Do While Not EOF(1)
     Line Input #1, t
     wml = wml & t & vbCrLf
   Loop
   Close #1
   frmPreview.Show
   frmPreview.Render wml
   frmPreview.Caption = "Preview of " & FileItem(SelectedFileIndex).text
 Case Else
   ShellExecute Me.hwnd, vbNullString, FileItem(SelectedFileIndex).Path, vbNullString, "C:\", SW_SHOWNORMAL
End Select
End Sub

Private Sub mnuSetAsItm_Click(Index As Integer)
For i = 1 To UBound(FileItem)
  If FileItem(i).Special = Index + 1 Then FileItem(i).Special = 0
Next
  FileItem(SelectedFileIndex).Special = Index + 1
End Sub

Private Sub Socket1_Connect()
  Socket1.SendData "GET / HTTP/1.1" & vbCrLf & "Connection: Keep -Alive" & vbCrLf & vbCrLf
End Sub

Private Sub Socket1_DataArrival(ByVal bytesTotal As Long)
  If bytesTotal > 0 Then
    Dim a$
    Socket1.GetData a
    Text1.text = "http://" & Mid(a, InStr(1, a, vbCrLf & vbCrLf) + 4)
  End If
End Sub

Sub SavePostResponce(Responce As String)
  Open App.Path & "\Logs\PostResponce.txt" For Append As #1
     Print #1, Responce
  Close #1
End Sub

Sub SavePostForm(Responce As String)
  Dim t$, a$, z&
  Open App.Path & "\wap site\postings.wml" For Input As #1
     Do While Not EOF(1)
        Line Input #1, t
        a = a & t & vbCrLf
     Loop
  Close #1
  z = InStr(1, a, "<p>" & vbCrLf) + 5
  a = Mid(a, 1, z) & "      " & Responce & "<br/>" & vbCrLf & Mid(a, z + 1)
  a = Mid(a, 1, Len(a) - 2)
  Kill App.Path & "\wap site\postings.wml"
  Open App.Path & "\wap site\postings.wml" For Output As #1
     Print #1, a
  Close #1
End Sub

Function FileNameFromPath(Path As String) As String
   For i = Len(Path) To 1 Step -1
      If Mid(Path, i, 1) = "\" Then
         FileNameFromPath = Mid(Path, i + 1)
         Exit For
      End If
   Next
End Function

'DO URL DECODE
Function UrlDecode(text As String) As String
On Error Resume Next
Dim i&, Out$
Do
i = i + 1
If Mid(text, i, 1) = "%" Then
  Out = Out & Chr(Val("&H" & Mid(text, i + 1, 2) & "&"))
  i = i + 2
Else
  Out = Out & Mid(text, i, 1)
End If
Loop Until i >= Len(text)
UrlDecode = Out
End Function


Function URLEncode(strBefore As String) As String
    Dim strAfter As String
    Dim intLoop As Integer
    If Len(strBefore) > 0 Then
        For intLoop = 1 To Len(strBefore) Step 2
            Select Case Val("&H" & Mid(strBefore, intLoop, 2) & "&")
                Case 48 To 57, 65 To 90, 97 To 122, 46, 45, 95, 42 '0-9, A-Z, a-z . - _ *
                   strAfter = strAfter & Chr(Val("&H" & Mid(strBefore, intLoop, 2) & "&"))
                Case 32
                   strAfter = strAfter & "+"
                Case Else
                   strAfter = strAfter & "%" & Mid(strBefore, intLoop, 2)
            End Select
    Next
End If
URLEncode = strAfter
End Function

