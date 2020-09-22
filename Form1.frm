VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "HTTP File Download"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save File"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download File"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private total, alldata As String
Private totalsize
Private strHost As String, strFile As String

Private Sub cmdSave_Click()
    Call SaveFile
    
    strHost = vbNull
    strFile = vbNull
    totalsize = vbNull
    total = vbNull
    alldata = vbNull
    Label1.Caption = ""
    Label2.Caption = ""
    ProgressBar.Value = 0
    Command1.Enabled = True
    
End Sub

Private Sub Command1_Click()
    Winsock1.Close
    
    '''parse strHost and strFile out of Text2.text'''
    Dim strAddress As String, intSlash As Integer
    
    strAddress = InputBox("Enter the HTTP URL for the file to Download:" & vbNewLine & "(ex. http://pietschsoft.com/images/logo.gif)", "Enter URL", "http://pietschsoft.com/images/logo.gif")
    
    If UCase(Left(strAddress, 7)) = "HTTP://" Then
        strAddress = Right(strAddress, (Len(strAddress) - 7))
    End If
    
    intSlash = InStr(1, strAddress, "/", vbTextCompare) - 1
    
    If intSlash <= 0 Then intSlash = Len(strAddress)
    
    strHost = Left(strAddress, intSlash)
    strFile = Right(strAddress, (Len(strAddress) - (Len(strHost))))
    
    If Left(strFile, 1) <> "/" Then strFile = "/" & strFile
    If Len(strFile) = 0 Then strFile = "/"
    '''''''''''''''''''''''''''''''''''''''''''''''''
    
    Label1.Caption = strAddress
    Command1.Enabled = False
    
    '''''''''''''''''''''''''''''''''''''''''''''''''
    
    'connect to the server
    Winsock1.Protocol = sckTCPProtocol
    Winsock1.Connect strHost, 80
    
End Sub

Private Sub Form_Load()
    cmdSave.Enabled = False
    
End Sub

Private Sub Winsock1_Connect()
    'send request to server for the file, once connected to server
    Winsock1.SendData "GET " & strFile & " HTTP/1.0" & vbCrLf
    Winsock1.SendData vbCrLf
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    
    'put entire file into the alldata variable
    Winsock1.GetData data
    alldata = alldata & data
    
    'display file completion status
    total = total + bytesTotal
    If totalsize = 0 Then
        totalsize = GetHTTPInfo(data, HTTP_Data_Content_Length)
        Label1.Caption = Label1.Caption & " -- " & totalsize & " Bytes"
        ProgressBar.Max = totalsize
    End If
    If total > ProgressBar.Max Then total = ProgressBar.Max
    ProgressBar.Value = total
    
    If ProgressBar.Max <= ProgressBar.Value Then 'if download complete
        cmdSave.Enabled = True
    End If
    
    Label2.Caption = Format((ProgressBar.Value / ProgressBar.Max * 100), "0") & "% Complete"
        
End Sub

Sub SaveFile()
    Dim strSaveFile As String, strFileName As String
    Dim i As Integer
    Do Until Left(strFileName, 1) = "/"
        strFileName = Right(strFile, i)
        i = i + 1
    Loop
    strFileName = Right(strFileName, Len(strFileName) - 1)
    
    CommonDialog.DialogTitle = "Save File As"
    CommonDialog.FileName = strFileName
    CommonDialog.ShowSave
    
    If CommonDialog.CancelError = True Then Exit Sub
    
    strSaveFile = CommonDialog.FileName
    
    'write the entire file contents recieved to an actual file file
    Open strSaveFile For Binary As #1
    Put #1, , GetHTTPInfo(alldata, HTTP_Data_Data)
    Close #1
    
    cmdSave.Enabled = False
    
End Sub
