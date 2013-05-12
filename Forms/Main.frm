VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form SendButt 
   BackColor       =   &H00000000&
   Caption         =   "NetCheat PC Manager"
   ClientHeight    =   6765
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9660
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Main.frx":08CA
   ScaleHeight     =   6765
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox uleBox 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   405
      Left            =   135
      TabIndex        =   3
      Text            =   "uLaunchELF path"
      Top             =   6255
      Width           =   6105
   End
   Begin VB.TextBox LogBox 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   5580
      Left            =   6030
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   180
      Width           =   3450
   End
   Begin VB.TextBox CodeWindow 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   5505
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Main.frx":D290C
      Top             =   630
      Width           =   3540
   End
   Begin VB.TextBox IPBox 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   135
      TabIndex        =   0
      Text            =   "IP Address"
      Top             =   135
      Width           =   3540
   End
   Begin MSWinsockLib.Winsock WSock1 
      Left            =   4680
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label StopDisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stop Disc"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   3800
      TabIndex        =   10
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Bootule 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Boot uLaunchELF"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   6345
      TabIndex        =   9
      Top             =   6255
      Width           =   3180
   End
   Begin VB.Label LogClear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   5805
      Width           =   3465
   End
   Begin VB.Label ConPS2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   3750
      TabIndex        =   7
      Top             =   135
      Width           =   2205
   End
   Begin VB.Label DisconnectBtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disconnect"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   3750
      TabIndex        =   6
      Top             =   675
      Width           =   2205
   End
   Begin VB.Label SendButt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send Codes"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   3750
      TabIndex        =   5
      Top             =   1215
      Width           =   2205
   End
   Begin VB.Label StartGame 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   3750
      TabIndex        =   4
      Top             =   4440
      Width           =   2205
   End
   Begin VB.Menu ToolBox 
      Caption         =   "Tools"
      Begin VB.Menu JokerTB 
         Caption         =   "Joker That"
      End
      Begin VB.Menu mnu_cbc 
         Caption         =   "CBCImport"
      End
   End
End
Attribute VB_Name = "SendButt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bootule_Click()
If WSock1.State = 7 Then
    Send HexToString("4")
    Delay 1
    Send uleBox.Text
    SaveSett 1
    WSock1.Close
Else
    MsgBox "Not Connected!"
End If

End Sub

Private Sub Bootule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bootule.ForeColor = &HFFFF00
End Sub

'Gets rid of the annoying beep...
Private Sub CodeWindow_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then
    KeyAscii = 0
End If

End Sub

Private Sub CodeWindow_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 65 And Shift = 2 Then
    CodeWindow.SelStart = 0
    CodeWindow.SelLength = Len(CodeWindow.Text)
End If

End Sub

Private Sub ConPS2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConPS2.ForeColor = &HFFFF00
End Sub

Private Sub DisconnectBtn_Click()
If WSock1.State = 7 Then
    Send HexToString("2")
    WSock1.Close
Else
    WSock1.Close
End If
End Sub

Private Sub ConPS2_Click()
    ConToPS2 IPBox.Text
End Sub

Private Sub DisconnectBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DisconnectBtn.ForeColor = &HFFFF00
End Sub

Private Sub Form_Load()
SendC = False
sfile = App.Path & "\NetCheat.ini"
LoadSett
IPBox.Text = ip_addr
uleBox.Text = alt_boot
SetLengths
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StartGame.ForeColor = &HC0C000
    LogClear.ForeColor = &HC0C000
    ConPS2.ForeColor = &HC0C000
    DisconnectBtn.ForeColor = &HC0C000
    SendButt.ForeColor = &HC0C000
    Bootule.ForeColor = &HC0C000
    StopDisc.ForeColor = &HC0C000
End Sub

Private Sub JokerTB_Click()
    JokerThat.Show
End Sub

Private Sub LogClear_Click()
    LogBox.Text = ""
End Sub

Private Sub LogClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LogClear.ForeColor = &HFFFF00
End Sub

Private Sub mnu_cbc_Click()
    cbcUtil.Show
End Sub

Private Sub SendButt_Click()
    Dim Size As String, GetString As String, Codes As String, Old As String, tempArr() As String, MastCodeLine As Long
    Dim GetByte As Integer, RetSend As Integer
    
    If WSock1.State = 7 Then
    
    'Find mastercode and move it to the top. If there is no mastercode, append a "00000000 00000000" to the top
    Codes = FindMC(CodeWindow.Text & vbCrLf)
    
    Codes = StringFlip(ParseCodes(Codes))
    Old = Codes
    Codes = HexToString(Codes)
    Size = Len(Codes)
    'Size = SaveAsBin(Codes, App.Path & "\temp.bin")
    
'    Dim File As String
'    File = (App.Path & "\temp.bin")
    
'    If Dir(File) = "" Then: MsgBox "Error: temp.bin doesn't exist! Try again.": Exit Sub
'    Open File For Binary As #2
'        For X = 1 To Size
'            Get #2, X, GetByte
'            GetString = GetString & Trim(Hex(GetByte))
'        Next X
'    Close #2
    
    Log "Sending " & Trim(Str(Size / 8)) & " lines of code"
    
    SendRet = SendWait(HexToString("3"), "K")
    
    SendRet = SendWait(Size, "K")
    
    SendRet = SendWait(Codes, "", 1)
    
    'Delay 5
    'Send "K"
    
    Else
        MsgBox "Not Connected!"
    End If
    
End Sub

Private Sub SendButt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SendButt.ForeColor = &HFFFF00
End Sub

Private Sub StartGame_Click()
    Dim SendRet As Integer
    
    If WSock1.State = 7 Then
        SendRet = SendWait(HexToString("1"), "", 1)
        WSock1.Close
    Else
        MsgBox "Please connect to the PS2 first!"
    End If
End Sub

Private Sub StartGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartGame.ForeColor = &HFFFF00
End Sub

Private Sub StopDisc_Click()
Dim SendRet As Integer

If WSock1.State = 7 Then
    SendRet = SendWait(HexToString("5"), "K")
Else
    MsgBox "Not Connected!"
End If
End Sub

Private Sub StopDisc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StopDisc.ForeColor = &HFFFF00
End Sub

Sub Log(Text As String)
    Text = Text & vbCrLf
    If Len(LogBox.Text) > 1000 Then: LogBox.Text = ""
    
    LogBox.Text = LogBox.Text & Text
End Sub

Private Sub WSock1_SendComplete()
    SendC = True
End Sub
