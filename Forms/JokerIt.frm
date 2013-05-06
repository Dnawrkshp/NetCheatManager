VERSION 5.00
Begin VB.Form JokerIt 
   Caption         =   "Joker It by Dnawrkshp - Based off of Joker It! by Rathlar"
   ClientHeight    =   5865
   ClientLeft      =   210
   ClientTop       =   3105
   ClientWidth     =   8595
   Icon            =   "JokerIt.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "JokerIt.frx":08CA
   ScaleHeight     =   5865
   ScaleWidth      =   8595
   Begin VB.CheckBox RmvFmt 
      BackColor       =   &H00242410&
      Caption         =   "Remove Comments"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   5040
      Width           =   2415
   End
   Begin VB.OptionButton Width 
      BackColor       =   &H0044431A&
      Caption         =   "16 bit"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   1
      Left            =   4455
      TabIndex        =   22
      Top             =   4230
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton Width 
      BackColor       =   &H0044431A&
      Caption         =   "8 bit"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   21
      Top             =   4260
      Width           =   735
   End
   Begin VB.ComboBox TypeCombo 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   315
      ItemData        =   "JokerIt.frx":A5A14
      Left            =   3120
      List            =   "JokerIt.frx":A5A16
      TabIndex        =   20
      Text            =   "Equal"
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Jokered 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   4935
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox JAddr 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Text            =   "Joker Address"
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox Codes 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "JokerIt.frx":A5A18
      Top             =   840
      Width           =   2895
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00555620&
      Caption         =   "Square"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   15
      Left            =   4440
      TabIndex        =   15
      Top             =   3780
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00555620&
      Caption         =   "Cross"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   14
      Left            =   3120
      TabIndex        =   14
      Top             =   3780
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H006B6C28&
      Caption         =   "Circle"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   13
      Left            =   4440
      TabIndex        =   13
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H006B6C28&
      Caption         =   "Triangle"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   12
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H007F7D2E&
      Caption         =   "R1"
      Height          =   255
      Index           =   11
      Left            =   4440
      TabIndex        =   11
      Top             =   2820
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H007F7D2E&
      Caption         =   "L1"
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   10
      Top             =   2820
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00919434&
      Caption         =   "R2"
      Height          =   255
      Index           =   9
      Left            =   4440
      TabIndex        =   9
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00919434&
      Caption         =   "L2"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   8
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00A8A13C&
      Caption         =   "Left"
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   7
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00A8A13C&
      Caption         =   "Down"
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   6
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00B8BB41&
      Caption         =   "Right"
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   5
      Top             =   1380
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00B8BB41&
      Caption         =   "Up"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   4
      Top             =   1380
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00CACD47&
      Caption         =   "Start"
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      Top             =   900
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00CACD47&
      Caption         =   "R3"
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   900
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00E7EC51&
      Caption         =   "L3"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   1
      Top             =   450
      Width           =   1095
   End
   Begin VB.CheckBox BCheck 
      BackColor       =   &H00E7EC51&
      Caption         =   "Select"
      Height          =   255
      Index           =   0
      Left            =   3150
      TabIndex        =   0
      Top             =   450
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Jokered code"
      Height          =   255
      Left            =   5640
      TabIndex        =   29
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label JokLab 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Joker It"
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
      Height          =   495
      Left            =   840
      TabIndex        =   28
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label InsLab 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Insert"
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
      Height          =   495
      Left            =   5760
      TabIndex        =   27
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label ExitLab 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
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
      Height          =   495
      Left            =   7320
      TabIndex        =   26
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codes to Joker"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   25
      Top             =   630
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Joker Address"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   135
      TabIndex        =   24
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Based off of Joker It by Rathlar"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   16
      Top             =   5520
      Width           =   2295
   End
End
Attribute VB_Name = "JokerIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WidthSel As Long

Const bE As String = "1110" 'Binary of 0xE
Const bD As String = "1101" 'Binary of 0xD

Private Sub Codes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 65 And Shift = 2 Then
    Codes.SelStart = 0
    Codes.SelLength = Len(Codes.Text)
End If
End Sub

Private Sub Codes_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then
    KeyAscii = 0
End If
End Sub

Private Sub ExitLab_Click()
Unload Me
End Sub

Private Sub ExitLab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ExitLab.ForeColor = &HFFFF00
End Sub

Private Sub Form_Load()
WidthSel = 1

'Add type options
TypeCombo.Clear
TypeCombo.AddItem "Equal", 0
TypeCombo.AddItem "Not Equal", 1
TypeCombo.AddItem "Less Than or Equal", 2
TypeCombo.AddItem "Greater Than or Equal", 3
TypeCombo.AddItem "Mask Unset", 4
TypeCombo.AddItem "Mask Set", 5
TypeCombo.AddItem "(Mask Unset) - COMBO", 6
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ExitLab.ForeColor = &HC0C000
InsLab.ForeColor = &HC0C000
JokLab.ForeColor = &HC0C000
End Sub

Private Sub InsLab_Click()
SendButt.CodeWindow.Text = SendButt.CodeWindow.Text & Jokered.Text
End Sub

Private Sub InsLab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
InsLab.ForeColor = &HFFFF00
End Sub

Private Sub Jokered_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 65 And Shift = 2 Then
    Jokered.SelStart = 0
    Jokered.SelLength = Len(Jokered.Text)
End If
End Sub

Private Sub Jokered_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then
    KeyAscii = 0
End If
End Sub

Function GetSizeOf(Of As String)
    GetSizeOf = Len(Of) / 16
End Function

Private Sub JokLab_Click()
Dim Size As Integer, Code As String, Com As String, CompType As Integer, JokVal As String, Width As Integer
Dim Addr As String, Val As String

Width = Abs(WidthSel - 1)
Code = ParseCodes(Codes.Text & vbCrLf)
Size = GetSizeOf(Code)
JokVal = GetJokVal
CompType = TypeCombo.ListIndex
If CompType < 0 Then: CompType = 0

Dim temp1 As String, temp2 As String, temp3 As String, temp4 As String

If Size > 1 And Len(JAddr.Text) = 8 Then
    temp1 = Pad(Right(Hex2Bin(Right(JAddr.Text, 7)), 25), 28)
    temp2 = Trim(Str(Width))
    temp3 = Pad(Dec2Bin(Trim(Str(CompType))), 3)
    
    Addr = bE & Pad(temp2, 4) & Pad(Dec2Bin(Trim(Str(Size))), 8)
    Addr = Hex(Bin2Dec(Left(Addr, 16))) & Right(JokVal, 4)
    Val = Pad(temp3, 4) & temp1
    Val = Hex(Bin2Dec(Left(Val, 16))) & Hex(Bin2Dec(Right(Val, 16)))
ElseIf Size = 1 Then
    temp1 = Pad(Right(Hex2Bin(Right(JAddr.Text, 7)), 25), 28)
    temp2 = Trim(Str(Width))
    temp3 = Pad(Dec2Bin(Trim(Str(CompType))), 3)
    
    Addr = bD & temp1
    Addr = Hex(Bin2Dec(Left(Addr, 16))) & Hex(Bin2Dec(Right(Addr, 16)))
    Val = "00000000" & Pad(temp3, 4) & Pad(temp2, 4)
    Val = Hex(BinToDec(Left(Val, 16))) & Right(JokVal, 4)
Else
    MsgBox "Invalid parameters!"
    Exit Sub
End If

If RmvFmt.Value = 0 Then
    temp1 = Codes.Text & vbCrLf
Else
    temp1 = RemoveComments(Codes.Text & vbCrLf)
End If
Jokered.Text = Pad(Addr, 8) & " " & Pad(Val, 8) & vbCrLf & temp1
End Sub

Private Sub JokLab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
JokLab.ForeColor = &HFFFF00
End Sub

Private Sub Width_Click(Index As Integer)
WidthSel = Index
End Sub

Function GetJokVal()
    Dim temp As String, X As Long
    temp = ""
    
    For X = 0 To 15
        If BCheck(X).Value = 0 Then
            temp = "0" & temp 'Button is off
        Else
            temp = "1" & temp 'Button is on
        End If
    Next X
    
    X = BinToDec(temp)
    If TypeCombo.ListIndex < 4 Then: X = -1 - X
    GetJokVal = Pad(Hex(X), 4)
End Function

