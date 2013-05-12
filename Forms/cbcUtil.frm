VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form cbcUtil 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CBC Extractor By Xodus - cb2util by misfire"
   ClientHeight    =   6330
   ClientLeft      =   3630
   ClientTop       =   3540
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "cbcUtil.frx":0000
   ScaleHeight     =   6330
   ScaleWidth      =   8550
   Begin VB.CommandButton Command 
      Caption         =   "Copy to NetCheat"
      Height          =   330
      Left            =   6345
      TabIndex        =   10
      Top             =   5895
      Width           =   2085
   End
   Begin VB.FileListBox File2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   2235
      Left            =   180
      TabIndex        =   8
      Top             =   3555
      Width           =   4110
   End
   Begin VB.ListBox List1 
      Height          =   5715
      Left            =   11280
      TabIndex        =   7
      Top             =   180
      Width           =   2715
   End
   Begin VB.CheckBox crypt 
      BackColor       =   &H007F7D2E&
      Caption         =   "Decrypt"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   2820
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.OptionButton encr 
      BackColor       =   &H007F7D2E&
      Caption         =   "V8+"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   5
      Top             =   2820
      Width           =   675
   End
   Begin VB.OptionButton encr 
      BackColor       =   &H007F7D2E&
      Caption         =   "V7"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   2820
      Value           =   -1  'True
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Unpack CBC"
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   1275
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   2235
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Width           =   4110
   End
   Begin RichTextLib.RichTextBox cbcbox 
      Height          =   5700
      Left            =   4365
      TabIndex        =   0
      Top             =   90
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   10054
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"cbcUtil.frx":A514A
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Text files / Unpacked cbc files (click to view)"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   3195
      Width           =   4110
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "CBC files"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   4065
   End
End
Attribute VB_Name = "cbcUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()
    SendButt.CodeWindow.text = SendButt.CodeWindow.text & vbCrLf & cbcbox.text
End Sub

Private Sub Command1_Click()

Dim sh As Long

Dim inputfile As String
Dim outputfile As String
Dim extract As String
Dim version As String
Dim plug As String
Dim p1, p2
'MsgBox CurDir

cbcbox.text = ""
plug = "cb2util "
inputfile = File1.filename
If InStr(1, inputfile, " ") <> 0 Then MsgBox "Cbc file name must not containt spaces. Rename file and try again": Exit Sub
p1 = Left(inputfile, Len(inputfile) - 4)



outputfile = p1 & ".txt"
If crypt.Value = 1 Then extract = "-d "
If encr(0).Value = True Then version = "cbc -7 "
If encr(1).Value = True Then version = "cbc "

Dim cbfn As String
cbfn = plug & version & extract & inputfile & " >" & outputfile
Open "ex.bat" For Output As #1
Print #1, cbfn
Close


sh = Shell("ex.bat", vbHide)
Dim cnt
tryagain:
    DoEvents
    On Error GoTo nope
    Open outputfile For Input As #1
    Close
    
    List1.AddItem "waiting"  'DoEvents
    File2.Refresh
    Exit Sub
    
nope:
    Resume tryagain
    




End Sub

Function fileexists(filename As String) As Boolean
On Error GoTo NOFILEYET
    Dim FL
    FL = FileLen(filename)
    fileexists = True
    Exit Function
NOFILEYET:
    fileexists = False

End Function


Private Sub Dir1_Change()
File1.Path = dir1.Path
File2.Path = dir1.Path
End Sub

Private Sub Drive1_Change()
dir1.Path = Drive1.Drive

End Sub

Private Sub File2_Click()
    cbcbox.LoadFile File2.filename
    cbcbox.SelStart = 0
    cbcbox.SelLength = Len(cbcbox.text)
    cbcbox.SelColor = &HE7EC51
End Sub

Private Sub Form_Load()
File1.Path = App.Path
File1.Pattern = "*.cbc"
File2.Path = App.Path
File2.Pattern = "*.txt"
cbcbox.SelStart = 0
cbcbox.SelLength = Len(cbcbox.text)
cbcbox.SelColor = &HE7EC51
End Sub
