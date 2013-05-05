Attribute VB_Name = "Math"
'By ORCXodus

'---------------------------------------------------------------------------------------
' Module: Math
' DateTime: 9/21/2012 12:10:21 AM
' Author: Xodus
' Description: Module to handle mathmatics calculations
'---------------------------------------------------------------------------------------
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'declarations for mathmatical formula variables
Public MaxHexInput As Long
Public MaxDecInput As Long
Public MaxBinInput As Long

Public DecBytes() As Byte
Public MaxDecLen As Long
Public HexString$
Public DecString$
Public BinString$    ' General

Public HexString1$, DecString1$, BinString1$ ' 1st number
Public HexString2$, DecString2$, BinString2$ ' 2nd number
Public HexResult$, DecResult$, BinResult$    ' Logic result
Public HexRemain$, DecRemain$, BinRemain$    ' Div remainder

Public HexBytes() As Byte
Public MaxHexLen As Long
Public Dec1()
Public Dec2()
Public PDec1() As Byte
Public BinBytes() As Byte
Public MaxByteLen As Long
Public BinBits() As Byte
Public MaxBinLen As Long
'public aVBASM As Boolean
Public ptrBinBytes As Long
Public ptrDecBytes As Long
Public ptrHexBytes As Long
Public ptrBinBits As Long
Public Counter() As Byte                     ' Div subtraction counter

Const MaxLong = 10000
Public TLong(MaxLong) As Long
Public LenTLong As Long
'' For machine code if used
' All separated for easier debugging

Public ptrMC As Long
Public ptrMC2 As Long
Public ptrMC3 As Long
Public ptrMC4 As Long
'---------------------------------------------------------------------------------------
' Procedure : Dec2Bin
' Author    : Xodus
' Date      : 9/23/2012
' Purpose   : Converts a decimal value to a binary value
'---------------------------------------------------------------------------------------
Function Dec2Bin(DecValue As String)
method1:
    Dec2B2H (DecValue)
    Dec2Bin = BinString$
method2:
    'Dec2Bin = DecimalToBinary(DecValue)
    
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : Dec2Hex
' Author    : Xodus
' Date      : 9/23/2012
' Purpose   : Converts a decimal value to a hex value
'---------------------------------------------------------------------------------------
Function Dec2Hex(DecValue As String)
   If Val(DecValue) < 1 Then Dec2Hex = Hex(Val(DecValue)): Exit Function
    
    Dec2B2H (DecValue)
    Dec2Hex = HexString$
    Exit Function
End Function
'---------------------------------------------------------------------------------------
' Procedure : Hex2Bin
' Author    : Xodus
' Date      : 9/23/2012
' Purpose   : Converts a hex value to a binary value
'---------------------------------------------------------------------------------------
Function Hex2Bin(HexValue As String)
method1:
    Hex2Bin2Dec (HexValue)
    Hex2Bin = BinString$
   
End Function
'---------------------------------------------------------------------------------------
' Procedure : Hex2Dec
' Author    : Xodus
' Date      : 9/23/2012
' Purpose   : Converts a hex value to a decimal value
'---------------------------------------------------------------------------------------
Function Hex2Dec(HexValue As String, method)
  If HexValue = "00" Then Hex2Dec = 0: Exit Function
  
  Select Case method
    Case 1
        Hex2Bin2Dec (HexValue)
        Hex2Dec = DecString$
    Case 2
        Hex2Dec = HEX2DECIMAL(HexValue)
  End Select
End Function
'---------------------------------------------------------------------------------------
' Procedure : Bin2Dec
' Author    : Xodus
' Date      : 9/23/2012
' Purpose   : Converts binary a value to a decimal value
'---------------------------------------------------------------------------------------
Function Bin2Dec(BinValue As String)
    B2H2Dec (BinValue)
    Bin2Dec = DecString$
End Function
'---------------------------------------------------------------------------------------
' Procedure : B2H
' Author    : Xodus
' Date      : 9/23/2012
' Purpose   : Converts a binary value to a hex value
'---------------------------------------------------------------------------------------
Function B2H(BinValue As String)
    'Debug.Print CurrentAddress
    On Error GoTo method2
    Dim d As Long
    d = BinToDec(BinValue)
    B2H = DECIMAL2HEX(Str(d))
    Exit Function

method2:
    B2H2Dec (BinValue)
    B2H = HexString$
End Function
'---------------------------------------------------------------------------------------
' Procedure : HexUnsigned2Dbl / Mathmatics.bas
' Author    : Xodus
' Date      : 9/23/2012 07:49
' Purpose   : Converts a hex value to a unsigned double decimal value
'---------------------------------------------------------------------------------------
Function HexUnsigned2Dbl(sHex As String) As Double
    Dim I As Long
    Dim lExpCtr As Long
    sHex = "0" + sHex
    For I = Len(sHex) To 2 Step -2
        lExpCtr = lExpCtr + 1
        HexUnsigned2Dbl = HexUnsigned2Dbl + CDbl("&H" + Right$("00" + Mid$(sHex, I - 1, 2), 2)) * 256# ^ (lExpCtr - 1)
    Next I
End Function

'====================================================================================
'====================================================================================
Private Sub Hex2Bin2Dec(a$)
' IN:  A$ = Hex string
' OUT: DecString$, BinString$
Dim k As Long, J As Long
Dim b As Byte
   b = 48
   FillMemory HexBytes(1), MaxHexLen, b    ' "0" to HexBytes()
   ' Fill HexBytes() from A$
   ' NB HexBytes(1) is Right char of A$ ie @ Len(A$)
   ' ie CopyMemory cannot be used here
   For k = 1 To Len(a$)
      HexBytes(k) = Asc(Mid$(a$, Len(a$) - (k - 1), 1))
   Next k
    
   Hex2Bytes a$
   a$ = ""
   Bytes2Bits   ' BinBits()
   Bytes2Dec    ' DecBytes()   ' Slow in VB
   
   ' Get Dec result
   DecString$ = ""
   For k = MaxDecLen To 1 Step -1
      If DecBytes(k) <> 48 Then Exit For
   Next k
   For J = k To 1 Step -1
         DecString$ = DecString$ + Chr$(DecBytes(J))
   Next J
   
   ' Get Binary result
   BinString$ = ""
   For k = MaxBinLen To 1 Step -1
      If BinBits(k) <> 48 Then Exit For
   Next k
   For J = k To 1 Step -1
         BinString$ = BinString$ + Chr$(BinBits(J))
   Next J
End Sub

Private Sub Dec2B2H(a$)
'IN:  A$ = Dec string
'OUT: HexString$, BinString$
Dim k As Long, J As Long
Dim b As Byte
Dim U As Long
   b = 48
   FillMemory DecBytes(1), MaxDecLen, b    ' "0" to DecBytes()
   ' Fill DecBytes() from A$ (DecString$)
   U = UBound(DecBytes(), 1)
   For k = 1 To Len(a$)
      DecBytes(k) = Asc(Mid$(a$, Len(a$) - (k - 1), 1))
   Next k
   a$ = ""
    
   Dec2Bytes
   Bytes2Bits   ' BinBits()
   Bytes2Hex    ' HexBytes()   ' Bit slow in VB
   
   ' Get Hex result
   HexString$ = ""
   For k = MaxHexLen To 1 Step -1
      If HexBytes(k) <> 48 Then Exit For
   Next k
   For J = k To 1 Step -1
         HexString$ = HexString$ + Chr$(HexBytes(J))
   Next J
   
   ' Get Binary result
   BinString$ = ""
   For k = MaxBinLen To 1 Step -1
      If BinBits(k) <> 48 Then Exit For
   Next k
   For J = k To 1 Step -1
         BinString$ = BinString$ + Chr$(BinBits(J))
   Next J
End Sub

Private Sub B2H2Dec(a$)
'IN:  A$ = Bin string
'OUT: HexString$, DecString$
Dim k As Long, J As Long
   ' Zero BinBits
   ReDim BinBits(MaxBinLen)  ' to zero
   
   ' Fill BinBits() from A$
   k = Len(a$)
   For k = 1 To Len(a$)
      BinBits(k) = Asc(Mid$(a$, Len(a$) - (k - 1), 1))
   Next k
   a$ = ""
   
   Bits2Bytes
   Bytes2Hex    ' HexBytes()    ' Bit slow in VB
   Bytes2Dec    ' DecBytes()    ' Very Slow in VB
   
   ' Get Hex result
   HexString$ = ""
   For k = MaxHexLen To 1 Step -1
      If HexBytes(k) <> 48 Then Exit For
   Next k
   For J = k To 1 Step -1
         HexString$ = HexString$ + Chr$(HexBytes(J))
   Next J
   
   ' Get Dec result
   DecString$ = ""
   For k = MaxDecLen To 1 Step -1
      If DecBytes(k) <> 48 Then Exit For
   Next k
   For J = k To 1 Step -1
         DecString$ = DecString$ + Chr$(DecBytes(J))
   Next J
End Sub

'################################################################################

Private Sub Hex2Bytes(HexString$)
'IN:  HexString$
'OUT: BinBytes(MaxByteLen)
Dim a$
Dim LengthHexStr As Long
Dim k As Long, N As Long
Dim b As Byte
   
   ' Ensure LengthHexStr even
   If (Len(HexString$) And 1) <> 0 Then HexString$ = "0" & HexString$
   LengthHexStr = Len(HexString$)
   b = 48
   FillMemory DecBytes(1), MaxDecLen, b    ' "0" to DecBytes()
   ReDim BinBytes(MaxByteLen)  ' to zero
   ' Transfer 2 nybble values to BinBytes()
   N = 1
   For k = LengthHexStr To 2 Step -2
      a$ = Mid$(HexString$, (k - 1), 2)
      BinBytes(N) = Val("&H" & a$)
      N = N + 1
      If N > MaxByteLen Then Exit For
   Next k
End Sub

Private Sub Bits2Bytes()
'IN:  BinBits(MaxBinLen)
'OUT: BinBytes(MaxByteLen)
Dim I As Long, k As Long
Dim sum As Byte, bit As Byte
Dim Carry As Byte
   ReDim BinBytes(MaxByteLen)  ' to zero
   For I = 1 To MaxByteLen
      sum = 0
      Carry = 0
      For k = 8 To 1 Step -1
         bit = BinBits(k + (I - 1) * 8)
         If (bit And 1) <> 0 Then Carry = 1
         sum = sum * 2 + Carry
         Carry = 0
      Next k
      BinBytes(I) = sum
   Next I
End Sub

Private Sub Bytes2Bits()
'IN:  BinBytes(MaxByteLen)
'OUT: BinBits(MaxBinLen)
Dim I As Long, J As Long, k As Long
Dim one As Byte
Dim b As Byte
   b = 48
   FillMemory BinBits(1), MaxBinLen, b    ' "0" to BinBits()
   
   
   ' VB routine
   one = 49 ' "1"
   I = 1
   For J = 1 To MaxByteLen
      b = BinBytes(J)
      For k = 0 To 7
         If (b And 1) <> 0 Then BinBits(I) = one
         b = b \ 2
         I = I + 1
      Next k
   Next J
End Sub

Private Sub Bytes2Hex()
' Bit slow in VB
'IN:  BinBytes(MaxByteLen)
'OUT: HexBytes(MaxHexLen)
Dim I As Long, k As Long
Dim b As Byte, lo As Byte, hi As Byte
   b = 48
   FillMemory HexBytes(1), MaxHexLen, b    ' "0" to HexBytes()
  
   ' VB routine
   I = 1
   For k = 1 To MaxByteLen
      b = BinBytes(k)
      lo = b And &HF
      hi = (b And &HF0) \ 16
      lo = lo + 48
      hi = hi + 48
      If lo > 57 Then lo = lo + 7
      If hi > 57 Then hi = hi + 7
      HexBytes(I) = lo
      HexBytes(I + 1) = hi
      I = I + 2
   Next k
End Sub

Private Sub Dec2Bytes()
' SLOW IN VB!
'IN:  DecBytes(MaxDecLen) from Dec String
'OUT: BinBytes(MaxByteLen)
Dim I As Long, J As Long
Dim byt As Long, lo As Long, hi As Long
Dim ival As Long
   
   ReDim BinBytes(MaxByteLen)  ' to zero

   ' VB routine
   For I = MaxDecLen - 1 To 1 Step -1
      byt = 0
      For J = 1 To MaxByteLen
         ival = 10 * BinBytes(J)
         lo = ival And &HFF
         hi = (ival And &HFF00) \ 256
         lo = lo + byt
         If lo > 255 Then
            lo = lo - 256
            hi = hi + 1
         End If
         BinBytes(J) = CByte(lo)
         byt = hi
      Next J
      J = 1
      ival = DecBytes(I)
      ival = ival - 48  ' 0 - 9
      ival = BinBytes(J) + ival
      If ival > 255 Then
         ival = ival - 256
         byt = 1
         BinBytes(J) = CByte(ival)
         Do
            J = J + 1
            ival = 1& * BinBytes(J) + byt '1
            If ival > 255 Then
               ival = ival - 256
               byt = 1
            Else
               byt = 0
            End If
            BinBytes(J) = CByte(ival)
         Loop While byt = 1
      Else
         byt = 0
         BinBytes(J) = CByte(ival)
      End If
   Next I
End Sub

Private Sub Bytes2Dec()
' SLOW IN VB!
'IN:  BinBytes(MaxByteLen)
'OUT: DecBytes(MaxDecLen)
Dim I As Long, k As Long
Dim Carry1 As Byte, Carry2 As Byte
Dim bits As Integer, sum As Integer
Dim b As Byte
   b = 48
   ReDim DecBytes(MaxDecLen + 4)
   FillMemory DecBytes(1), MaxDecLen, b    ' "0" to DecBytes()
   
   ' VB routine
   k = 1
   Do
XX:
      bits = MaxBinLen
      sum = 0
      Do Until bits = 0
         bits = bits - 1
         Carry1 = 0
         ' Shift bits to left with carry
         For I = 1 To MaxByteLen
            ' Check if * 2 will give a carry
            If BinBytes(I) > 127 Then
               Carry2 = 1
               BinBytes(I) = BinBytes(I) - 128
            Else
               Carry2 = 0
            End If
            
            BinBytes(I) = BinBytes(I) * 2 + Carry1  ' Shift << 1 + 1/0
            Carry1 = Carry2
         Next I
         sum = sum * 2 + Carry1    ' Shift << 1 + 1/0
         If sum >= 10 Then
            sum = sum - 10
            BinBytes(1) = BinBytes(1) + 1
         End If
      Loop
      DecBytes(k) = sum + 48  ' Store ASCII digit
      k = k + 1
      ' Check if finished
      For I = MaxByteLen To 1 Step -1
         If BinBytes(I) <> 0 Then GoTo XX ' GoTo used for comparison with ASM
      Next I
      Exit Do
   Loop

End Sub
Sub SetLengths()
' IN: MaxHexInput
   MaxHexInput = 16
   MaxHexLen = MaxHexInput + 4
   MaxDecLen = MaxHexLen + (MaxHexLen \ 5) + 1
   MaxBinLen = 4 * MaxHexLen
   MaxBinLen = 8 * (MaxBinLen \ 8)  ' Make multiple of 8
   ReDim HexBytes(MaxHexLen)
   ReDim DecBytes(MaxDecLen)
   ReDim BinBits(MaxBinLen)
   MaxByteLen = MaxBinLen \ 8
   ReDim BinBytes(MaxByteLen)
   MaxDecInput = MaxDecLen - 2
   MaxBinInput = MaxBinLen - 8
End Sub
Function Add2(ByVal a1$, ByVal a2$) As String
' Dec A1$ + Dec A2$
Dim L As Long, k As Long
Dim N1 As Byte, N2 As Byte
Dim PartNum As Byte
Dim Carry As Byte
Dim R1$ ', C2$
Dim aZero As Boolean
      
      If Len(a2$) < Len(a1$) Then
         a2$ = String$(Len(a1$) - Len(a2$), "0") & a2$
      End If
      a1$ = "0" & a1$
      a2$ = "0" & a2$
      L = Len(a1$)
      
      ReDim Dec1(L), Dec2(L)
      
      For k = 1 To L
         Dec1(k) = Val(Mid$(a1$, Len(a1$) - (k - 1), 1))
         Dec2(k) = Val(Mid$(a2$, Len(a2$) - (k - 1), 1))
      Next k
      Carry = 0
      DoEvents ' Help stop ridiculous XP whiteout
      For k = 1 To L                   ' 1 2 3 4
         N1 = Dec1(k)
         N2 = Dec2(k)         ' 0 5 2 3
         PartNum = N1 + N2 + Carry
         Carry = PartNum \ 10
         Dec1(k) = PartNum - Carry * 10
      Next k
      R1$ = ""
      aZero = True
      For k = L To 1 Step -1
         If Dec1(k) <> 0 Then aZero = False
         If Not aZero Then
            R1$ = R1$ & (Dec1(k))
         End If
      Next k
      Add2 = R1$
      R1$ = ""
End Function


Private Function Subtract(ByVal a1$, ByVal a2$) As String
' Dec A1$ - Dec A2$
Dim L As Long, k As Long
Dim N1 As Byte, N2 As Byte
Dim PartNum As Byte
Dim Carry As Byte
Dim R1$ ', C2$
Dim aZero As Boolean
      
      L = Len(a1$)
      
      ReDim Dec1(L), Dec2(L)
      
      If Len(a2$) < Len(a1$) Then
         a2$ = String$(Len(a1$) - Len(a2$), "0") & a2$
      End If
      
      For k = 1 To L
         Dec1(k) = Val(Mid$(a1$, Len(a1$) - (k - 1), 1))
         Dec2(k) = Val(Mid$(a2$, Len(a2$) - (k - 1), 1))
      Next k
      Carry = 0
      DoEvents ' Help stop ridiculous XP whiteout
      For k = 1 To L                   ' 1 2 3 4
         N1 = Dec1(k)
         N2 = Dec2(k) + Carry          ' 0 5 2 3
         Carry = 0
         If N1 >= N2 Then
            PartNum = N1 - N2
         Else
            PartNum = (N1 + 10) - N2
            Carry = 1
         End If
         Dec1(k) = PartNum
      Next k
      R1$ = ""
      aZero = True
      For k = L To 1 Step -1
         If Dec1(k) <> 0 Then aZero = False
         If Not aZero Then
            R1$ = R1$ & (Dec1(k))
         End If
      Next k
      Subtract = R1$
      R1$ = ""
End Function


Function Float2Hex(ByVal TmpFloat As Single) As String

    Dim TmpBytes(0 To 3) As Byte
    Dim TmpSng As Single
    Dim tmpStr As String
    Dim X As Long
    TmpSng = TmpFloat
    Call CopyMemory(ByVal VarPtr(TmpBytes(0)), ByVal VarPtr(TmpSng), 4)

    For X = 3 To 0 Step -1
        If Len(Hex(TmpBytes(X))) = 1 Then tmpStr = tmpStr & "0" & Hex(TmpBytes(X)) Else tmpStr = tmpStr & Hex(TmpBytes(X))
    Next X
    Float2Hex = tmpStr

End Function

Function HEX2DECIMAL(HexVal As String)

    On Error GoTo Err
    Dim H As String
    H = HexVal
    Dim Tmp$
    Dim lo1 As Integer, lo2 As Integer
    Dim hi1 As Long, hi2 As Long
    Const hx = "&H"
    Const BigShift = 65536
    Const LilShift = 256, Two = 2
    Tmp = H
    If UCase(Left$(H, 2)) = "&H" Then Tmp = Mid$(H, 3)
    Tmp = Right$("0000000" & Tmp, 8)
    If IsNumeric(hx & Tmp) Then
        lo1 = CInt(hx & Right$(Tmp, Two))
        hi1 = CLng(hx & Mid$(Tmp, 5, Two))
        lo2 = CInt(hx & Mid$(Tmp, 3, Two))
        hi2 = CLng(hx & Left$(Tmp, Two))
        HEX2DECIMAL = CCur(hi2 * LilShift + lo2) * BigShift + (hi1 * LilShift) + lo1
    End If
    Exit Function
Err:
    MsgBox "There was an error: Hex2Decimal/Math"
    End Function
Function DecimalToBinary(DecVal As String)

On Error Resume Next
Dim BinVal As String, InputVal As Long
Dim Remain As Integer, Output As Long
InputVal = CLng(Val(DecVal))
Output = InputVal
Do
    Remain = Output Mod 2
    Output = Output \ 2
    BinVal = Trim$(Str(Remain) & BinVal)
Loop Until Output = 0

Do While Len(BinVal) < 16
    BinVal = "0" + BinVal

Loop

DecimalToBinary = BinVal


End Function


Function HEX2FLOAT(ByVal tmpHex As String) As Single
    If Left(tmpHex, 1) = "c" Then GoTo skipchk
    If tmpHex > "4fff0000" Then Exit Function
skipchk:
    If InStr(tmpHex, "-") <> 0 Then Exit Function
    On Error Resume Next
    Dim TmpSng As Single
    Dim tmpLng As Long
    tmpLng = CLng("&H" & tmpHex)
    Call CopyMemory(ByVal VarPtr(TmpSng), ByVal VarPtr(tmpLng), 4)
    HEX2FLOAT = TmpSng
End Function

Function DECIMAL2HEX(Data As Long) As String
    On Error GoTo Err
    Dim DECNUM As Long
    DECNUM = Data
    Dim NextHexDigit As Double
    Dim HEXNUM As String
    HEXNUM = ""
    While DECNUM <> 0
        NextHexDigit = DECNUM - (Int(DECNUM / 16) * 16)
        If NextHexDigit < 10 Then
            HEXNUM = Chr(Asc(NextHexDigit)) & HEXNUM
        Else
            HEXNUM = Chr(Asc("A") + NextHexDigit - 10) & HEXNUM
        End If
        DECNUM = Int(DECNUM / 16)
    Wend
    If HEXNUM = "" Then HEXNUM = "0"
    DECIMAL2HEX = HEXNUM
    Exit Function
Err:
    MsgBox "There was an error - Decimal2Hex"

End Function

Function AsciiToHex(Data As String) As String
Dim Ret As String, I As Double

Ret = ""
For I = 1 To Len(Data)
    Ret = Ret & Right("00" & Hex(Asc(Mid(Data, I, 1))), 2)
Next I

AsciiToHex = Ret
End Function

Function HexToAscii(Data As String) As String
Dim Ret As String, I As Double

Ret = ""
For I = 1 To Len(Data)
    Ret = Ret & Chr(Val("&H" & Mid(Data, I, 2)))
    I = I + 1
Next I

HexToAscii = Ret
End Function

Sub HexToStr2(HexToStr As String)
Dim strTemp   As String
Dim strReturn As String
Dim I         As Long
    For I = 1 To Len(HexToStr) Step 3
        strTemp = Chr$(Val("&H" & Mid$(HexToStr, I, 2)))
        strReturn = strReturn & strTemp
    Next I
    HexToStr = strReturn
End Sub

Sub StringToHex(Data As String)
Dim strTemp   As String
Dim strReturn As String
Dim I         As Long
    For I = 1 To Len(Data)
        strTemp = Hex$(Asc(Mid$(Data, I, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & strTemp 'Space$(1) & strTemp
    Next I
    Data = strReturn
End Sub

Function HexToDec(hx As String) As Variant
Dim I As Double, M As String, T As Double, R As Variant

R = 0
For I = 0 To Len(hx) - 1
    M = Mid(hx, Len(hx) - I, 1)
    Select Case LCase(M)
        Case "f"
            T = 15
        Case "e"
            T = 14
        Case "d"
            T = 13
        Case "c"
            T = 12
        Case "b"
            T = 11
        Case "a"
            T = 10
        Case "9"
            T = 9
        Case "8"
            T = 8
        Case "7"
            T = 7
        Case "6"
            T = 6
        Case "5"
            T = 5
        Case "4"
            T = 4
        Case "3"
            T = 3
        Case "2"
            T = 2
        Case "1"
            T = 1
        Case "0"
            T = 0
    End Select
    
    R = CDec(R) + CDec((CDec(T) * (16 ^ CDec(I))))
    
Next I
'Dim abc As Object


'MsgBox abc
HexToDec = CDec(R)
End Function

Function DecConvert(dc As Variant, bits As Variant) As String
Dim Ret As String, Upper As Variant, Lower As Variant, ts() As String, tStr As String

Ret = ""

If CDec(dc) = 1 Then
    DecConvert = "1"
    Exit Function
End If

If CDec(dc) = 0 Then
    DecConvert = "0"
    Exit Function
End If

If CDec(bits) = 16 And CDec(dc) < 16 Then
    If CDec(dc) > 9 Then
        Select Case CDec(dc)
            Case 10
                DecConvert = "A"
            Case 11
                DecConvert = "B"
            Case 12
                DecConvert = "C"
            Case 13
                DecConvert = "D"
            Case 14
                DecConvert = "E"
            Case 15
                DecConvert = "F"
        End Select
        Exit Function
    Else
        DecConvert = CDec(dc)
        Exit Function
    End If
End If


tStr = CDec(CDec(dc) / CDec(bits))
ts = Split(tStr & ".00", ".")
Upper = CDec(ts(0))
Lower = CDec(ts(1))
Do Until CDec(Upper) = 0
    ts = Split(tStr & ".0", ".")
    Upper = CDec(CDec(ts(0)) / CDec(bits))
    Lower = CDec(CDec(("." & ts(1))) * CDec(bits))
    
    tStr = CDec(Upper) & ".0"
    
    Select Case Lower
        Case 10
            Ret = "A" & Ret
        Case 11
            Ret = "B" & Ret
        Case 12
            Ret = "C" & Ret
        Case 13
            Ret = "D" & Ret
        Case 14
            Ret = "E" & Ret
        Case 15
            Ret = "F" & Ret
        Case Else
            Ret = Lower & Ret
    End Select

    
    ts = Split(tStr, ".")
Loop

DecConvert = Ret
End Function

Function BinToDec(Bin As String) As Long
Dim I As Double, Exp As Variant, TotOut As Variant

Exp = CDec(0)
TotOut = CDec(0)
For I = 0 To Len(Bin) - 1
    
    If Mid(Bin, Len(Bin) - I, 1) = "1" Then
        TotOut = CDec(CDec(TotOut) + CDec((2 ^ CDec(Exp))))
    End If
    
    Exp = CDec(CDec(Exp) + 1)
Next I
'BinString$ = TotOut
BinToDec = CDec(TotOut)
End Function

Function DecToFloat32(DecVal As Variant) As String
Dim tSp() As String, tBin1 As String, tBin2 As String, tBin3 As String
Dim tVar1 As Variant, I As Double
Dim Exponent As String, Sign As String, Mantissa As String

If CDec(DecVal) = 0 Then
    DecToFloat32 = "00000000"
    Exit Function
End If

If CDec(DecVal) < 0 Then Sign = "1"
If CDec(DecVal) > 0 Then Sign = "0"

tVar1 = CDec(DecVal)
If CDec(tVar1) < 0 Then tVar1 = CDec(0 - CDec(tVar1))

tSp = Split("0" & DecVal & ".0", ".")

tBin1 = DecConvert(CDec(tSp(0)), 2)

tVar1 = CDec(CDec("." & tSp(1)) * 2)
tSp = Split(CDec(tVar1) & ".0000", ".")
tBin2 = CDec(tSp(0))
Do Until CDec(tSp(1)) = 0
    tVar1 = CDec(CDec("." & tSp(1)) * 2)
    tSp = Split(CDec(tVar1) & ".0000", ".")
    tBin2 = tBin2 & CDec(tSp(0))
    
    If Len(tBin2) = 23 Then GoTo mantissaOverflow
Loop
mantissaOverflow:

'Strip non-significant numbers
Do Until Left(tBin1, 1) = "1"
    If Len(tBin1) = 1 Then GoTo stripNonSigsLeave
    tBin1 = Right(tBin1, Len(tBin1) - 1)
Loop
stripNonSigsLeave:

If tBin1 = "0" Then
    I = -1
    Do Until Mid(tBin2, Abs(I), 1) = "1"
        I = I - 1
    Loop
    Exponent = DecConvert(CDec(127 + I), 2)
    
    Mantissa = Right(tBin2, Len(tBin2) - Abs(I))
Else
    Exponent = DecConvert(CDec(127 + CDec(Len(tBin1) - 1)), 2)
    
    tBin3 = tBin1 & tBin2
    Mantissa = Right(tBin3, Len(tBin3) - 1)
End If

Sign = Right("0" & Sign, 1)
Exponent = Right("00000000" & Exponent, 8)
Mantissa = Left(Mantissa & "00000000000000000000000", 23)
tBin1 = Right("00000000" & DecConvert(CDec(BinToDec(Sign & Exponent & Mantissa)), 16), 8)

DecToFloat32 = tBin1

End Function

Function Float32ToDec(FloatVal As String) As Variant
Dim Sign As String, Exponent As String, Mantissa As String
Dim tBin1 As String, tBin2 As String
Dim tVar1 As Variant, tVar2 As Variant, tVar3 As Variant, tVar4 As Variant

tVar1 = CDec(HexToDec(FloatVal))
tBin1 = Right("00000000000000000000000000000000" & DecConvert(CDec(tVar1), 2), 32)

Sign = Left(tBin1, 1)
tBin1 = Right(tBin1, Len(tBin1) - 1)
Exponent = Left(tBin1, 8)
tBin1 = Right(tBin1, Len(tBin1) - 8)
Mantissa = tBin1

tVar1 = CDec(CDec(BinToDec(Exponent)) - 127)

tBin1 = "1" & Left(Mantissa, Abs(CDec(tVar1)))
tBin2 = Right(Mantissa, Len(Mantissa) - (Len(tBin1) - 1))

tVar2 = CDec(BinToDec(tBin1))
tVar3 = CDec(CDec(BinToDec(tBin2)) * (2 ^ (0 - Len(tBin2))))

tVar4 = CDec(tVar2) + CDec(tVar3)

If Sign = "1" Then tVar4 = CDec(0 - CDec(tVar4))

Float32ToDec = CDec(tVar4)
End Function





Function Hex32AND(hex1 As String, hex2 As String) As String
Dim v1 As Variant, v2 As Variant, v3 As String, v4 As String, v5 As String
Dim t1 As String, t2 As String

t1 = Right("00000000" & hex1, 8)
t2 = Right("00000000" & hex2, 8)

v1 = CDec(HexToDec(t1))
v2 = CDec(HexToDec(t2))
v3 = DecConvert(CDec(v1), 2)
v4 = DecConvert(CDec(v2), 2)
v5 = Right("00000000000000000000000000000000" & bitAND(v3, v4), 32)

Hex32AND = DecConvert(CDec(BinToDec(v5)), 16)
End Function

Function HexAND(hex1 As String, hex2 As String) As String
Dim v1 As Double, v2 As Double, v3 As String, v4 As String, v5 As String
Dim t1 As String, t2 As String, t3 As String, t4 As String, D1 As Boolean, D2 As Boolean

t1 = Right("0000000000000000" & hex1, 16)
t2 = Right("0000000000000000" & hex2, 16)

D1 = False
D2 = False
v5 = ""

t3 = Right(t1, 8)
t4 = Right(t2, 8)

v1 = HexToDec(t3)
v2 = HexToDec(t4)
v3 = DecConvert(v1, 2)
v4 = DecConvert(v2, 2)
v5 = Right("00000000000000000000000000000000" & bitAND(v3, v4), 32)

t3 = Left(t1, 8)
t4 = Left(t2, 8)

v1 = HexToDec(t3)
v2 = HexToDec(t4)
v3 = DecConvert(v1, 2)
v4 = DecConvert(v2, 2)
v5 = v5 & Right("00000000000000000000000000000000" & bitAND(v3, v4), 32)

t3 = Right("00000000" & DecConvert(BinToDec(Left(v5, 32)), 16), 8)
t4 = Right("00000000" & DecConvert(BinToDec(Right(v5, 32)), 16), 8)

HexAND = t4 & t3
End Function

'---------------------------------------------------------------------------------------
' Procedure : HexOR / Mathmatics.bas
' Author    : Xodus
' Date      : 9/27/2012 11:51
' Purpose   :
'---------------------------------------------------------------------------------------

Function HexOR(hex1 As String, hex2 As String) As String
Dim v1 As Double, v2 As Double, v3 As String, v4 As String, v5 As String
Dim t1 As String, t2 As String, t3 As String, t4 As String, D1 As Boolean, D2 As Boolean

t1 = Right("0000000000000000" & hex1, 16)
t2 = Right("0000000000000000" & hex2, 16)

D1 = False
D2 = False
v5 = ""

t3 = Right(t1, 8)
t4 = Right(t2, 8)

v1 = HexToDec(t3)
v2 = HexToDec(t4)
v3 = DecConvert(v1, 2)
v4 = DecConvert(v2, 2)
v5 = Right("00000000000000000000000000000000" & bitOR(v3, v4), 32)

t3 = Left(t1, 8)
t4 = Left(t2, 8)

v1 = HexToDec(t3)
v2 = HexToDec(t4)
v3 = DecConvert(v1, 2)
v4 = DecConvert(v2, 2)
v5 = v5 & Right("00000000000000000000000000000000" & bitOR(v3, v4), 32)

t3 = Right("00000000" & DecConvert(BinToDec(Left(v5, 32)), 16), 8)
t4 = Right("00000000" & DecConvert(BinToDec(Right(v5, 32)), 16), 8)


HexOR = t4 & t3
End Function

'---------------------------------------------------------------------------------------
' Procedure : HexXOR / Mathmatics.bas
' Author    : Xodus
' Date      : 9/27/2012 11:51
' Purpose   :
'---------------------------------------------------------------------------------------

Function HexXOR(hex1 As String, hex2 As String) As String
Dim v1 As Double, v2 As Double, v3 As String, v4 As String, v5 As String
Dim t1 As String, t2 As String, t3 As String, t4 As String, D1 As Boolean, D2 As Boolean

t1 = Right("0000000000000000" & hex1, 16)
t2 = Right("0000000000000000" & hex2, 16)

D1 = False
D2 = False
v5 = ""

t3 = Right(t1, 8)
t4 = Right(t2, 8)

v1 = HexToDec(t3)
v2 = HexToDec(t4)
v3 = DecConvert(v1, 2)
v4 = DecConvert(v2, 2)
v5 = Right("00000000000000000000000000000000" & bitXOR(v3, v4), 32)

t3 = Left(t1, 8)
t4 = Left(t2, 8)

v1 = HexToDec(t3)
v2 = HexToDec(t4)
v3 = DecConvert(v1, 2)
v4 = DecConvert(v2, 2)
v5 = v5 & Right("00000000000000000000000000000000" & bitXOR(v3, v4), 32)

t3 = Right("00000000" & DecConvert(BinToDec(Left(v5, 32)), 16), 8)
t4 = Right("00000000" & DecConvert(BinToDec(Right(v5, 32)), 16), 8)


HexXOR = t4 & t3
End Function

'---------------------------------------------------------------------------------------
' Procedure : bitAND / Mathmatics.bas
' Author    : Xodus
' Date      : 9/27/2012 11:51
' Purpose   :
'---------------------------------------------------------------------------------------

Function bitAND(bin1 As String, bin2 As String) As String
Dim I As Double, bits1 As String, bits2 As String, bits3 As String

If Len(bin1) < Len(bin2) Then
    bits1 = bin2
    bits2 = bin1
Else
    bits1 = bin1
    bits2 = bin2
End If

Do Until Len(bits2) = Len(bits1)
    bits2 = "0" & bits2
Loop

bits3 = ""
For I = 1 To Len(bits1)
    If Mid(bits1, I, 1) = "1" And Mid(bits2, I, 1) = "1" Then
        bits3 = bits3 & "1"
    Else
        bits3 = bits3 & "0"
    End If
Next I

bitAND = bits3
End Function

'---------------------------------------------------------------------------------------
' Procedure : bitXOR / Mathmatics.bas
' Author    : Xodus
' Date      : 9/27/2012 11:51
' Purpose   :
'---------------------------------------------------------------------------------------

Function bitXOR(bin1 As String, bin2 As String) As String
Dim I As Double, bits1 As String, bits2 As String, bits3 As String

If Len(bin1) < Len(bin2) Then
    bits1 = bin2
    bits2 = bin1
Else
    bits1 = bin1
    bits2 = bin2
End If

Do Until Len(bits2) = Len(bits1)
    bits2 = "0" & bits2
Loop

bits3 = ""
For I = 1 To Len(bits1)
    If Mid(bits1, I, 1) = "1" And Mid(bits2, I, 1) = "0" Then
        bits3 = bits3 & "1"
    ElseIf Mid(bits1, I, 1) = "0" And Mid(bits2, I, 1) = "1" Then
        bits3 = bits3 & "1"
    Else
        bits3 = bits3 & "0"
    End If
Next I

bitXOR = bits3
End Function

'---------------------------------------------------------------------------------------
' Procedure : bitOR / Mathmatics.bas
' Author    : Xodus
' Date      : 9/27/2012 11:51
' Purpose   :
'---------------------------------------------------------------------------------------

Function bitOR(bin1 As String, bin2 As String) As String
Dim I As Double, bits1 As String, bits2 As String, bits3 As String

If Len(bin1) < Len(bin2) Then
    bits1 = bin2
    bits2 = bin1
Else
    bits1 = bin1
    bits2 = bin2
End If

Do Until Len(bits2) = Len(bits1)
    bits2 = "0" & bits2
Loop

bits3 = ""
For I = 1 To Len(bits1)
    If Mid(bits1, I, 1) = "0" And Mid(bits2, I, 1) = "0" Then
        bits3 = bits3 & "0"
    Else
        bits3 = bits3 & "1"
    End If
Next I

bitOR = bits3
End Function

'---------------------------------------------------------------------------------------
' Procedure : bitADD / Mathmatics.bas
' Author    : Xodus
' Date      : 9/27/2012 11:51
' Purpose   :
'---------------------------------------------------------------------------------------

Function bitADD(bin1 As String, bin2 As String) As String
    Dim I As Double, bits1 As String, bits2 As String, bits3 As String

    If Len(bin1) < Len(bin2) Then
        bits1 = bin2
        bits2 = bin1
    Else
        bits1 = bin1
        bits2 = bin2
    End If

    Do Until Len(bits2) = Len(bits1)
        bits2 = "0" & bits2
    Loop

    bits3 = ""
    For I = 0 To Len(bits1) - 1
        If Mid(bits1, Len(bits1) - I, 1) = "0" And Mid(bits2, Len(bits2) - I, 1) = "1" Then
            bits3 = "1" & bits3
        ElseIf Mid(bits1, Len(bits1) - I, 1) = "1" And Mid(bits2, Len(bits2) - I, 1) = "0" Then
            bits3 = "1" & bits3
        ElseIf Mid(bits1, Len(bits1) - I, 1) = "1" And Mid(bits2, Len(bits2) - I, 1) = "1" Then
            bits3 = "10" & bits3
        ElseIf Mid(bits1, Len(bits1) - I, 1) = "0" And Mid(bits2, Len(bits2) - I, 1) = "0" Then
            bits3 = "0" & bits3
        End If
    Next I

    bitADD = bits3
End Function

Function neg(Bin As String)
    Dim X As Long, Y As Long
    
    Y = InStr(1, Bin, "1")
    If Y = 0 Then
        neg = "1111111111111111"
        Exit Function
    End If
    neg = Right(Bin, Len(Bin) - Y - 1)
    Do While X <= Y
        neg = "1" & neg
        X = X + 1
    Loop
    
End Function
