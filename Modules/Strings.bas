Attribute VB_Name = "Strings"
'String manipulation module for NetCheat Manager by Dnawrkshp

'Written by ORCXodus
Function StrToHex(Data As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim I         As Long
    For I = 1 To Len(Data)
        strTemp = Hex$(Asc(Mid$(Data, I, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & strTemp 'Space$(1) & strTemp
    Next I
    StrToHex = strReturn
End Function

'Written by ORCXodus
Function HexToString(ByVal HexToStr As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim I         As Long
    For I = 1 To Len(HexToStr) Step 2
        strTemp = Chr$(Val("&H" & Mid$(HexToStr, I, 2)))
        strReturn = strReturn & strTemp
    Next I
    HexToString = strReturn
End Function

'Written by Dnawrkshp
Function Pad(String1 As String, Size As Integer)

Pad = String1
If Len(String1) >= Size Then
Exit Function

Else
Dim X As Integer
X = Len(String1)

Do While X < Size
    Pad = "0" & Pad
    X = X + 1
Loop

End If


End Function

'Written by Dnawrkshp
Function StringFlip(Bytes As String)
StringFlip = ""
Dim Counter As Integer
Counter = 1

Do While Counter <= Len(Bytes)

StringFlip = StringFlip & Mid(Bytes, Counter + 6, 2)
StringFlip = StringFlip & Mid(Bytes, Counter + 4, 2)
StringFlip = StringFlip & Mid(Bytes, Counter + 2, 2)
StringFlip = StringFlip & Mid(Bytes, Counter, 2)

Counter = Counter + 8
Loop

End Function

Function ParseCodes(Codes As String)
    ParseCodes = Replace(Replace(Codes, " ", ""), vbCrLf, "")
End Function

Function RemoveComments(Text As String)
    Dim temp As String, X As Long, Y As Long
    temp = Text
    X = 1
    Do While X > 0 And X <= Len(temp)
        X = InStr(temp, "//")
        If X = 0 Then: Exit Do
        
        Y = InStr(X, temp, vbCrLf)
        
        If X = 1 Then
            temp = Right(temp, Len(temp) - Y - 1)
        Else
            temp = Left(temp, X - 2) + Right(temp, Len(temp) - Y)
        End If
    Loop
    X = 1
    Do While X > 0 And X <= Len(temp)
        X = InStr(temp, "/*")
        If X = 0 Then: Exit Do
        
        Y = InStr(X, temp, "*/")
        temp = Left(temp, X - 1) + Right(temp, Len(temp) - Y - 1)
        X = InStr(temp, "/*")
    Loop

    RemoveComments = temp
End Function

Function FindMC(Code As String)
    Dim tempArr() As String, X As Long, Y As Long
    Code = RemoveComments(Code)
    
    tempArr = Split(Code, vbCrLf)
    For X = 0 To UBound(tempArr)
        
        If InStr(1, tempArr(X), "9") = 1 Then
            SendButt.Log "Mastercode: " & tempArr(X)
            Code = "0" & Right(tempArr(X), Len(tempArr(X)) - 1)
            
            For Y = 0 To UBound(tempArr)
                If Y <> X Then
                    Code = Code & tempArr(Y)
                End If
            Next Y
            
            Exit For
        ElseIf X = UBound(tempArr) Then
            SendButt.Log "No mastercode found"
            Code = "0000000000000000"
            
            For Y = 0 To UBound(tempArr)
                Code = Code & tempArr(Y)
            Next Y
            
        End If
    Next X
    
    FindMC = Code
End Function
