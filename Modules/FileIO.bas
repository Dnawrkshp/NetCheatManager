Attribute VB_Name = "FileIO"
'FileIO module for NetCheat manager by Dnawrkshp

Global ip_addr As String
Global alt_boot As String
Global sfile As String

Function SaveAsBin(CodeAsHex As String, FileName As String) 'Saves codes into bin file (with correct LE format) and returns size
    Dim Counter As Long, Size As Long, SubCode As String
    Size = Len(CodeAsHex) + 1
    Counter = 1
    
    If Dir(FileName) <> "" Then: Kill FileName
    
    Open FileName For Binary As #1
        Do While Counter < Size
            SubCode = StringFlip(Mid(CodeAsHex, Counter, 4))
            Put #1, Counter, SubCode
        Counter = Counter + 4
        Loop
    Close #1
    
    SaveAsBin = FileLen(FileName)
End Function

Sub LoadSett()
    If Dir(sfile) = "" Or sfile = "" Then
        Open sfile For Output As #1
            Print #1, "ip_addr = IP Address" & vbCrLf & "alt_boot = uLaunchELF Path"
        Close #1
        
        ip_addr = "IP Address"
        alt_boot = "uLaunchELF Path"
        Exit Sub
    Else
        Dim temp As String, X As Long, Y As Long
        Open sfile For Input As #1
            Do Until EOF(1)
                Line Input #1, temp
                If Y = 0 Then
                    X = InStrRev(temp, "ip_addr = ")
                    ip_addr = Right(temp, Len(temp) - Len("ip_addr = ") + X - 1)
                    Y = Y + 1
                ElseIf Y = 1 Then
                    X = InStrRev(temp, "alt_boot = ")
                    alt_boot = Right(temp, Len(temp) - Len("alt_boot = ") + X - 1)
                    Y = Y + 1
                End If
            Loop
        Close #1
        
        Exit Sub
    End If
    
End Sub

Sub SaveSett(Mode As Integer)
    Dim temp As String, Y As Long
    Dim oldip As String, oldboot As String
    
    If Mode > 1 Then: Exit Sub
    
    oldboot = SendButt.uleBox.Text
    oldip = SendButt.IPBox.Text
    LoadSett
    
    
    Select Case Mode
        Case 0
            ip_addr = "ip_addr = " & oldip
            alt_boot = "alt_boot = " & alt_boot
            GoTo SaveSettMode
        Case 1
            ip_addr = "ip_addr = " & ip_addr
            alt_boot = "alt_boot = " & oldboot
            GoTo SaveSettMode
    End Select
SaveSettMode:
    
    Open sfile For Output As #1
            Print #1, ip_addr
            Print #1, alt_boot
    Close #1
    
    ip_addr = oldip
    alt_boot = oldboot
    
End Sub
