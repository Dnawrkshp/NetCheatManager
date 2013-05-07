Attribute VB_Name = "Networking"
Option Explicit
Private Declare Function GetRTTAndHopCount Lib "iphlpapi.dll" (ByVal lDestIPAddr As Long, _
                                                               ByRef lHopCount As Long, _
                                                               ByVal lMaxHops As Long, _
                                                               ByRef lRTT As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

Public SendC As Boolean 'Set to yes when sendingis complete
Const SUCCESS  As Integer = 1

'Connects to the PS2
Sub ConToPS2(IP As String)
    Dim Timer As Long
    
    If IP = "" Then: SendButt.Log "Please enter an IP Address!": Exit Sub
    
    If SendButt.WSock1.State = 7 Then: SendButt.Log "Already connected": Exit Sub
    
    SendButt.WSock1.Close
    SendButt.WSock1.RemoteHost = IP
    SendButt.WSock1.RemotePort = 2345
    SendButt.Log "Connecting"
    DoEvents
    If Ping(IP, 1) Then
        SendButt.WSock1.Connect
        Do While SendButt.WSock1.State = 6
            DoEvents
            Timer = Timer + 1
            If Timer > 4000000 Then: SendButt.Log "Connection timeout": SendButt.WSock1.Close: Exit Sub
        Loop
        If SendButt.WSock1.State = 7 Then
            SendButt.Log "Connected!"
            SaveSett 0
        Else
            If SendButt.WSock1.State = 9 Then
                SendButt.Log "An error occured"
            Else
                SendButt.Log "Failed to Connect"
            End If
            SendButt.WSock1.Close
        End If
    Else
        SendButt.Log "Failed to Connect"
    End If
End Sub

'From http://www.vbforums.com/showthread.php?384015-winsock-timeout-Resolved
Function Ping(sIPadr As String, Pings As Long) As Boolean
 
Dim lIPadr     As Long
Dim lHopsCount As Long
Dim lRTT       As Long
Dim lMaxHops   As Long
 
 
    lMaxHops = Pings
    lIPadr = inet_addr(sIPadr)
    Ping = (GetRTTAndHopCount(lIPadr, lHopsCount, lMaxHops, lRTT) = SUCCESS)
 
End Function

'Keeps the PS2 and PC in a form of sync
Sub Delay(Y As Long)
    Dim X As Long
    For X = 0 To 10000 * Y
    
    Next X
    DoEvents
End Sub

Function WaitForReply(Old As String)
    Dim temp As String
    SendButt.WSock1.GetData (temp)
    While temp <> Old
        
        Delay 1
        SendButt.WSock1.GetData (temp)
        DoEvents
    Wend
End Function

Function SendWait(Send As String, Wait As String, Optional Mode As Integer)
    If SendButt.WSock1.State <> 7 Then: Exit Function
    
    Dim temp As String
    SendButt.WSock1.SendData Send
    
    While Not SendC
    DoEvents
    Wend
    
    SendC = False
    
    'Delay 1
    SendButt.WSock1.GetData temp
    If Wait <> "" Then
        'temp = Left(temp, Len(Wait))
        Do While InStr(1, temp, Wait) = 0 'temp <> Wait 'Wait for specific response
            
            Delay 1
            If SendButt.WSock1.State <> 7 Then: Exit Function
            SendButt.WSock1.GetData temp
            'temp = Left(temp, Len(Wait))
        Loop
    Else
        Do While temp = Wait 'Wait for any respone
            Delay 1
            SendButt.WSock1.GetData temp
        Loop
    End If
    
    If Mode = 1 Then: SendButt.Log temp
    
    Delay 1000 'Delay. The PS2 lags behind otherwise...
    
End Function

Function SendArrayWait(Send() As String, Wait As String, Size As Long, Optional Mode As Integer)
    If SendButt.WSock1.State <> 7 Then: Exit Function
    
    Dim temp As String, X As Long
    
    For X = 0 To Size
        SendButt.WSock1.SendData Send(X)
        While Not SendC
            DoEvents
        Wend
        SendC = False
    Next X
    
    SendButt.WSock1.GetData temp
    If Wait <> "" Then
        'temp = Left(temp, Len(Wait))
        Do While InStr(1, temp, Wait) = 0 'Wait for specific response
            
            Delay 1
            SendButt.WSock1.GetData temp
            'temp = Left(temp, Len(Wait))
        Loop
    Else
        Do While temp = "" 'Wait for any respone
            
            If SendButt.WSock1.State <> 7 Then: Exit Function
            Delay 1
            SendButt.WSock1.GetData temp
        Loop
    End If
    
    If Mode = 1 Then: SendButt.Log temp
    
    Delay 1000 'Delay. The PS2 lags behind otherwise...
End Function

Sub Send(Send As String)
    SendButt.WSock1.SendData Send
    Wait
    SendC = False
End Sub

Sub Wait()
    Dim X As Long
  While Not SendC
    If X > 10000000 Then
        SendC = True
        MsgBox "Sending timed out!"
    End If
    X = X + 1
    DoEvents
  Wend
End Sub
