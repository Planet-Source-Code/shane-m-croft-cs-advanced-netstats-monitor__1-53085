Attribute VB_Name = "ModGetStack"
Option Explicit

Type MIB_UDPROW
    dwLocalAddr As String * 4 'address on local computer
    dwLocalPort As String * 4 'port number on local computer
End Type

'Private Type MIB_TCPROW
'    dwState As Long
'    dwLocalAddr As Long
'    dwLocalPort As Long
'    dwRemoteAddr As Long
'    dwRemotePort As Long
'End Type

Type MIB_TCPROW
    dwState As Long        'state of the connection
    dwLocalAddr As String * 4    'address on local computer
    dwLocalPort As String * 4    'port number on local computer
    dwRemoteAddr As String * 4   'address on remote computer
    dwRemotePort As String * 4   'port number on remote computer
End Type

Private Declare Function GetProcessHeap Lib "kernel32" () As Long


Private Declare Function htons Lib "ws2_32.dll" (ByVal dwLong As Long) As Long


Public Declare Function AllocateAndGetTcpExTableFromStack Lib "iphlpapi.dll" (pTcpTableEx As Any, ByVal bOrder As Long, ByVal heap As Long, ByVal zero As Long, ByVal Flags As Long) As Long

Public Declare Function AllocateAndGetUdpExTableFromStack Lib "iphlpapi.dll" (pUdpTableEx As Any, ByVal bOrder As Long, ByVal heap As Long, ByVal zero As Long, ByVal Flags As Long) As Long


Private Declare Function SetTcpEntry Lib "iphlpapi.dll" (pTcpTableEx As MIB_TCPROW) As Long


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private pTablePtr As Long
    Private pDataRef As Long
    Public nRows As Long
    Public oRows As Long
    Private nCurrentRow As Long
    Private udtRow As MIB_TCPROW
    Private nState As Long
    Private nLocalAddr As Long
    Private nLocalPort As Long
    Private nLocalAddr2 As Long
    Private nLocalPort2 As Long
    Private nRemoteAddr As Long
    Private nRemotePort As Long
    Private nProcId As Long
    Private nProcName As String
    Public nRet As Long

Private Type Connection_
    FileName As String
    ProcessID As Long
    ProcessName As String
    LocalPort As Long
    RemotePort As Long
    LocalHost As Long
    RemoteHost As Long
    State As String
End Type

Public Connection(2000) As Connection_



Public Function GetIPAddress(dwAddr As Long) As String
    Dim arrIpParts(3) As Byte
    CopyMemory arrIpParts(0), dwAddr, 4
    GetIPAddress = CStr(arrIpParts(0)) & "." & _
    CStr(arrIpParts(1)) & "." & _
    CStr(arrIpParts(2)) & "." & _
    CStr(arrIpParts(3))
End Function


Public Function GetPort(ByVal dwPort As Long) As Long
    GetPort = htons(dwPort)
End Function

Public Sub RefreshStack()
Dim i As Long
    pDataRef = 0

For i = 0 To nRows '// read 24 bytes at a time
    CopyMemory nState, ByVal pTablePtr + (pDataRef + 4), 4
    CopyMemory nLocalAddr, ByVal pTablePtr + (pDataRef + 8), 4
    CopyMemory nLocalPort, ByVal pTablePtr + (pDataRef + 12), 4
    CopyMemory nRemoteAddr, ByVal pTablePtr + (pDataRef + 16), 4
    CopyMemory nRemotePort, ByVal pTablePtr + (pDataRef + 20), 4
    CopyMemory nProcId, ByVal pTablePtr + (pDataRef + 24), 4

    DoEvents

        'If nRemoteAddr <> 0 Or nRemotePort <> 0 Or nLocalPort <> 0 Then
            Connection(i).State = nState
            Connection(i).LocalHost = nLocalAddr
            Connection(i).LocalPort = nLocalPort
            Connection(i).RemoteHost = nRemoteAddr
            Connection(i).RemotePort = nRemotePort
            Connection(i).ProcessID = nProcId
            Connection(i).ProcessName = GetProcessName(nProcId)
        'End If
        pDataRef = pDataRef + 24

        DoEvents
        Next i

End Sub

Public Function GetEntryCount() As Long
    GetEntryCount = nRows - 1 '// The last entry is always an EOF of sorts
End Function
Public Function TerminateThisConnection(rNum As Long) As Boolean
    
    On Error GoTo ErrorTrap
    
    udtRow.dwLocalAddr = Connection(rNum).LocalHost
    udtRow.dwLocalPort = Connection(rNum).LocalPort
    udtRow.dwRemoteAddr = Connection(rNum).RemoteHost
    udtRow.dwRemotePort = Connection(rNum).RemotePort
    udtRow.dwState = 12
    SetTcpEntry udtRow
    
    TerminateThisConnection = True
    Exit Function
    
ErrorTrap:
    TerminateThisConnection = False
End Function

Public Function GetRefreshTCP() As Boolean
    nRet = AllocateAndGetTcpExTableFromStack(pTablePtr, 0, GetProcessHeap, 0, 2)
    
    If nRet = 0 Then
        CopyMemory nRows, ByVal pTablePtr, 4
    Else
        GetRefreshTCP = False
        Exit Function
    End If

    If nRows = 0 Or pTablePtr = 0 Then
    GetRefreshTCP = False
    Exit Function
    End If

   If oRows <> nRows Then
   GetRefreshTCP = True
   Else
   GetRefreshTCP = False
   End If
   
oRows = nRows

End Function
Function c_state(s) As String
  Select Case s
  Case "0": c_state = "UNKNOWN"
  Case "1": c_state = "CLOSED"
  Case "2": c_state = "LISTENING"
  Case "3": c_state = "SYN_SENT"
  Case "4": c_state = "SYN_RCVD"
  Case "5": c_state = "ESTABLISHED"
  Case "6": c_state = "FIN_WAIT1"
  Case "7": c_state = "FIN_WAIT2"
  Case "8": c_state = "CLOSE_WAIT"
  Case "9": c_state = "CLOSING"
  Case "10": c_state = "LAST_ACK"
  Case "11": c_state = "TIME_WAIT"
  Case "12": c_state = "DELETE_TCB"
  End Select
End Function
