VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' BE-Softwaredeveloper ©
' Written by egbert (Egberttheone@hotmail.com)
' HTTP:\\www.besoftwaredeveloper.tk

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim WithEvents Sock As Winsock
Attribute Sock.VB_VarHelpID = -1
Dim MyID As Long
Dim Start As Long
Dim CurrentPingReturn As Long
Dim Pinging As Boolean
Dim RTimeOut As Long
Event RecievedEcho(Ping As Long, TimedOut As Boolean)

Private Sub Class_Initialize()
Set Sock = New Winsock
Sock.Protocol = sckUDPProtocol
Sock.RemotePort = 7
RTimeOut = 10000 ' 10 sec
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
If GetTickCount - Start > RTimeOut Then
CurrentPingReturn = RTimeOut
Remove_Timer_From_Class MyID
Pinging = False
RaiseEvent RecievedEcho(RTimeOut, True)
Else
CurrentPingReturn = GetTickCount - Start
Remove_Timer_From_Class MyID
Pinging = False
RaiseEvent RecievedEcho(Me.CurrentPing, False)
End If
End Sub

Public Function Ping()
If Not Pinging Then
Start = GetTickCount
Sock.SendData "/Echo/"
Pinging = True
MyID = Add_Timer_To_Class(Me, RTimeOut)
End If
End Function

Property Get CurrentPing() As Long
CurrentPing = CurrentPingReturn
End Property

Property Get IP() As String
IP = Sock.RemoteHost
End Property

Property Let IP(NewIP As String)
Sock.RemoteHost = NewIP
End Property

Property Get TimeOut() As Long
TimeOut = RTimeOut
End Property

Property Let TimeOut(NewTime As Long)
RTimeOut = NewTime
End Property

Public Function TimerUpdate()
If GetTickCount - Start > RTimeOut Then
Sock_DataArrival 1
End If
End Function
