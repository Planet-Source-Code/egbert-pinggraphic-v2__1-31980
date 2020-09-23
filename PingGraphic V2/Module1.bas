Attribute VB_Name = "TimerModule"
' BE-Softwaredeveloper ©
' Written by egbert (Egberttheone@hotmail.com)
' HTTP:\\www.besoftwaredeveloper.tk

Public Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Type TypeIds
InUse As Boolean
Object As ClsPing
End Type

Dim IDs() As TypeIds
Dim First As Boolean

Public Sub CallBack(ByVal hWnd As Long, ByVal uMsg As Long, ByVal IDEvent As Long, ByVal dwTime As Long)
IDs(IDEvent).Object.TimerUpdate
End Sub

Public Function Add_Timer_To_Class(Object As ClsPing, Interval As Long) As Long
Dim A As Long

If Not First Then ReDim IDs(1 To 1): First = True

For A = 1 To UBound(IDs)
If Not IDs(A).InUse Then
SetTimer Form1.hWnd, A, Interval, AddressOf CallBack
Add_Timer_To_Class = A
IDs(A).InUse = True
Set IDs(A).Object = Object
Exit Function
End If
Next A

' if not engouht timers are availeble then add one!
ReDim Preserve IDs(1 To UBound(IDs) + 1)
SetTimer Form1.hWnd, UBound(IDs), Interval, AddressOf CallBack
Add_Timer_To_Class = UBound(IDs)
IDs(UBound(IDs)).InUse = True
Set IDs(UBound(IDs)).Object = Object
End Function

Public Function Remove_Timer_From_Class(ID As Long)
KillTimer Form1.hWnd, ID
End Function
