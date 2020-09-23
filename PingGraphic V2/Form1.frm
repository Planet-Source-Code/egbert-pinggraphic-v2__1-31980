VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MultiPingGraphic"
   ClientHeight    =   11430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11670
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11430
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Graphic"
      Height          =   4215
      Left            =   120
      TabIndex        =   52
      Top             =   7080
      Width           =   11415
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   10800
         Top             =   3360
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   255
         LargeChange     =   100
         Left            =   10560
         Max             =   5000
         Min             =   10
         SmallChange     =   10
         TabIndex        =   65
         Top             =   3480
         Value           =   100
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "100"
         Top             =   3480
         Width           =   495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2415
         LargeChange     =   1000
         Left            =   10800
         Max             =   100
         Min             =   10000
         SmallChange     =   100
         TabIndex        =   62
         Top             =   600
         Value           =   500
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Do  not follow"
         Height          =   255
         Index           =   4
         Left            =   6360
         TabIndex        =   59
         Top             =   3480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Follow Yellow"
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   58
         Top             =   3480
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Follow Blue"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   57
         Top             =   3480
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Follow Green"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   56
         Top             =   3480
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Follow Red"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   55
         Top             =   3480
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   400
         Left            =   120
         Max             =   0
         SmallChange     =   10
         TabIndex        =   54
         Top             =   3840
         Width           =   11175
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   3135
         Left            =   120
         ScaleHeight     =   205
         ScaleMode       =   0  'User
         ScaleWidth      =   701
         TabIndex        =   53
         Top             =   240
         Width           =   10575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh interval :"
         Height          =   195
         Left            =   8880
         TabIndex        =   64
         Top             =   3480
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   10800
         TabIndex        =   61
         Top             =   3120
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "1000"
         Height          =   195
         Index           =   0
         Left            =   10800
         TabIndex        =   60
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Yellow Line"
      Height          =   3375
      Index           =   3
      Left            =   5880
      TabIndex        =   39
      Top             =   3600
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "Draw graphic"
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   73
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Option2 
         Caption         =   "FollowLog"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   69
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Index           =   3
         Left            =   240
         Top             =   1200
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   47
         Text            =   "208.246.240.6"
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   46
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Index           =   3
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   44
         Text            =   "5000"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   43
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear line"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   42
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear log"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   41
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   40
         Text            =   "1000"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address :"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   51
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log :"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   50
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TimeOut Time (MS) :"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   49
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interval (MS) :"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   48
         Top             =   2880
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Blue Line"
      Height          =   3375
      Index           =   2
      Left            =   120
      TabIndex        =   26
      Top             =   3600
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "Draw graphic"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   72
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Option2 
         Caption         =   "FollowLog"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   68
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Index           =   2
         Left            =   240
         Top             =   1200
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   34
         Text            =   "208.246.240.6"
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   31
         Text            =   "5000"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear line"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   29
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear log"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   28
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   27
         Text            =   "1000"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address :"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log :"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TimeOut Time (MS) :"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   36
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interval (MS) :"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   35
         Top             =   2880
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Green Line"
      Height          =   3375
      Index           =   1
      Left            =   5880
      TabIndex        =   13
      Top             =   120
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "Draw graphic"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   71
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Option2 
         Caption         =   "FollowLog"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   67
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Index           =   1
         Left            =   240
         Top             =   1200
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   21
         Text            =   "208.246.240.6"
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   18
         Text            =   "5000"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear line"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   16
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear log"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   15
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Text            =   "1000"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address :"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log :"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TimeOut Time (MS) :"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interval (MS) :"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   22
         Top             =   2880
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Red Line"
      Height          =   3375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "Draw graphic"
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   70
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Option2 
         Caption         =   "FollowLog"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   66
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Index           =   0
         Left            =   240
         Top             =   1200
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   11
         Text            =   "1000"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear log"
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear line"
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Text            =   "5000"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Text            =   "208.246.240.6"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interval (MS) :"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   2880
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TimeOut Time (MS) :"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log :"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address :"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' BE-Softwaredeveloper ©
' Written by egbert (Egberttheone@hotmail.com)
' HTTP:\\www.besoftwaredeveloper.tk

Dim WithEvents PingProtocol1 As ClsPing
Attribute PingProtocol1.VB_VarHelpID = -1
Dim WithEvents PingProtocol2 As ClsPing
Attribute PingProtocol2.VB_VarHelpID = -1
Dim WithEvents PingProtocol3 As ClsPing
Attribute PingProtocol3.VB_VarHelpID = -1
Dim WithEvents PingProtocol4 As ClsPing
Attribute PingProtocol4.VB_VarHelpID = -1
Dim LineData1() As Long
Dim LineData2() As Long
Dim LineData3() As Long
Dim LineData4() As Long

Private Sub Command1_Click(Index As Integer)
If Not IsNumeric(Text2(Index)) Or Val(Text2(Index)) < 100 Or Val(Text2(Index)) > 10000 Then MsgBox "Enter a valid timeout time (between 100 and 10000)", vbExclamation + vbOKOnly, "Error": Text2(Index).SetFocus: Exit Sub
If Not IsNumeric(Text3(Index)) Or Val(Text3(Index)) < 10 Or Val(Text3(Index)) > 10000 Then MsgBox "Enter a valid interval time (between 10 and 10000)", vbExclamation + vbOKOnly, "Error": Text3(Index).SetFocus: Exit Sub

Select Case Index
Case 0
PingProtocol1.IP = Text1(Index)
PingProtocol1.TimeOut = Val(Text2(Index))
Case 1
PingProtocol2.IP = Text1(Index)
PingProtocol2.TimeOut = Val(Text2(Index))
Case 2
PingProtocol3.IP = Text1(Index)
PingProtocol3.TimeOut = Val(Text2(Index))
Case 3
PingProtocol4.IP = Text1(Index)
PingProtocol4.TimeOut = Val(Text2(Index))
End Select

Timer1(Index).Interval = Val(Text3(Index))
Timer1(Index).Enabled = True
Command2(Index).Enabled = True
Command1(Index).Enabled = False
End Sub

Private Sub Command2_Click(Index As Integer)
Timer1(Index).Enabled = False
Command2(Index).Enabled = False
Command1(Index).Enabled = True
End Sub

Private Sub Command3_Click(Index As Integer)
Select Case Index
Case 0
Erase LineData1
ReDim LineData1(0 To 0)
Case 1
Erase LineData2
ReDim LineData2(0 To 0)
Case 2
Erase LineData3
ReDim LineData3(0 To 0)
Case 3
Erase LineData4
ReDim LineData4(0 To 0)
End Select
Command2_Click Index
CalcScroll
End Sub

Private Sub Command4_Click(Index As Integer)
List1(Index).Clear
End Sub

Private Sub Form_Load()
Set PingProtocol1 = New ClsPing
Set PingProtocol2 = New ClsPing
Set PingProtocol3 = New ClsPing
Set PingProtocol4 = New ClsPing
ReDim LineData1(0 To 0)
ReDim LineData2(0 To 0)
ReDim LineData3(0 To 0)
ReDim LineData4(0 To 0)
VScroll1_Change
CalcScroll
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set PingProtocol1 = Nothing
Set PingProtocol2 = Nothing
Set PingProtocol3 = Nothing
Set PingProtocol4 = Nothing
End Sub

Private Sub HScroll1_Change()
'Option1_Click 4
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).Value = True And Index <> 0 Then Option1(0).Value = False
If Option1(1).Value = True And Index <> 1 Then Option1(1).Value = False
If Option1(2).Value = True And Index <> 2 Then Option1(2).Value = False
If Option1(3).Value = True And Index <> 3 Then Option1(3).Value = False
If Option1(4).Value = True And Index <> 4 Then Option1(4).Value = False

Option1(Index).Value = True
End Sub

Private Sub PingProtocol1_RecievedEcho(Ping As Long, TimedOut As Boolean)
ReDim Preserve LineData1(0 To UBound(LineData1) + 1)
If TimedOut = False Then
List1(0).AddItem "Echo recieved at " & Time & " Ping : " & Ping
Else
List1(0).AddItem "No echo recieved at " & Time & " before timeout"
End If

LineData1(UBound(LineData1)) = Ping
If Option2(0).Value = 1 Then List1(0).ListIndex = List1(0).ListCount - 1
If Command2(0).Enabled = True Then Timer1(0).Enabled = True
CalcScroll
End Sub

Private Function PaintLines()
Dim A As Long, B As Long

Picture1.Cls

If Check1(0).Value = 1 Then
B = -1
For A = HScroll1.Value + 1 To HScroll1.Value + Picture1.ScaleWidth ' draw Red line
If A <= UBound(LineData1) Then
If B = -1 Then B = LineData1(A)
Picture1.Line (A - HScroll1.Value, Picture1.ScaleHeight - B)-(A - HScroll1.Value + 1, Picture1.ScaleHeight - LineData1(A)), vbRed
B = LineData1(A)
End If
Next A
End If

If Check1(1).Value = 1 Then
B = -1
For A = HScroll1.Value + 1 To HScroll1.Value + Picture1.ScaleWidth ' draw Green line
If A <= UBound(LineData2) Then
If B = -1 Then B = LineData2(A)
Picture1.Line (A - HScroll1.Value, Picture1.ScaleHeight - B)-(A - HScroll1.Value + 1, Picture1.ScaleHeight - LineData2(A)), vbGreen
B = LineData2(A)
End If
Next A
End If

If Check1(2).Value = 1 Then
B = -1
For A = HScroll1.Value + 1 To HScroll1.Value + Picture1.ScaleWidth ' draw Blue line
If A <= UBound(LineData3) Then
If B = -1 Then B = LineData3(A)
Picture1.Line (A - HScroll1.Value, Picture1.ScaleHeight - B)-(A - HScroll1.Value + 1, Picture1.ScaleHeight - LineData3(A)), vbBlue
B = LineData3(A)
End If
Next A
End If

If Check1(3).Value = 1 Then
B = -1
For A = HScroll1.Value + 1 To HScroll1.Value + Picture1.ScaleWidth ' draw Yellow line
If A <= UBound(LineData4) Then
If B = -1 Then B = LineData4(A)
Picture1.Line (A - HScroll1.Value, Picture1.ScaleHeight - B)-(A - HScroll1.Value + 1, Picture1.ScaleHeight - LineData4(A)), vbYellow
B = LineData4(A)
End If
Next A
End If


' follow the line's
If Option1(0).Value = True Then
If UBound(LineData1) > Picture1.ScaleWidth Then
HScroll1.Value = UBound(LineData1) - Picture1.ScaleWidth
End If
End If

If Option1(1).Value = True Then
If UBound(LineData2) > Picture1.ScaleWidth Then
HScroll1.Value = UBound(LineData2) - Picture1.ScaleWidth
End If
End If

If Option1(2).Value = True Then
If UBound(LineData3) > Picture1.ScaleWidth Then
HScroll1.Value = UBound(LineData3) - Picture1.ScaleWidth
End If
End If

If Option1(3).Value = True Then
If UBound(LineData4) > Picture1.ScaleWidth Then
HScroll1.Value = UBound(LineData4) - Picture1.ScaleWidth
End If
End If
End Function

Private Sub PingProtocol2_RecievedEcho(Ping As Long, TimedOut As Boolean)
ReDim Preserve LineData2(0 To UBound(LineData2) + 1)
If TimedOut = False Then
List1(1).AddItem "Echo recieved at " & Time & " Ping : " & Ping
Else
List1(1).AddItem "No echo recieved at " & Time & " before timeout"
End If

LineData2(UBound(LineData2)) = Ping
If Option2(1).Value = 1 Then List1(1).ListIndex = List1(1).ListCount - 1
If Command2(1).Enabled = True Then Timer1(1).Enabled = True
CalcScroll
End Sub

Private Sub PingProtocol3_RecievedEcho(Ping As Long, TimedOut As Boolean)
ReDim Preserve LineData3(0 To UBound(LineData3) + 1)
If TimedOut = False Then
List1(2).AddItem "Echo recieved at " & Time & " Ping : " & Ping
Else
List1(2).AddItem "No echo recieved at " & Time & " before timeout"
End If

LineData3(UBound(LineData3)) = Ping
If Option2(2).Value = 1 Then List1(2).ListIndex = List1(2).ListCount - 1
If Command2(2).Enabled = True Then Timer1(2).Enabled = True
CalcScroll
End Sub

Private Sub PingProtocol4_RecievedEcho(Ping As Long, TimedOut As Boolean)
ReDim Preserve LineData4(0 To UBound(LineData4) + 1)
If TimedOut = False Then
List1(3).AddItem "Echo recieved at " & Time & " Ping : " & Ping
Else
List1(3).AddItem "No echo recieved at " & Time & " before timeout"
End If

LineData4(UBound(LineData4)) = Ping
If Option2(3).Value = 1 Then List1(3).ListIndex = List1(3).ListCount - 1
If Command2(3).Enabled = True Then Timer1(3).Enabled = True
CalcScroll
End Sub

Private Sub Timer1_Timer(Index As Integer)
Timer1(Index).Enabled = False

Select Case Index
Case 0
PingProtocol1.Ping
Case 1
PingProtocol2.Ping
Case 2
PingProtocol3.Ping
Case 3
PingProtocol4.Ping
End Select
End Sub

Private Sub Timer2_Timer()
PaintLines
End Sub

Private Sub VScroll1_Change()
Label5(0) = VScroll1.Value
Picture1.ScaleHeight = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
Text4 = VScroll2
Timer2.Interval = Text4
End Sub

Private Function CalcScroll()
Dim A As Long

If UBound(LineData1) > A Then A = UBound(LineData1)
If UBound(LineData2) > A Then A = UBound(LineData2)
If UBound(LineData3) > A Then A = UBound(LineData3)
If UBound(LineData4) > A Then A = UBound(LineData4)
If A > Picture1.ScaleWidth Then HScroll1.Max = A - Picture1.ScaleWidth Else HScroll1.Max = 0
End Function
