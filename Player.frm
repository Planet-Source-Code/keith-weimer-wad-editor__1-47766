VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl mciPlayer 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      UpdateInterval  =   1
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSComctlLib.Slider sldPosition 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      TickFrequency   =   15000
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim UpdatePosition As Boolean

Property Get Position() As Long
    Position = mciPlayer.Position
End Property

Property Let Position(Position As Long)
    If mciPlayer.Mode <> mciModeNotOpen Then
        If mciPlayer.Mode = mciModePlay Then
            mciPlayer.From = Position
            mciPlayer.Command = "Play"
        Else
            mciPlayer.To = Position
            mciPlayer.Command = "Seek"
        End If
    End If
End Property

Property Get Length() As Long
    Length = mciPlayer.Length
End Property

Sub OpenDevice()
    mciPlayer.Command = "Open"
    
    If mciPlayer.Mode <> mciModeNotOpen Then
        sldPosition.Max = Length
        sldPosition.Value = Position
        sldPosition.Enabled = True
    End If
End Sub

Sub CloseDevice()
    If mciPlayer.Mode <> mciModeNotOpen Then
        mciPlayer.Command = "Stop"
        mciPlayer.Command = "Close"
        
        sldPosition.Enabled = False
        sldPosition.Value = 0
        sldPosition.Max = 1
    End If
End Sub

Sub OpenFile(ByVal FileName As String)
    CloseDevice
    mciPlayer.FileName = FileName
    OpenDevice
End Sub

Private Sub Form_Load()
    mciPlayer.Wait = True
    UpdatePosition = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseDevice
End Sub

Private Sub mciPlayer_Done(NotifyCode As Integer)
    If NotifyCode = mciNotifySuccessful Then
        mciPlayer.Command = "Stop"
        Position = 0
    End If
End Sub

Private Sub mciPlayer_StatusUpdate()
    If UpdatePosition Then sldPosition.Value = Position
End Sub

Private Sub sldPosition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UpdatePosition = False
End Sub

Private Sub sldPosition_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Position = sldPosition.Value
    UpdatePosition = True
End Sub
