VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Callback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Fail As Boolean

Event Progress(ByVal Value As Long, ByVal Total As Long)
Event StatusText(ByVal Text As String)

Function Progress(ByVal Value As Long, ByVal Total As Long)
    RaiseEvent Progress(Value, Total)
End Function

Function StatusText(ByVal Text As String)
    RaiseEvent StatusText(Text)
End Function
