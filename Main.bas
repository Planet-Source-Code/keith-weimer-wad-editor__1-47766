Attribute VB_Name = "modMain"
Option Explicit

Public FileSystem As FileSystemObject

Sub Main()
    Set FileSystem = New FileSystemObject
    
    frmMain.Show
End Sub
