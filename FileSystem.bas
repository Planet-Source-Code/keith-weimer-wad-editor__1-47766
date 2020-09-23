Attribute VB_Name = "modFileSystem"
Option Explicit
Option Compare Text

'Handling for File System

'Written by Keith R. Weimer
'Way Too Happy Software

'Include: Memory.bas

Enum SlashEnum
    BackSlash = 1
    ForwardSlash = 2
End Enum

Public Declare Function GetWindowsDirectoryA Lib "kernel32.dll" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectoryA Lib "kernel32.dll" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPathA Lib "kernel32.dll" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function GetShortPathNameA Lib "kernel32.dll" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetLongPathNameA Lib "kernel32.dll" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetFullPathNameA Lib "kernel32.dll" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

Function AddSlash(ByVal Text As String, Optional ByVal Slash As SlashEnum = BackSlash) As String
    Dim SlashChar As String
    
    AddSlash = Text
    SlashChar = Choose(Slash, "\", "/")
    
    If Right$(Text, 1) <> SlashChar Then AddSlash = AddSlash & SlashChar
End Function

Function RemoveSlash(ByVal Text As String, Optional ByVal Slash As SlashEnum) As String
    Dim SlashChar As String
    
    RemoveSlash = Text
    SlashChar = Right$(Text, 1)
    
    If SlashChar = "\" Or SlashChar = "/" Then
        If (SlashChar = "\" And Slash = ForwardSlash) Or (SlashChar = "/" And Slash = BackSlash) Then Exit Function
        
        RemoveSlash = Left$(RemoveSlash, Len(RemoveSlash) - 1)
    End If
End Function

'%AppPath% - Application path
'%TempPath% - Temp path
'%WinPath% - Windows path
'%SysPath% - System path
'%CurPath% - Current path

Function GetFakePath(ByVal Path As String, Optional ByVal UseCurrentPath As Boolean = True) As String
    GetFakePath = GetFullPathName(Path)
    GetFakePath = Replace(GetFakePath, RemoveSlash(App.Path), "%AppPath%")
    GetFakePath = Replace(GetFakePath, GetTempPath, "%TempPath%")
    GetFakePath = Replace(GetFakePath, GetWindowsDirectory, "%WinPath%")
    GetFakePath = Replace(GetFakePath, GetSystemDirectory, "%SysPath%")
    If UseCurrentPath Then GetFakePath = Replace(GetFakePath, RemoveSlash(CurDir), "%CurPath%")
End Function

Function GetRealPath(ByVal Path As String) As String
    GetRealPath = Path
    GetRealPath = Replace(GetRealPath, "%AppPath%", RemoveSlash(App.Path))
    GetRealPath = Replace(GetRealPath, "%TempPath%", GetTempPath)
    GetRealPath = Replace(GetRealPath, "%WinPath%", GetWindowsDirectory)
    GetRealPath = Replace(GetRealPath, "%SysPath%", GetSystemDirectory)
    GetRealPath = Replace(GetRealPath, "%CurPath%", RemoveSlash(CurDir))
    GetRealPath = GetFullPathName(GetRealPath)
End Function

Function GetEXEFileName() As String
    GetEXEFileName = AddSlash(App.Path) & App.EXEName & ".exe"
End Function

Function GetWindowsDirectory() As String
    Dim Length As Long
    
    Length = GetWindowsDirectoryA(GetWindowsDirectory, 0)
    If Length <> 0 Then
        GetWindowsDirectory = String$(Length, 0)
        GetWindowsDirectoryA GetWindowsDirectory, Length
        GetWindowsDirectory = RemoveSlash(Left$(GetWindowsDirectory, Len(GetWindowsDirectory) - 1))
    End If
End Function

Function GetSystemDirectory() As String
    Dim Length As Long
    
    Length = GetSystemDirectoryA(GetSystemDirectory, 0)
    If Length <> 0 Then
        GetSystemDirectory = String$(Length, 0)
        GetSystemDirectoryA GetSystemDirectory, Length
        GetSystemDirectory = RemoveSlash(Left$(GetSystemDirectory, Len(GetSystemDirectory) - 1))
    End If
End Function

Function GetTempPath() As String
    Dim Length As Long
    
    Length = GetTempPathA(0, GetTempPath)
    If Length <> 0 Then
        GetTempPath = String$(Length, 0)
        GetTempPathA Length, GetTempPath
        GetTempPath = RemoveSlash(Left$(GetTempPath, Len(GetTempPath) - 1))
    End If
End Function

Function GetShortPathName(ByVal Path As String) As String
    Dim Temp As String
    Dim Length As Long
    
    Length = GetShortPathNameA(Path, Temp, 0)
    If Length <> 0 Then
        GetShortPathName = String$(Length, 0)
        GetShortPathNameA Path, GetShortPathName, Length
        GetShortPathName = Left$(GetShortPathName, Len(GetShortPathName) - 1)
    End If
End Function

Function GetLongPathName(ByVal Path As String) As String
    Dim Length As Long
    
    Length = GetLongPathNameA(Path, GetLongPathName, 0)
    If Length <> 0 Then
        GetLongPathName = String$(Length, 0)
        GetLongPathNameA Path, GetLongPathName, Length
        GetLongPathName = Left$(GetLongPathName, Len(GetLongPathName) - 1)
    End If
End Function

Function GetFullPathName(ByVal Path As String) As String
    Dim Length As Long
    
    Length = GetFullPathNameA(Path, 0, GetFullPathName, vbNullString)
    If Length <> 0 Then
        GetFullPathName = String$(Length, 0)
        GetFullPathNameA Path, Length, GetFullPathName, vbNullString
        GetFullPathName = Left$(GetFullPathName, Len(GetFullPathName) - 1)
    End If
End Function

Function GetDriveName(ByVal Path As String) As String
    Dim Start As Long
    
    Start = InStr(1, Path, ":")
    If Start <= 3 Then GetDriveName = Left$(Path, Start)
End Function

Function GetParentFolderName(ByVal Path As String)
    Dim Start As Long
    
    Start = InStrRev(Path, "\")
    If Start = 0 Then Start = InStrRev(Path, "/")
    
    If Start <> 0 Then GetParentFolderName = Left$(Path, Start - 1)
End Function

Function GetFileName(ByVal Path As String) As String
    'On Error Resume Next
    
    Dim Start As Long
    
    Start = InStrRev(Path, "\")
    If Start = 0 Then Start = InStrRev(Path, "/")
    
    If Start = 0 Then
        GetFileName = Path
    Else
        GetFileName = Mid$(Path, Start + 1)
    End If
End Function

Function GetExtensionName(ByVal Path As String) As String
    'On Error Resume Next
    
    Dim FileName As String
    
    FileName = GetFileName(Path)
    If FileName <> Empty Then
        Dim Start As Long
        
        Start = InStrRev(FileName, ".")
        If Start <> 0 Then GetExtensionName = Mid$(FileName, Start + 1)
    End If
End Function

Function GetBaseName(ByVal Path As String) As String
    'On Error Resume Next
    
    Dim FileName As String
    
    FileName = GetFileName(Path)
    If FileName <> Empty Then
        Dim Start As Long
        
        Start = InStrRev(FileName, ".")
        If Start <> 0 Then GetBaseName = Left$(FileName, Start - 1)
    End If
End Function

Sub FileMove(ByVal Source As String, ByVal Destination As String)
    'On Error Resume Next
    
    FileCopy Source, Destination
    If Err.Number = 0 Then Kill Source
End Sub

Function IsFolder(ByVal Path As String) As Boolean
    'On Error Resume Next
    
    If PathExists(Path) Then IsFolder = GetAttr(Path) And vbDirectory
End Function

Function IsFile(ByVal Path As String) As Boolean
    'On Error Resume Next
    
    IsFile = Not IsFolder(Path)
End Function

Function PathExists(ByVal Path As String) As Boolean
    'On Error Resume Next
    
    If Path <> Empty Then PathExists = Dir(Path, vbReadOnly Or vbHidden Or vbSystem Or vbDirectory) <> Empty
End Function

Function FolderExists(ByVal Path As String) As Boolean
    'On Error Resume Next
    
    If Path <> Empty Then
        If Dir(Path, vbDirectory) <> Empty Then
            FolderExists = IsFolder(Path)
        End If
    End If
End Function

Function FileExists(ByVal Path As String) As Boolean
    'On Error Resume Next
    
    If Path <> Empty Then FileExists = Dir(Path, vbReadOnly Or vbHidden Or vbSystem) <> Empty
End Function

Function ReadFile(ByVal FileName As String, Optional ByVal BufferSize As Long) As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    Open FileName For Binary Access Read As #FileNum
        If BufferSize <= 0 Then
            ReadFile = Input$(LOF(FileNum), FileNum)
        Else
            Do Until EOF(FileNum)
                ReadFile = ReadFile & Input$(BufferSize, FileNum)
                DoEvents
            Loop
        End If
    Close #FileNum
End Function

Sub WriteFile(ByVal FileName As String, Data As String, Optional ByVal BufferSize As Long)
    Dim FileNum As Integer
    
    If FileExists(FileName) Then Kill FileName
    
    FileNum = FreeFile
    Open FileName For Binary Access Write As #FileNum
        If BufferSize = 0 Then
            Put #FileNum, , Data
        Else
            Dim Start As Long
            
            For Start = 1 To Len(Data) Step BufferSize
                Put #FileNum, , Mid$(Data, Start, BufferSize)
            Next Start
        End If
    Close #FileNum
End Sub
