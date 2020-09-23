Attribute VB_Name = "modIFF"
Option Explicit
Option Compare Text

'Handling for Interchange Format Files

'Written by Keith R. Weimer
'Way Too Happy Software

'NOTICE: I'm still working out a universal reader for all variations of this file format.
'        Even though there shouldn't be variation in the first place.
'        For the time being lists and cataloges are not supported but should be readable
'        by parsing the data contained within.

'Include: Memory.bas
'Include: Callback.cls

'IFF reference:

'--- MIDI ---
'   MID (MIDI file)
'       UseMicrosoft = False
'       UseMasterChunk = False
'   RMI (RMID file)
'       RMID files are essentially MIDI files with a master chunk.
'       UseMicrosoft = True?
'       UseMasterChunk = True
'       --Master Chunk--
'           ID = "RIFF"
'           Type = "RMID"
'Note: An RMI file can also be disguised as a MID file.  The reverse is unlikely but entirely possible.
'      Detection of this should be easy since MIDI files start with MTHd and RMID files start with RIFF.
'
'   Chunk Names
'       "MTHd" - Header
'       "MTTk" - A track (may contain several)

'--- WAVE ---
'   Windows PCM
'       UseMicrosoft = True
'       UseMasterChunk = True
'       --Master Chunk--
'           ID = "RIFF"
'           Type = "WAVE"
'
'   Chunk Names
'       "fmt " - Format information
'       "data" - Data

Type MasterChunk
    GroupID As String * 4
    Size As Long
    Type As String * 4
End Type

Type Chunk
    ID As String * 4
    Offset As Long 'Offset to data
    Size As Long
    DataAcquired As Boolean
    Data As String
End Type

Type IFF
    FileName As String
    UseMicrosoft As Boolean
    UseMasterChunk As Boolean
    MasterChunk As MasterChunk
    Chunk() As Chunk
End Type

Function GetChunkIndex(IFF As IFF, ID As String) As Long
    Dim ChunkIndex As Long
    
    ID = FixedLengthString(4, ID)
    
    For ChunkIndex = LBound(IFF.Chunk) To UBound(IFF.Chunk)
        If IFF.Chunk(ChunkIndex).ID = ID Then
            GetChunkIndex = ChunkIndex
            Exit Function
        End If
    Next ChunkIndex
    
    GetChunkIndex = -1
End Function

Function GetChunk(IFF As IFF, ByVal ChunkIndex As Long) As Chunk
    On Error GoTo Escape
    
    With IFF.Chunk(ChunkIndex)
        If Not .DataAcquired Then
            Dim FileNum As Integer
            
            FileNum = FreeFile
            Open IFF.FileName For Binary Access Read As #FileNum
                Seek #FileNum, .Offset
                .Data = Input$(.Size, FileNum)
                .DataAcquired = True
            Close #FileNum
        End If
    End With
    
    GetChunk = IFF.Chunk(ChunkIndex)
Escape:
End Function

Function SetChunk(IFF As IFF, Chunk As Chunk, Optional ChunkIndex As Long = -1)

End Function

Function GetIFF(ByVal FileName As String, Optional ByVal UseMicrosoft As Boolean = True, Optional ByVal UseMasterChunk As Boolean = True, Optional AcquireData As Boolean = True, Optional Callback As Callback) As IFF
    'On Error Resume Next
    
    Dim FileNum As Integer
    Dim IFF As IFF
    Dim Data As String
    Dim ChunkIndex As Long
    
    With IFF
        FileNum = FreeFile
        .FileName = FileName
        Open .FileName For Binary Access Read As #FileNum
            .UseMicrosoft = UseMicrosoft
            .UseMasterChunk = UseMasterChunk
            If .UseMasterChunk Then
                Get #FileNum, , .MasterChunk
                If Not .UseMicrosoft Then .MasterChunk.Size = SwapEndianDWord(.MasterChunk.Size)
            End If
            
            Do
                DoEvents
                
                ReDim Preserve .Chunk(0 To ChunkIndex)
                
                With .Chunk(ChunkIndex)
                    .ID = Input$(4, FileNum)
                    Get #FileNum, , .Size
                    If Not IFF.UseMicrosoft Then .Size = SwapEndianDWord(.Size)
                    
                    .Offset = Loc(FileNum)
                    If AcquireData Then
                        .Data = Input$(.Size, FileNum)
                        .DataAcquired = True
                    Else
                        Seek #FileNum, Loc(FileNum) + .Size 'Skip over data
                    End If
                    
                    'The non-Microsoft (Electronic Arts) version of the IFF format requires that data in chunks be even.
                    'However, the last chunk can be odd so odd sized files are possible.
                    'This next If...End If block just skips over the unused byte if there is one.
                    If Not IFF.UseMicrosoft And Not EOF(FileNum) And IsOdd(.Size) Then Seek #FileNum, Loc(FileNum) + 1
                End With
                
                ChunkIndex = ChunkIndex + 1
                
                If Not Callback Is Nothing Then Callback.Progress Loc(FileNum), LOF(FileNum)
            Loop Until EOF(FileNum)
        Close #FileNum
    End With
    
    GetIFF = IFF
End Function

Sub SetIFF(ByVal FileName As String, IFF As IFF, Optional Callback As Callback)
    Set Callback = New Callback
    ''On Error Goto Handle
    
    Dim FileNum As Integer
    Dim ChunkIndex As Long
    
    With IFF
        If FileExists(FileName) Then Kill FileName
        
        FileNum = FreeFile
        Open FileName For Binary Access Write As #FileNum
            For ChunkIndex = LBound(.Chunk) To UBound(.Chunk)
                With .Chunk(ChunkIndex)
                    .Size = Len(.Data)
                    If Not IFF.UseMicrosoft And IsOdd(.Size) Then .Data = .Data & Chr$(0)
                End With
                
                .MasterChunk.Size = .MasterChunk.Size + 8 + Len(.Chunk(ChunkIndex).Data)
            Next ChunkIndex
            .MasterChunk.Size = .MasterChunk.Size + 4
            
            If .UseMasterChunk Then
                With .MasterChunk
                    Put #FileNum, , .GroupID
                    If IFF.UseMicrosoft Then
                        Put #FileNum, , .Size
                    Else
                        Put #FileNum, , SwapEndianDWord(.Size)
                    End If
                    Put #FileNum, , .Type
                End With
            End If
            
            For ChunkIndex = LBound(.Chunk) To UBound(.Chunk)
                DoEvents
                
                With .Chunk(ChunkIndex)
                    Put #FileNum, , .ID
                    If IFF.UseMicrosoft Then
                        Put #FileNum, , .Size
                    Else
                        Put #FileNum, , SwapEndianDWord(.Size)
                    End If
                    .Offset = Loc(FileNum)
                    Put #FileNum, , .Data
                End With
                
                If Not Callback Is Nothing Then Callback.Progress ChunkIndex, UBound(.Chunk)
            Next ChunkIndex
        Close #FileNum
        
        .FileName = FileName
    End With
    
    Exit Sub
handle:
    Callback.Fail = True
End Sub
