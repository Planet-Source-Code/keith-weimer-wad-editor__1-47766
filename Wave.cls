VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Handling for Windows PCM waves

'Written by Keith R. Weimer
'Way Too Happy Software

'Include: Memory.bas
'Include: IFF.bas

Dim intChannels As Integer
Dim lngSampleRate As Long
Dim lngAvgBytesPerSec As Long
Dim intBlockAlign As Integer
Dim intBitsPerSample As Integer
Public Data As String

Property Get Channels() As Integer
    Channels = intChannels
End Property

Property Get SampleRate() As Long
    SampleRate = lngSampleRate
End Property

Property Get AvgBytesPerSec() As Long
    AvgBytesPerSec = lngAvgBytesPerSec
End Property

Property Get BlockAlign() As Integer
    BlockAlign = intBlockAlign
End Property

Property Get BitsPerSample() As Integer
    BitsPerSample = intBitsPerSample
End Property

Private Sub CalculateWave()
    intBlockAlign = Ceiling(intChannels * (intBitsPerSample / 8))
    lngAvgBytesPerSec = lngSampleRate * intBlockAlign
End Sub

Sub Create(ByVal Channels As Integer, ByVal SampleRate As Long, ByVal BitsPerSample As Integer, Optional Data As String)
    If Channels <> 0 And SampleRate <> 0 And BitsPerSample <> 0 Then
        intChannels = Channels
        lngSampleRate = SampleRate
        intBitsPerSample = BitsPerSample
        Me.Data = Data
        
        CalculateWave
    End If
End Sub

Sub Resample(Optional ByVal Channels As Integer = -1, Optional ByVal SampleRate As Long = -1, Optional ByVal BitsPerSample As Integer = -1, Optional Callback As Callback)

End Sub

Friend Function FromIFF(IFF As IFF)
    With IFF
        If .UseMicrosoft And .MasterChunk.GroupID = "RIFF" And .MasterChunk.Type = "WAVE" Then
            Dim FormatChunkIndex As Long
            Dim DataChunkIndex As Long
            
            FormatChunkIndex = GetChunkIndex(IFF, "fmt ")
            DataChunkIndex = GetChunkIndex(IFF, "data")
            
            If FormatChunkIndex = -1 Then
                MsgBox "Format information missing from wave.", vbExclamation
            Else
                Dim FormatChunk As Chunk
                Dim FormatTag As Integer
                
                FormatChunk = GetChunk(IFF, FormatChunkIndex)
                
                With FormatChunk
                    FormatTag = StringToWord(Mid$(.Data, 1, 2))
                    If FormatTag <> 1 Then
                        MsgBox "Compression is not supported."
                    Else
                        intChannels = StringToWord(Mid$(.Data, 3, 2))
                        lngSampleRate = StringToDWord(Mid$(.Data, 5, 4))
                        lngAvgBytesPerSec = StringToDWord(Mid$(.Data, 9, 4))
                        intBlockAlign = StringToWord(Mid$(.Data, 13, 2))
                        intBitsPerSample = StringToWord(Mid$(.Data, 15, 2))
                        
                        If DataChunkIndex <> -1 Then Data = GetChunk(IFF, DataChunkIndex).Data
                        
                        CalculateWave
                    End If
                End With
            End If
        Else
            MsgBox "Not a Windows PCM wave.", vbExclamation
        End If
    End With
End Function

Friend Function ToIFF() As IFF
    With ToIFF
        'Windows PCM
        .UseMicrosoft = True
        .UseMasterChunk = True
        .MasterChunk.GroupID = "RIFF"
        .MasterChunk.Type = "WAVE"
        
        ReDim .Chunk(0 To 1)
        .Chunk(0).ID = "fmt "
                        'FormatTag (Uncompressed only supported)
        .Chunk(0).Data = WordToString(1) & WordToString(intChannels) & DWordToString(lngSampleRate) & DWordToString(lngAvgBytesPerSec) & WordToString(intBlockAlign) & WordToString(intBitsPerSample)
        .Chunk(1).ID = "data"
        .Chunk(1).Data = Data
    End With
End Function

Sub OpenFile(ByVal FileName As String)
    Dim IFF As IFF
    
    IFF = GetIFF(FileName, True, True, True)
    FromIFF IFF
End Sub

Function SaveFile(ByVal FileName As String)
    Dim IFF As IFF
    
    IFF = ToIFF
    SetIFF FileName, IFF
End Function
