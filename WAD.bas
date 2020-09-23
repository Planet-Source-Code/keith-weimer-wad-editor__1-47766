Attribute VB_Name = "modWAD"
Option Explicit
Option Compare Text

'Handling for WADs

'Written by Keith R. Weimer
'Way Too Happy Software

'Include: Memory.bas
'Include: Callback.cls

'*** Converts ***
'Include: Bitmap.cls
'Include: Wave.cls

Enum WADType
    WT_Internal
    WT_Patch
End Enum

Enum PictureType
    PT_Default
    PT_Flat
End Enum

Type Lump
    Name As String
    Offset As Long
    Size As Long
    Data As String
    
    DataAcquired As Boolean
    'Optimized As Boolean
End Type

Type WAD
    FileName As String
    Type As WADType
    LumpCount As Long
    Lump() As Lump
End Type

Type PictureInfo
    Width As Integer
    Height As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

Public Const WADFilter As String = "WAD Files (*.wad)|*.wad|All Files (*.*)|*.*"

'This will either gather the lump index from a lump name or pass the lump index if provided.
Function GetLumpIndex(WAD As WAD, ByVal Identifier As Variant, Optional ByVal Start As Long = 1) As Long
    If VarType(Identifier) = vbString Then
        Dim Index As Long
        
        For Index = Start To WAD.LumpCount
            If WAD.Lump(Index).Name = Identifier Then
                GetLumpIndex = Index
                Exit Function
            End If
        Next Index
        
        GetLumpIndex = 0
    Else
        GetLumpIndex = Identifier
    End If
End Function

'GetLump should always be used to collect data from a lump.
'Check the DataAquired element of the lump to check for failure.
Function GetLump(WAD As WAD, ByVal LumpIndex As Variant) As Lump
    On Error GoTo Escape
    
    LumpIndex = GetLumpIndex(WAD, LumpIndex)
    
    If LumpIndex >= 1 Or LumpIndex <= WAD.LumpCount Then
        If Not WAD.Lump(LumpIndex).DataAcquired Then
            Dim FileNum As Integer
            
            FileNum = FreeFile
            Open WAD.FileName For Binary Access Read As #FileNum
                With WAD.Lump(LumpIndex)
                    If .Size > 0 Then
                        Seek #FileNum, .Offset
                        .Data = Input$(.Size, FileNum)
                    End If
                    .DataAcquired = True
                End With
            Close #FileNum
        End If
        
        GetLump = WAD.Lump(LumpIndex)
    End If
Escape:
End Function

Sub GetAllLumps(WAD As WAD, Optional Callback As Callback)
    If Callback Is Nothing Then Set Callback = New Callback 'Create dummy callback
    
    Dim FileNum As Integer
    Dim LumpIndex As Long
    
    Callback.StatusText "Reading WAD (Lump Data)..."
    
    FileNum = FreeFile
    Open WAD.FileName For Binary Access Read As #FileNum
        For LumpIndex = 1 To WAD.LumpCount
            DoEvents
            
            With WAD.Lump(LumpIndex)
                If .Size > 0 Then
                    Seek #FileNum, .Offset
                    .Data = Input$(.Size, FileNum)
                End If
                .DataAcquired = True
            End With
            
            Callback.Progress LumpIndex, WAD.LumpCount
        Next LumpIndex
    Close #FileNum
End Sub

Function SetLump(WAD As WAD, ByVal LumpIndex As Variant, Lump As Lump)
    LumpIndex = GetLumpIndex(WAD, LumpIndex)
    
    If LumpIndex > 0 Then
        If Lump.Size > Len(Lump.Data) Or Lump.Size = 0 Then Lump.Size = Len(Lump.Data)
        Lump.DataAcquired = True
        
        WAD.Lump(LumpIndex) = Lump
    End If
End Function

Function GetWAD(ByVal FileName As String, Optional ByVal AcquireData As Boolean, Optional Callback As Callback) As WAD
    If Callback Is Nothing Then Set Callback = New Callback 'Create dummy callback
    On Error GoTo handle
    
    Dim FileNum As Integer
    Dim Header As String * 4
    
    Callback.StatusText "Reading WAD..."
    
    FileNum = FreeFile
    
    GetWAD.FileName = FileName
    Open GetWAD.FileName For Binary Access Read As #FileNum
        Get #FileNum, , Header
        
        If Header = "IWAD" Or Header = "PWAD" Then
            Dim DirectoryOffset As Long
            Dim LumpIndex As Long
            
            Select Case Header
                Case "IWAD": GetWAD.Type = WT_Internal
                Case "PWAD": GetWAD.Type = WT_Patch
            End Select
            
            Get #FileNum, , GetWAD.LumpCount
                ReDim GetWAD.Lump(1 To GetWAD.LumpCount)
            Get #FileNum, , DirectoryOffset
            
            Callback.StatusText "Reading WAD (Lump Directory)..."
            Seek #FileNum, DirectoryOffset + 1
            For LumpIndex = 1 To GetWAD.LumpCount
                If DoEvents = 0 Then Exit For
                
                With GetWAD.Lump(LumpIndex)
                    Get #FileNum, , .Offset
                    .Offset = .Offset + 1 'WADs use 0 based offsets, VB uses 1
                    Get #FileNum, , .Size
                    
                    .Name = String$(8, 0)
                    Get #FileNum, , .Name
                    .Name = RTrimNull(.Name)
                End With
                
                Callback.Progress LumpIndex, GetWAD.LumpCount
            Next LumpIndex
            
            Callback.StatusText "Reading WAD (Lump Data)..."
            For LumpIndex = 1 To GetWAD.LumpCount
                DoEvents
                
                With GetWAD.Lump(LumpIndex)
                    If .Size = 0 Then
                        .DataAcquired = True
                    ElseIf AcquireData Then
                        Seek #FileNum, .Offset
                        .Data = Input$(.Size, FileNum)
                        .DataAcquired = True
                    End If
                End With
                
                Callback.Progress LumpIndex, GetWAD.LumpCount
            Next LumpIndex
            
            Callback.StatusText "Reading WAD...Complete."
        Else
            Callback.Fail = True
        End If
    Close #1
    
    Exit Function
handle:
    Callback.Fail = True
End Function

Sub SetWAD(ByVal FileName As String, WAD As WAD, Optional Callback As Callback)
    If Callback Is Nothing Then Set Callback = New Callback 'Create dummy callback
    On Error GoTo handle
    
    Dim DirectoryOffset As Long
    Dim FileNum As Integer
    Dim Header As String * 4
    
    Dim LumpIndex As Long
    
    GetAllLumps WAD, Callback
    
    Callback.StatusText "Writing WAD..."
    
    If FileExists(FileName) Then Kill FileName
    FileNum = FreeFile
    Open FileName For Binary Access Write As #FileNum
        Select Case WAD.Type
            Case WT_Internal: Header = "IWAD"
            Case WT_Patch: Header = "PWAD"
        End Select
        
        'Header
        Put #FileNum, , Header
        Put #FileNum, , WAD.LumpCount
        Put #FileNum, , 0& 'DirectoryOffset -- We'll come back to this.
        
        'Lump Data
        Callback.StatusText "Writing WAD (Lump Data)..."
        
        For LumpIndex = 1 To WAD.LumpCount
            With WAD.Lump(LumpIndex)
                .Offset = Seek(FileNum)
                                
                Put #FileNum, , .Data
                'Lumps are aligned by 4 bytes (32-bits) and
                'padded with the first byte of the lump's data
                If .Size Mod 4 > 0 Then Put #FileNum, , String$(4 - .Size Mod 4, Left$(.Data, 1))
            End With
            
            Callback.Progress LumpIndex, WAD.LumpCount
        Next LumpIndex
        
        'Lump Directory
        Callback.StatusText "Writing WAD (Lump Directory)..."
        
        DirectoryOffset = Seek(FileNum) - 1
        For LumpIndex = 1 To WAD.LumpCount
            With WAD.Lump(LumpIndex)
                If .Size <> 0 Or IsMapName(.Name) Then
                    Put #FileNum, , .Offset - 1
                Else
                    Put #FileNum, , 0&
                End If
                
                Put #FileNum, , .Size
                Put #FileNum, , FixedLengthString(8, .Name)
            End With
            
            Callback.Progress LumpIndex, WAD.LumpCount
        Next LumpIndex
        
        Put #FileNum, 9, DirectoryOffset 'See?  I told you we'd come back.
    Close #FileNum
    
    Callback.StatusText "Writing WAD...Complete."
    
    Exit Sub
handle:
    Callback.Fail = True
End Sub

'WADs can be insanely redundant (especially IWADs).  The same data is repeated several time.
'For example, several levels have the same music but the data for the music is repeated within the WAD.
'If you were to make 15 maps of the same 1MB map then it would take 15MB in the WAD.
'The purpose of this sub is to point lumps of the same data at each other, thereby optimizing the WAD.
Sub OptimizeWAD(WAD As WAD, Callback As Callback)

End Sub

Function IsMapName(ByVal LumpName As String) As Boolean
    IsMapName = LumpName Like "E#M#" Or LumpName Like "MAP##"
End Function

'--------------------------------------------------------------------------------
'   Converters
'--------------------------------------------------------------------------------

'----------------------------------------
'Sound
'----------------------------------------
Function LumpToWave(Lump As Lump) As Wave
    Dim Size As Long
    Dim Data As String
    
    Size = DWord(StringToWord(Mid$(Lump.Data, 5, 2)))
    Data = Mid$(Lump.Data, 8, Size)
    
    Set LumpToWave = New Wave
    LumpToWave.Create 1, DWord(StringToWord(Mid$(Lump.Data, 3, 2))), 8, Data
End Function

Function WaveToLump(Wave As Wave) As Lump
    If Wave.Channels <> 1 Or Wave.BitsPerSample <> 8 Then
        'Wave.Resample 1, , 8
        MsgBox "8-bit mono wave only supported."
        Exit Function
    End If
    
    With WaveToLump
        .Data = WordToString(3) & WordToString(Wave.SampleRate) & WordToString(Len(Wave.Data)) & WordToString(0) & Wave.Data
        .Size = Len(.Data)
        .DataAcquired = True
    End With
End Function

'----------------------------------------
'Graphics
'----------------------------------------
Function GetPalette(WAD As WAD, ByVal PaletteIndex As Integer) As Long()
    On Error Resume Next
    
    Dim Lump As Lump
    
    Lump = GetLump(WAD, "PLAYPAL")
    If Lump.DataAcquired Then
        Dim PaletteCount As Integer
        
        PaletteCount = Len(Lump.Data) \ 768
        If PaletteIndex >= 0 Or PaletteIndex <= PaletteCount - 1 Then
            Dim Palette(0 To 255) As Long
            Dim Color As Byte
            Dim Red As Byte
            Dim Green As Byte
            Dim Blue As Byte
            
            For Color = 0 To 255
                Red = Asc(Mid$(Lump.Data, PaletteIndex * 768 + Color * 3 + 1, 1))
                Green = Asc(Mid$(Lump.Data, PaletteIndex * 768 + Color * 3 + 2, 1))
                Blue = Asc(Mid$(Lump.Data, PaletteIndex * 768 + Color * 3 + 3, 1))
                
                Palette(Color) = RGB(Red, Green, Blue)
            Next Color
            
            GetPalette = Palette
        End If
    Else
        'TODO: Make this an error.
        MsgBox "Palettes not found."
    End If
End Function

'This function does what its supposed to but could be way better.
'I wanted it to return an OLE picture but apparently I don't know enough about GDI.
'Why must it be so hard to create a bitmap, set its palette to Doom's palette,
'load the picture data into it, and convert it to an OLE picture?  Huh?!?
'Is that so much to ask?  Why must it be?
Function LumpToPicture(WAD As WAD, ByVal LumpIndex As Variant, Target As Object, Optional ByVal Mirror As Boolean, Optional ByVal PictureType As PictureType, Optional ByVal PaletteIndex As Integer) As PictureInfo
    Dim Lump As Lump
    Dim Palette() As Long
    Dim Color As Byte
    
    Dim X As Integer
    Dim Y As Integer
    
    Lump = GetLump(WAD, LumpIndex)
    Palette = GetPalette(WAD, PaletteIndex)
    
    With LumpToPicture
        Select Case PictureType
            Case PT_Default
                'I thought this format was totally stupid before I actually
                'understood what it did.  Yay!  Ignorance!
                'Using DMGRAPH I thought the cyan mask color was actually stored
                'into the WAD but this isn't so.
                'What's transparent isn't actually stored but skipped over.
                'This format saved Doom's graphics engine from having to do
                'stupid Windows crap like masking.
                
                Dim Offset As Long
                Dim Pixels As Byte
                Dim Pixel As Byte
                                
                .Width = StringToWord(Mid$(Lump.Data, 1, 2))
                .Height = StringToWord(Mid$(Lump.Data, 3, 2))
                .OffsetX = StringToWord(Mid$(Lump.Data, 5, 2))
                .OffsetY = StringToWord(Mid$(Lump.Data, 7, 2))
                
                For X = 0 To .Width - 1
                    If Mirror Then
                        Offset = StringToDWord(Mid$(Lump.Data, 9 + (.Width - 1 - X) * 4, 4))
                    Else
                        Offset = StringToDWord(Mid$(Lump.Data, 9 + X * 4, 4))
                    End If
                    
                    'Read each post
                    Do
                        Y = Asc(Mid$(Lump.Data, Offset + 1, 1))
                        If Y = 255 Then Exit Do
                        
                        Pixels = Asc(Mid$(Lump.Data, Offset + 2, 1))
                        
                        For Pixel = 0 To Pixels - 1
                            Color = Asc(Mid$(Lump.Data, Offset + Pixel + 4))
                            
                            Target.PSet (X, Y + Pixel), Palette(Color)
                        Next Pixel
                        
                        Offset = Offset + Pixels + 4
                    Loop Until Offset > Len(Lump.Data)
                Next X
            Case PT_Flat
                'I liked this format!  Was easy to write code for.
                
                .Width = 64
                .Height = 64
                
                For X = 0 To .Width - 1
                    For Y = 0 To .Height - 1
                        If Mirror Then
                            Color = Asc(Mid$(Lump.Data, (.Width - 1 - X) * .Width + Y + 1, 1))
                        Else
                            Color = Asc(Mid$(Lump.Data, X * .Width + Y + 1, 1))
                        End If
                        
                        Target.PSet (X, Y), Palette(Color)
                    Next Y
                Next X
        End Select
    End With
End Function
