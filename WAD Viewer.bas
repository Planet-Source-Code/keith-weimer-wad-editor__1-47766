Attribute VB_Name = "modWADViewer"
Option Explicit

'WAD Viewer

'Written by Keith R. Weimer
'Way Too Happy Software

'Include: TreeView Control (Microsoft Common Controls 6.0)
'Include: WAD.bas
'Include: Callback.cls

'Description: Generates a categorized tree in a TreeView control of lumps in a WAD file.

'NOTE: This sorting is tailored mostly to Doom, Doom 2, and Final Doom since I don't have
'      the Heretic and Hexen WADs to look at for reference.

Private Enum BlockType
    BT_Single
    BT_Map
    BT_StartEnd
End Enum

Function IsChildNode(Parent As Node, Child As Node) As Boolean
    Dim Node As Node
    
    Set Node = Child.Parent
    Do Until Node Is Nothing
        If Node Is Parent Then
            IsChildNode = True
            Exit Do
        End If
        
        Set Node = Node.Parent
    Loop
End Function

Function GetLumpIndexFromNode(Node As Node) As Long
    If Not Node Is Nothing Then
        If IsNumeric(Node.Tag) Then GetLumpIndexFromNode = CLng(Node.Tag)
    End If
End Function

Sub CreateLumpTree(WAD As WAD, TreeView As TreeView, Optional Callback As Callback)
    If Callback Is Nothing Then Set Callback = New Callback 'Create dummy callback
    On Error GoTo handle
    
    Dim Index As Long
    
    Dim Sound As Node
        Dim SoundPC As Node
        Dim SoundWAV As Node
    Dim Music As Node
    Dim Graphics As Node
        Dim Sprites As Node
        Dim Patches As Node
        Dim Flats As Node
        Dim Menu As Node
        Dim Status As Node
        Dim LevelStatus As Node
        Dim Border As Node
        Dim FullScreen As Node
    Dim Maps As Node
    Dim Demos As Node
    Dim Other As Node
    
    Dim BlockType As BlockType
    Dim Parent As Node
    Dim Temp As Node
    Dim Name As String
    Dim Skip As Boolean
    
    TreeView.Nodes.Clear
    
    'Create categories
    With TreeView.Nodes
        Set Sound = .Add(, , "Sound", "Sound")
            Set SoundPC = .Add(Sound, tvwChild, "PC Speaker", "PC Speaker")
            Set SoundWAV = .Add(Sound, tvwChild, "Wave", "Wave")
        Set Music = .Add(, , "Music", "Music")
        Set Graphics = .Add(, , "Graphics", "Graphics")
            Set Sprites = .Add(Graphics, tvwChild, "Sprites", "Sprites")
            Set Patches = .Add(Graphics, tvwChild, "Patches", "Patches")
            Set Flats = .Add(Graphics, tvwChild, "Flats", "Flats")
            Set Menu = .Add(Graphics, tvwChild, "Menu", "Menu")
            Set Status = .Add(Graphics, tvwChild, "Status", "Status")
            Set LevelStatus = .Add(Graphics, tvwChild, "Level Status", "Level Status")
            Set Border = .Add(Graphics, tvwChild, "Border", "Border")
            Set FullScreen = .Add(Graphics, tvwChild, "Full Screen", "Full Screen")
        Set Maps = .Add(, , "Maps", "Maps")
        Set Demos = .Add(, , "Demos", "Demos")
        Set Other = .Add(, , "Other", "Other")
    End With
    
    For Index = 1 To TreeView.Nodes.Count
        With TreeView.Nodes(Index)
            .Image = 1
            .ExpandedImage = 2
        End With
    Next Index
    
    BlockType = BT_Single
    Set Parent = Other
    
    For Index = 1 To WAD.LumpCount
        With WAD.Lump(Index)
            'Lump tree doesn't take that long to make and this just slows it down:
            'If DoEvents = 0 Then Exit Sub 'Stop loading if user is closing
            
            'Name can be changed to a fancy lump name.  WAD.Lump(LumpIndex).Name should not be changed.
            Name = .Name
            
            If BlockType = BT_Map Then
                Select Case .Name
                    'NOTICE: Keep these around even if you don't want fancy map data lump names.
                    '        They are used to identify if a lump is part of a map block.
                    'NOTICE: Theoretically lumps could become nonmap data lumps and then switch back to
                    '        map data lumps, causing them to become independant of the parent map, but
                    '        this is unlikely.  If this is a problem (you have a messed up WAD!) then
                    '        a LastMap node would likely need to be implemented.
                    
                    Case "THINGS" ': Name = "Things"
                    Case "LINEDEFS" ': Name = "Line Definitions"
                    Case "SIDEDEFS" ': Name = "Side Definitions"
                    Case "VERTEXES" ': Name = "Vertexes"
                    Case "SEGS" ': Name = "Segments"
                    Case "SSECTORS" ': Name = "Sub-sectors"
                    Case "NODES" ': Name = "Nodes"
                    Case "SECTORS" ': Name = "Sectors"
                    Case "REJECT" ': Name = "Reject"
                    Case "BLOCKMAP" ': Name = "BlockMap"
                    Case Else
                        BlockType = BT_Single
                        Set Parent = Other
                End Select
            End If
            
            Select Case True
                'NOTICE: This assumes for every *_START there is a *_END even if this is not true
                Case .Name Like "*#_START"
                    BlockType = BT_StartEnd
                    
                    Select Case .Name
                        Case "P1_START", "F1_START": Name = "Unregistered"
                        Case "P2_START", "F2_START": Name = "Registered"
                        Case "P3_START", "F3_START" 'What are these?
                    End Select
                Case .Name Like "*_START"
                    BlockType = BT_StartEnd
                    
                    Skip = True
                    Select Case .Name
                        Case "S_START": Set Parent = Sprites
                        Case "P_START": Set Parent = Patches
                        Case "F_START": Set Parent = Flats
                        Case Else: Skip = False
                    End Select
                    
                    If Skip Then Parent.Tag = Index
                    
                    'NOTICE: Blocks can theoretically nest indefinately.
                    'NOTICE: Skip is used in this manner to allow substandard *_START-*_ENDs to work.
                    '        A substandard block is placed in the Other category unless it is a child
                    '        of a otherwise catergorized block.
                'Case .Name Like "*_END": Skip = True
                Case IsMapName(.Name)
                    BlockType = BT_Map
                    Set Parent = Maps
                Case Else
                    If BlockType = BT_Single Then
                        Select Case True
                            Case .Name Like "DEMO#": Set Parent = Demos
                            Case .Name Like "DP*": Set Parent = SoundPC
                            Case .Name Like "DS*": Set Parent = SoundWAV
                            Case .Name Like "D_*", .Name = "GENMIDI", .Name Like "DMXGUS*": Set Parent = Music
                            Case .Name Like "TEXTURE#", .Name = "PNAMES", .Name = "PLAYPAL", .Name = "COLORMAP", .Name Like "AMMNUM#", .Name Like "END#": Set Parent = Graphics
                            Case .Name Like "WI*", .Name Like "CWI*": Set Parent = LevelStatus
                            Case .Name Like "ST*": Set Parent = Status
                            Case .Name Like "M_*": Set Parent = Menu
                            Case .Name Like "BRDR_*": Set Parent = Border
                            Case .Name Like "HELP*", .Name = "TITLEPIC", .Name = "CREDIT", .Name = "VICTORY2", .Name = "PFUB1", .Name = "PFUB2", .Name = "ENDPIC", .Name = "INTERPIC", .Name = "BOSSBACK": Set Parent = FullScreen
                            Case Else: Set Parent = Other
                        End Select
                    End If
            End Select
            
            If Skip Then
                Skip = False
            Else
                Set Temp = TreeView.Nodes.Add(Parent, tvwChild, , Name)
                
                'NOTICE: This allows us to get the LumpIndex from any node.
                '        Of course, if the Tag property is empty then the node is just
                '        there for show and doesn't actually refer to a lump.
                Temp.Tag = Index
                
                Select Case True
                    Case IsMapName(.Name), .Name Like "*_START"
                        Temp.Image = 1
                        Temp.ExpandedImage = 2
                End Select
                
                Select Case BlockType
                    Case BT_Map: If IsMapName(.Name) Then Set Parent = Temp
                    Case BT_StartEnd: If .Name Like "*_START" Then Set Parent = Temp
                End Select
            End If
            
            If BlockType = BT_StartEnd Then
                If .Name Like "*_END" Then
                    Set Parent = Parent.Parent
                    If Parent Is Nothing Then
                        BlockType = BT_Single
                        Set Parent = Other
                    End If
                End If
            End If
            
            Callback.Progress Index, WAD.LumpCount
        End With
    Next Index
    
    Exit Sub
handle:
    Callback.Fail = True
End Sub
