VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "WAD Editor"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   4680
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList 
      Left            =   480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   2910
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2655
      Visible         =   0   'False
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.TreeView treLump 
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4048
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   1
      ImageList       =   "ImageList"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save..."
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuPasteData 
         Caption         =   "Paste Data"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Preview..."
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import..."
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim WAD As WAD
Dim WithEvents Callback As Callback
Attribute Callback.VB_VarHelpID = -1

Private Sub Callback_Progress(ByVal Value As Long, ByVal Total As Long)
    ProgressBar.Value = (Value / Total) * 100
End Sub

Private Sub Callback_StatusText(ByVal Text As String)
    ProgressBar.Value = 0
    StatusBar.Panels(1).Text = Text
End Sub

Private Sub Form_Paint()
    On Error Resume Next
    
    treLump.Width = ScaleWidth
    treLump.Height = ScaleHeight - IIf(ProgressBar.Visible, ProgressBar.Height, 0) - StatusBar.Height
End Sub

Private Sub Form_Resize()
    Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Unload frmDisplay
    Unload frmPlayer
    If FileExists(AddSlash(App.Path) & "Temp.wav") Then Kill AddSlash(App.Path) & "Temp.wav"
End Sub

Private Sub mnuExport_Click()
    On Error Resume Next
    
    Dim LumpIndex As Long
    
    LumpIndex = GetLumpIndexFromNode(treLump.SelectedItem)
    If LumpIndex Then
        Dim Lump As Lump
        
        Lump = GetLump(WAD, LumpIndex)
        If Lump.DataAcquired Then
            With Lump
                Select Case True
                    Case .Name Like "DS*": cdlMain.Filter = "Wave Files (*.wav)|*.wav|"
                    Case .Name Like "D_*": cdlMain.Filter = "Music Files (*.mus)|*.mus|"
                    Case Else: cdlMain.Filter = Empty
                End Select
                
                cdlMain.Filter = cdlMain.Filter & "Lump Files (*.lmp)|*.lmp|All Files (*.*)|*.*"
            End With
            
            With cdlMain
                .FileName = Empty
                .ShowSave
                
                If Err.Number <> cdlCancel Then
                    Select Case GetExtensionName(.FileName)
                        Case "wav"
                            Dim Wave As New Wave
                            
                            Set Wave = LumpToWave(Lump)
                            Wave.SaveFile .FileName
                        Case Else: WriteFile .FileName, Lump.Data
                    End Select
                End If
            End With
        End If
    End If
End Sub

Private Sub mnuImport_Click()
    On Error Resume Next
    
    Dim LumpIndex As Long
    
    LumpIndex = GetLumpIndexFromNode(treLump.SelectedItem)
    If LumpIndex Then
        With WAD.Lump(LumpIndex)
            Select Case True
                Case .Name Like "DS*": cdlMain.Filter = "Wave Files (*.wav)|*.wav|"
                Case .Name Like "D_*": cdlMain.Filter = "Music Files (*.mus)|*.mus|"
                Case Else: cdlMain.Filter = Empty
            End Select
            
            cdlMain.Filter = cdlMain.Filter & "Lump Files (*.lmp)|*.lmp|All Files (*.*)|*.*"
        End With
        
        With cdlMain
            .FileName = Empty
            
            .ShowOpen
            
            If Err.Number <> cdlCancel Then
                Dim Lump As Lump
                
                Lump.Name = WAD.Lump(LumpIndex).Name
                If GetExtensionName(.FileName) = "wav" And WAD.Lump(LumpIndex).Name Like "DS*" Then
                    Dim Wave As New Wave
                    
                    Wave.OpenFile .FileName
                    Lump = WaveToLump(Wave)
                Else
                    Lump.Data = ReadFile(.FileName)
                End If
                
                SetLump WAD, LumpIndex, Lump
            End If
        End With
    End If
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next
    
    With cdlMain
        .FileName = Empty
        .Filter = WADFilter
        .ShowOpen
        
        If Err.Number <> cdlCancel Then
            ProgressBar.Visible = True
            Refresh
            
            Set Callback = New Callback
            WAD = GetWAD(.FileName, , Callback)
            
            If Callback.Fail Then
                StatusBar.Panels(1).Text = "Reading WAD...Failed."
            Else
                StatusBar.Panels(1).Text = "Reading WAD (Generating Lump Tree)..."
                
                CreateLumpTree WAD, treLump, Callback
                
                ProgressBar.Visible = False
                ProgressBar.Value = 0
                Refresh
                
                StatusBar.Panels(1).Text = "Reading WAD...Complete."
            End If
        End If
    End With
End Sub

Private Sub mnuPreview_Click()
    On Error Resume Next
    
    Dim LumpIndex As Long
    
    LumpIndex = GetLumpIndexFromNode(treLump.SelectedItem)
    If LumpIndex Then
        If WAD.Lump(LumpIndex).Name Like "DS*" Then
            Dim Wave As Wave
            
            Set Wave = LumpToWave(GetLump(WAD, LumpIndex))
            
            Load frmPlayer
            frmPlayer.CloseDevice
            Wave.SaveFile AddSlash(App.Path) & "Temp.wav"
            frmPlayer.Show , Me
            frmPlayer.OpenFile AddSlash(App.Path) & "Temp.wav"
            frmPlayer.mciPlayer.Command = "Play"
        ElseIf IsChildNode(treLump.Nodes("Flats"), treLump.SelectedItem) Then
            frmDisplay.Visible = True
            frmDisplay.Cls
            LumpToPicture WAD, LumpIndex, frmDisplay, , PT_Flat
        ElseIf IsChildNode(treLump.Nodes("Graphics"), treLump.SelectedItem) Then
            frmDisplay.Visible = True
            frmDisplay.Cls
            LumpToPicture WAD, LumpIndex, frmDisplay
        End If
    End If
End Sub

Private Sub mnuSave_Click()
    On Error Resume Next
    
    With cdlMain
        .FileName = Empty
        .Filter = WADFilter
        .ShowSave
        
        If Err.Number <> cdlCancel Then
            ProgressBar.Visible = True
            Refresh
            
            Set Callback = New Callback
            SetWAD .FileName, WAD, Callback
            
            ProgressBar.Visible = False
            ProgressBar.Value = 0
            Refresh
        End If
    End With
End Sub

Private Sub treLump_DblClick()
    mnuPreview_Click
End Sub

Private Sub treLump_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Dim Node As Node
        Dim LumpIndex As Long
        
        Set Node = treLump.HitTest(X, Y)
        
        LumpIndex = GetLumpIndexFromNode(Node)
        If LumpIndex Then
            Set treLump.SelectedItem = Node
            
            mnuPreview.Enabled = WAD.Lump(LumpIndex).Name Like "DS*" Or IsChildNode(treLump.Nodes("Graphics"), Node)
            
            mnuImport.Enabled = True
            mnuExport.Enabled = True
        Else
            mnuPreview.Enabled = False
            
            mnuImport.Enabled = False
            mnuExport.Enabled = False
        End If
        
        PopupMenu mnuEdit
    End If
End Sub

