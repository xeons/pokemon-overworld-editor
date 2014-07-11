VERSION 5.00
Begin VB.Form frmOverworldEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pokémon Overworld Sprite Editor"
   ClientHeight    =   6735
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOverworldEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame grpSpriteHeaderTwo 
      Caption         =   "Sprite Header #2 Info"
      Height          =   1215
      Left            =   3120
      TabIndex        =   40
      Top             =   2640
      Width           =   2535
      Begin VB.Label lblUnknownHdr2 
         Caption         =   "Unknown 1:"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblSpriteDataSizeHdr2 
         Caption         =   "Data Size:"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblSpritePointerHdr2 
         Caption         =   "Sprite Pointer:"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame grpNavigation 
      Caption         =   "Sprite Navigation"
      Height          =   1215
      Left            =   120
      TabIndex        =   30
      Top             =   2640
      Width           =   2895
      Begin VB.TextBox txtSpriteFrame 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   37
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdSpriteFrameBackwards 
         Height          =   255
         Left            =   840
         Picture         =   "frmOverworldEditor.frx":151A
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   735
         Width           =   255
      End
      Begin VB.CommandButton cmdSpriteFrameForward 
         Height          =   255
         Left            =   1800
         Picture         =   "frmOverworldEditor.frx":1585
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   735
         Width           =   255
      End
      Begin VB.CommandButton cmdIndexForward 
         Height          =   255
         Left            =   1800
         Picture         =   "frmOverworldEditor.frx":15F1
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   375
         Width           =   255
      End
      Begin VB.CommandButton cmdIndexBack 
         Height          =   255
         Left            =   840
         Picture         =   "frmOverworldEditor.frx":165D
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   375
         Width           =   255
      End
      Begin VB.TextBox txtSpriteIndex 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Height          =   615
         Left            =   2160
         TabIndex        =   31
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblFrame 
         Caption         =   "Frame "
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   735
         Width           =   615
      End
      Begin VB.Label lblSpriteIndex 
         Caption         =   "Index "
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   375
         Width           =   615
      End
   End
   Begin VB.Frame grpSpriteHeader1 
      Caption         =   "Sprite Header #1 Info"
      Height          =   2655
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   5535
      Begin VB.Label lblUnknownPointer4 
         Caption         =   "Unknown Pointer 4:"
         Height          =   255
         Left            =   2880
         TabIndex        =   29
         Top             =   1380
         Width           =   2415
      End
      Begin VB.Label lblSpritePointer 
         Caption         =   "Sprite Pointer:"
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   1125
         Width           =   2415
      End
      Begin VB.Label lblUnknownPointer3 
         Caption         =   "Unknown Pointer 3:"
         Height          =   255
         Left            =   2880
         TabIndex        =   27
         Top             =   870
         Width           =   2415
      End
      Begin VB.Label lblUnknownPointer2 
         Caption         =   "Unknown Pointer 2:"
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   615
         Width           =   2415
      End
      Begin VB.Label lblUnknownPointer1 
         Caption         =   "Unknown Pointer 1:"
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblPalleteNumber 
         Caption         =   "Pallete #:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   870
         Width           =   2415
      End
      Begin VB.Label lblStarterBytes 
         Caption         =   "Starter Bytes:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   615
         Width           =   2415
      End
      Begin VB.Label lblUnknownData2 
         Caption         =   "Unknown Data 2:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2145
         Width           =   2415
      End
      Begin VB.Label lblSpriteNumber 
         Caption         =   "Sprite #:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblSpriteDataSize 
         Caption         =   "Sprite Data Size: "
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1380
         Width           =   2415
      End
      Begin VB.Label lblUnknownData 
         Caption         =   "Unknown Data:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1125
         Width           =   2415
      End
      Begin VB.Label lblSpriteWidth 
         Caption         =   "Width:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1635
         Width           =   2415
      End
      Begin VB.Label lblSpriteHeight 
         Caption         =   "Height:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1890
         Width           =   2415
      End
   End
   Begin VB.PictureBox picMouseOverColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7200
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   12
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdFlipVertical 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      Picture         =   "frmOverworldEditor.frx":16C8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton cmdFlipHorizontal 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      Picture         =   "frmOverworldEditor.frx":1A6E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton cmdRotateRight 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      Picture         =   "frmOverworldEditor.frx":1E0E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton cmdRotateLeft 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      Picture         =   "frmOverworldEditor.frx":21C6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   495
   End
   Begin VB.PictureBox picSelectPalette 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   7920
      MouseIcon       =   "frmOverworldEditor.frx":257F
      MousePointer    =   99  'Custom
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   2880
      Width           =   480
      Begin VB.Line lnePalleteGrid 
         Index           =   7
         X1              =   16
         X2              =   16
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line lnePalleteGrid 
         Index           =   6
         X1              =   0
         X2              =   32
         Y1              =   112
         Y2              =   112
      End
      Begin VB.Line lnePalleteGrid 
         Index           =   5
         X1              =   0
         X2              =   32
         Y1              =   96
         Y2              =   96
      End
      Begin VB.Line lnePalleteGrid 
         Index           =   4
         X1              =   0
         X2              =   32
         Y1              =   80
         Y2              =   80
      End
      Begin VB.Line lnePalleteGrid 
         Index           =   3
         X1              =   0
         X2              =   32
         Y1              =   64
         Y2              =   64
      End
      Begin VB.Line lnePalleteGrid 
         Index           =   2
         X1              =   0
         X2              =   32
         Y1              =   48
         Y2              =   48
      End
      Begin VB.Line lnePalleteGrid 
         Index           =   1
         X1              =   0
         X2              =   32
         Y1              =   32
         Y2              =   32
      End
      Begin VB.Line lnePalleteGrid 
         Index           =   0
         X1              =   0
         X2              =   32
         Y1              =   16
         Y2              =   16
      End
      Begin VB.Shape shpPalleteBorder 
         Height          =   1920
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox picEditTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   5880
      MouseIcon       =   "frmOverworldEditor.frx":26D1
      MousePointer    =   99  'Custom
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   2
      Top             =   2880
      Width           =   1920
      Begin VB.Line lneTileEditGrid 
         Index           =   13
         X1              =   128
         X2              =   0
         Y1              =   112
         Y2              =   112
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   12
         X1              =   128
         X2              =   0
         Y1              =   96
         Y2              =   96
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   11
         X1              =   128
         X2              =   0
         Y1              =   80
         Y2              =   80
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   10
         X1              =   128
         X2              =   0
         Y1              =   64
         Y2              =   64
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   9
         X1              =   128
         X2              =   0
         Y1              =   48
         Y2              =   48
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   8
         X1              =   128
         X2              =   0
         Y1              =   32
         Y2              =   32
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   7
         X1              =   128
         X2              =   0
         Y1              =   16
         Y2              =   16
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   6
         X1              =   112
         X2              =   112
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   5
         X1              =   96
         X2              =   96
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   4
         X1              =   80
         X2              =   80
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   3
         X1              =   64
         X2              =   64
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   2
         X1              =   48
         X2              =   48
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   1
         X1              =   32
         X2              =   32
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line lneTileEditGrid 
         Index           =   0
         X1              =   16
         X2              =   16
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Shape shpCanvasBorder 
         Height          =   1920
         Left            =   0
         Top             =   0
         Width           =   1920
      End
   End
   Begin VB.PictureBox picViewSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   120
      MouseIcon       =   "frmOverworldEditor.frx":2823
      MousePointer    =   99  'Custom
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   72
      TabIndex        =   0
      Top             =   120
      Width           =   1080
      Begin VB.Shape shpSelectedTile 
         BorderColor     =   &H000000FF&
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox picSelectedColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
      Begin VB.Label lblPaletteIndex 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Index: 00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   15
         Top             =   255
         Width           =   1215
      End
   End
   Begin VB.Label lblCurrentGame 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   5535
   End
   Begin VB.Label lblMouseOverColor 
      Caption         =   "Mouse-over"
      Height          =   255
      Left            =   7200
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lblTransformationOptions 
      Caption         =   "Transformation Options"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label lblSelectedColor 
      Caption         =   "Selected Color"
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lblDrawingCanvas 
      Caption         =   "Drawing Canvas"
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblPallete 
      Caption         =   "Pallete"
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportBitmap 
         Caption         =   "Import Bitmap..."
      End
      Begin VB.Menu mnuExportBitmap 
         Caption         =   "Export Bitmap..."
      End
      Begin VB.Menu mnuSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuShowGridlines 
         Caption         =   "Show Gridlines"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmOverworldEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lngSpritePalleteHeaders As Long
Private m_lngSpriteBank As Long
Private m_lngSpriteMax As Long
Private m_lngGraphicWidthBlocks As Long
Private m_lngGraphicHeightBlocks As Long
Private m_lngGraphicStartOffset As Long
Private m_strCurrentROM As String
Private m_lngCurrentMouseX As Long
Private m_lngCurrentMouseY As Long
Private m_lngPaletteCurrentMouseX As Long
Private m_lngPaletteCurrentMouseY As Long
Private m_lngEditCurrentMouseX As Long
Private m_lngEditCurrentMouseY As Long
Private m_lngCurrentTile As Long
Private m_lngSelectedPaletteEntry As Byte
Private m_blnStartColorDrag As Boolean
Private m_blnROMOpened As Boolean

Dim Buffer() As Long
Dim EditBuffer() As Byte
Dim Data1 As Byte, Data2 As Byte
Dim Data3 As Byte, Data4 As Byte
Dim iFreeFile As Integer
Dim i As Integer
Dim X As Integer, Y As Integer
Dim PaletteData(0 To 15) As Integer
Dim PaletteData2(0 To 15) As Long
Dim XLine As Long, YLine As Long

Private Sub cmdGo_Click()
    If Val(txtSpriteIndex.Text) >= 0 And Val(txtSpriteIndex.Text) <= m_lngSpriteMax Then
        LoadSpriteStructure Val(txtSpriteIndex.Text), Val(txtSpriteFrame.Text)
    End If
End Sub

Private Sub cmdIndexBack_Click()
    If Val(txtSpriteIndex.Text) > 0 Then txtSpriteIndex.Text = txtSpriteIndex.Text - 1
    txtSpriteFrame.Text = 0
    LoadSpriteStructure Val(txtSpriteIndex.Text), Val(txtSpriteFrame.Text)
End Sub

Private Sub cmdIndexForward_Click()
    If Val(txtSpriteIndex.Text) < m_lngSpriteMax Then txtSpriteIndex.Text = txtSpriteIndex.Text + 1
    txtSpriteFrame.Text = 0
    LoadSpriteStructure Val(txtSpriteIndex.Text), Val(txtSpriteFrame.Text)
End Sub

Private Sub cmdSpriteFrameBackwards_Click()
    If Val(txtSpriteFrame.Text) > 0 Then txtSpriteFrame.Text = txtSpriteFrame.Text - 1
    LoadSpriteStructure Val(txtSpriteIndex.Text), Val(txtSpriteFrame.Text)
End Sub

Private Sub cmdSpriteFrameForward_Click()
    If Val(txtSpriteFrame.Text) < 255 Then txtSpriteFrame.Text = txtSpriteFrame.Text + 1
    LoadSpriteStructure Val(txtSpriteIndex.Text), Val(txtSpriteFrame.Text)
End Sub

Private Sub Form_Load()
    ToggleEditing False
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExportBitmap_Click()
    'Dim bmpHeader As BITMAPINFOHEADER
    Dim bmpFileHeader As BITMAPFILEHEADER
    Dim bmpInfoHeader As BITMAPINFOHEADER
    Dim bytCurrentRow() As Byte
    Dim bytTempBuffer() As Byte
    Dim bytTempBuffer2() As Byte
    Dim lngCurrentPosition As Long
    Dim lWidth As Long, lHeight As Long
    lWidth = (m_lngGraphicWidthBlocks * 8)
    lHeight = (m_lngGraphicHeightBlocks * 8)

    ReDim bytTempBuffer(((m_lngGraphicWidthBlocks * 8) * (m_lngGraphicHeightBlocks * 8)) - 1)
    ReDim bytTempBuffer2((((m_lngGraphicWidthBlocks * 8) * (m_lngGraphicHeightBlocks * 8)) / 2) - 1)
    
    bmpFileHeader.bfType = 19778 'BM
    bmpFileHeader.bfSize = 118 + (((m_lngGraphicWidthBlocks * 8) * (m_lngGraphicHeightBlocks * 8)) \ 2)
    bmpFileHeader.bfOffBits = 118
    
    bmpInfoHeader.biSize = 40
    bmpInfoHeader.biWidth = m_lngGraphicWidthBlocks * 8
    bmpInfoHeader.biHeight = m_lngGraphicHeightBlocks * 8
    bmpInfoHeader.biPlanes = 1
    bmpInfoHeader.biCompression = 0
    bmpInfoHeader.biSizeImage = ((m_lngGraphicWidthBlocks * 8) * (m_lngGraphicHeightBlocks * 8)) \ 2
    bmpInfoHeader.biXPelsPerMeter = 0
    bmpInfoHeader.biYPelsPerMeter = 0
    bmpInfoHeader.biBitCount = 4
    bmpInfoHeader.biClrUsed = 0
    bmpInfoHeader.biClrImportant = 0
    
    For YLine = 0 To m_lngGraphicHeightBlocks - 1
        For XLine = 0 To m_lngGraphicWidthBlocks - 1
            For Y = 0 To 7
                For X = 0 To 7
                    bytTempBuffer((((XLine * 8) + X) + ((YLine * 8) + Y) * (m_lngGraphicWidthBlocks * 8))) = EditBuffer(lngCurrentPosition + X)
                Next X
                lngCurrentPosition = lngCurrentPosition + 8
            Next Y
        Next XLine
    Next YLine
    
    lngCurrentPosition = 0
    For i = 0 To UBound(bytTempBuffer) Step 2
        bytTempBuffer2(lngCurrentPosition) = (bytTempBuffer(i + 1) * 16) Or bytTempBuffer(i)
    Next i
    
    
    Open App.Path & "\Dump.bmp" For Binary As #1
        Put #1, , bmpFileHeader
        Put #1, , bmpInfoHeader
        Put #1, , PaletteData2
        Put #1, , bytTempBuffer2
    Close #1
End Sub

Private Sub mnuOpen_Click()
    Dim oOpenDialog As New clsCommonDialog
    Dim sResult As String
    Dim sGameCode As String * 4
    sResult = oOpenDialog.ShowOpen(Me.hWnd, "Open ROM Image...", , "GameBoy Advance ROM's (*.gba)|*.gba|", FILEMUSTEXIST Or PATHMUSTEXIST)
    If Len(sResult) > 0 Then
        m_strCurrentROM = sResult
        If FileExists(m_strCurrentROM) Then
            m_blnROMOpened = True
            iFreeFile = FreeFile
            Open m_strCurrentROM For Binary As #iFreeFile
                Seek #iFreeFile, &HAD&
                Get #iFreeFile, , sGameCode
                If Len(ReadINI(sGameCode, "Name", App.Path & "\Sprites.ini")) > 0 Then
                    lblCurrentGame.Caption = sGameCode & " - " & ReadINI(sGameCode, "Name", App.Path & "\Sprites.ini")
                    m_lngSpriteBank = Val(ReadINI(sGameCode, "SpriteBank", App.Path & "\Sprites.ini")) + 1
                    m_lngSpritePalleteHeaders = Val(ReadINI(sGameCode, "SpritePalleteHeaders", App.Path & "\Sprites.ini")) + 1
                    m_lngSpriteMax = Val(ReadINI(sGameCode, "SpriteCount", App.Path & "\Sprites.ini"))
                    If m_lngSpriteBank = 1 Or m_lngSpritePalleteHeaders = 1 Then
                        MsgBox "Error Loading INI Settings for this game...", vbExclamation, "Error"
                        Exit Sub
                    End If
                Else
                    MsgBox "Error Loading INI Settings for this game...", vbExclamation, "Error"
                    Exit Sub
                End If
            Close #iFreeFile
            
            txtSpriteFrame.Text = 0
            txtSpriteIndex.Text = 0
            
            LoadSpriteStructure Val(txtSpriteIndex.Text), Val(txtSpriteFrame.Text)
            ToggleEditing True
        End If
    End If
End Sub

Private Sub mnuSave_Click()

    SaveSprite m_strCurrentROM, m_lngGraphicStartOffset, m_lngGraphicHeightBlocks, m_lngGraphicWidthBlocks

End Sub

Private Sub mnuShowGridlines_Click()
    mnuShowGridlines.Checked = Not mnuShowGridlines.Checked
    For i = 0 To lnePalleteGrid.UBound
        lnePalleteGrid(i).Visible = mnuShowGridlines.Checked
    Next i
    For i = 0 To lneTileEditGrid.UBound
        lneTileEditGrid(i).Visible = mnuShowGridlines.Checked
    Next i
    shpCanvasBorder.Visible = mnuShowGridlines.Checked
    shpPalleteBorder.Visible = mnuShowGridlines.Checked
End Sub

Private Sub picEditTile_Click()
    EditBuffer((m_lngCurrentTile * 64) + (m_lngEditCurrentMouseX + m_lngEditCurrentMouseY * 8)) = m_lngSelectedPaletteEntry
    Call DrawTileEdit
End Sub

Private Sub picEditTile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X + 1 > picEditTile.Width Or Y + 1 > picEditTile.Height Or X < 0 Or Y < 0 Then Exit Sub
    'Make sure its left click
    If Button = 1 Then
        EditBuffer((m_lngCurrentTile * 64) + (m_lngEditCurrentMouseX + m_lngEditCurrentMouseY * 8)) = m_lngSelectedPaletteEntry
        Call DrawTileEdit
        m_blnStartColorDrag = True
    End If
End Sub

Private Sub picEditTile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picMouseOverColor.BackColor = Colour15To24RGB(PaletteData(EditBuffer((m_lngCurrentTile * 64) + (m_lngEditCurrentMouseX + m_lngEditCurrentMouseY * 8))))
    If Fix(X \ 16) = m_lngEditCurrentMouseX And m_lngEditCurrentMouseY = Fix(Y \ 16) Then Exit Sub
    m_lngEditCurrentMouseX = Fix(X \ 16)
    m_lngEditCurrentMouseY = Fix(Y \ 16)
    If m_blnStartColorDrag = True Then
        If X + 1 > picEditTile.Width Or Y + 1 > picEditTile.Height Or X < 0 Or Y < 0 Then Exit Sub
        EditBuffer((m_lngCurrentTile * 64) + (m_lngEditCurrentMouseX + m_lngEditCurrentMouseY * 8)) = m_lngSelectedPaletteEntry
        Call DrawTileEdit
        Call DrawSpriteView
    End If
End Sub

Private Sub picEditTile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(m_strCurrentROM) <= 0 Or m_blnROMOpened = False Then Exit Sub
    If Button = 1 Then
        m_blnStartColorDrag = False
        If X + 1 > picEditTile.Width Or Y + 1 > picEditTile.Height Or X < 0 Or Y < 0 Then Exit Sub
        EditBuffer((m_lngCurrentTile * 64) + (m_lngEditCurrentMouseX + m_lngEditCurrentMouseY * 8)) = m_lngSelectedPaletteEntry
        Call DrawTileEdit
        Call DrawSpriteView
    End If
    If Button = 2 Then
        If X + 1 > picEditTile.Width Or Y + 1 > picEditTile.Height Or X < 0 Or Y < 0 Then Exit Sub
        m_lngSelectedPaletteEntry = EditBuffer((m_lngCurrentTile * 64) + (m_lngEditCurrentMouseX + m_lngEditCurrentMouseY * 8))
        picSelectedColor.BackColor = Colour15To24RGB(PaletteData(m_lngSelectedPaletteEntry))
        lblPaletteIndex.Caption = "Index: " & Right("00" & Hex(m_lngSelectedPaletteEntry), 2)
        Exit Sub
    End If
End Sub

Private Sub picSelectPalette_Click()
    m_lngSelectedPaletteEntry = m_lngPaletteCurrentMouseX + (m_lngPaletteCurrentMouseY * 2)
    picSelectedColor.BackColor = Colour15To24RGB(PaletteData(m_lngSelectedPaletteEntry))
    lblPaletteIndex.Caption = "Index: " & Right("00" & Hex(m_lngSelectedPaletteEntry), 2)
End Sub

Private Sub picSelectPalette_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_lngPaletteCurrentMouseX = Fix(X \ 16)
    m_lngPaletteCurrentMouseY = Fix(Y \ 16)
    picMouseOverColor.BackColor = Colour15To24RGB(PaletteData(m_lngPaletteCurrentMouseX + (m_lngPaletteCurrentMouseY * 2)))
End Sub

Private Sub picViewSprite_Click()
    m_lngCurrentTile = m_lngCurrentMouseX + (m_lngCurrentMouseY * m_lngGraphicWidthBlocks)
    shpSelectedTile.Top = m_lngCurrentMouseY * 16
    shpSelectedTile.Left = m_lngCurrentMouseX * 16
    Call DrawTileEdit
End Sub

Private Sub picViewSprite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_lngCurrentMouseX = Fix(X \ 16)
    m_lngCurrentMouseY = Fix(Y \ 16)
End Sub

Private Sub LoadSprite(FilePath As String, SpriteOffset As Long, PaletteOffset As Long, Height As Long, Width As Long)
    'Set the global image properties
    m_lngGraphicHeightBlocks = Height
    m_lngGraphicWidthBlocks = Width
    'Gets us a new file handle.
    iFreeFile = FreeFile
    'Current position in decoding
    Dim lngCurrentPosition As Long
    'Resizes the buffers for the required space.
    ReDim Buffer(0 To ((m_lngGraphicHeightBlocks * 8) * (m_lngGraphicWidthBlocks * 8)))
    ReDim EditBuffer(0 To ((m_lngGraphicHeightBlocks * 8) * (m_lngGraphicWidthBlocks * 8)))
    'Resize the Sprite View (Double-Size mode!)
    picViewSprite.Width = (m_lngGraphicWidthBlocks * 8) * 2
    picViewSprite.Height = (m_lngGraphicHeightBlocks * 8) * 2
    'Reset current tile
    m_lngCurrentTile = 0
    shpSelectedTile.Top = 0
    shpSelectedTile.Left = 0
    'Set offsets
    m_lngGraphicStartOffset = SpriteOffset
    'Opens the ROM
    Open FilePath For Binary As #iFreeFile
        'Goto the offset of the Palette
        Seek #iFreeFile, PaletteOffset
        'Load the Palette up
        Get #iFreeFile, , PaletteData
        'Goto offset of sprite
        Seek #iFreeFile, SpriteOffset
        'The Y-Line is the row of blocks we are reading
        For YLine = 0 To m_lngGraphicHeightBlocks - 1
            'The X-Line is the column of blocks we are reading
            For XLine = 0 To m_lngGraphicWidthBlocks - 1
                For Y = 0 To 7
                    'We are reading the data in chunks of DWORD's! HOW EFFECIENT OF ME!
                    'I MUST BE GAWD!
                    Get #iFreeFile, , Data1
                    Get #iFreeFile, , Data2
                    Get #iFreeFile, , Data3
                    Get #iFreeFile, , Data4
                    '4BPP graphics are 4-bits per pixel, so 4 bytes is 8 pixels or 1 row.
                    'they are in reverse order. XXXXRRRR
                    'RRRR = First Pixel
                    'XXXX = Second Pixel
                    'Below ...that could probely be shortend to a loop.
                    EditBuffer(lngCurrentPosition + 0) = (Data1 And &HF)  'High Nibble
                    EditBuffer(lngCurrentPosition + 1) = (Data1 \ 16)     'Low Nibble
                    EditBuffer(lngCurrentPosition + 2) = (Data2 And &HF)  'High Nibble
                    EditBuffer(lngCurrentPosition + 3) = (Data2 \ 16)     'Low Nibble
                    EditBuffer(lngCurrentPosition + 4) = (Data3 And &HF)  'High Nibble
                    EditBuffer(lngCurrentPosition + 5) = (Data3 \ 16)     'Low Nibble
                    EditBuffer(lngCurrentPosition + 6) = (Data4 And &HF)  'High Nibble
                    EditBuffer(lngCurrentPosition + 7) = (Data4 \ 16)     'Low Nibble
                    'Increase current position by 8
                    lngCurrentPosition = lngCurrentPosition + 8
                Next Y
            Next XLine
        Next YLine
    Close #iFreeFile
End Sub

Private Sub SaveSprite(FilePath As String, SpriteOffset As Long, Height As Long, Width As Long)
    'Current position in decoding
    Dim lngCurrentPosition As Long
    Dim bytTempBuffer() As Byte
    On Error GoTo ErrSaveSprite
    ReDim bytTempBuffer((UBound(EditBuffer) \ 2) - 1)
    'Make sure file is open
    If Len(m_strCurrentROM) <= 0 Or m_blnROMOpened = False Then Exit Sub
    'Gets us a new file handle.
    iFreeFile = FreeFile
    'Opens the ROM
    Open m_strCurrentROM For Binary As #iFreeFile
        'Goto offset of sprite
        Seek #iFreeFile, SpriteOffset
        For i = 0 To UBound(EditBuffer) - 1 Step 2
            bytTempBuffer(lngCurrentPosition) = EditBuffer(i) Or (EditBuffer(i + 1) * 16)
            lngCurrentPosition = lngCurrentPosition + 1
        Next i
        Put #iFreeFile, , bytTempBuffer
    Close #iFreeFile

    On Error GoTo 0
    Exit Sub

ErrSaveSprite:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveSprite of frmOverworldEditor", vbCritical

    
End Sub

Private Sub DrawPalette()
    'Display Palette
    For i = 0 To 15
        PaletteData2(i) = Colour15To24(PaletteData(i))
    Next i
    Blit32 PaletteData2, picSelectPalette, 2, 8
    picSelectPalette.Refresh
End Sub

Private Sub DrawTileEdit()
    Dim lngTempBuffer(0 To 63) As Long
    Dim i As Integer
    For Y = 0 To 7
        For X = 0 To 7
            lngTempBuffer(X + Y * 8) = Colour15To24(PaletteData(EditBuffer((m_lngCurrentTile * 64) + (X + Y * 8))))
        Next X
    Next Y
    Blit32 lngTempBuffer, picEditTile, 8, 8
    picEditTile.Refresh
End Sub

Private Sub DrawSpriteView()
    Dim lngCurrentPosition As Long
    For YLine = 0 To m_lngGraphicHeightBlocks - 1
        For XLine = 0 To m_lngGraphicWidthBlocks - 1
            For Y = 0 To 7
                For X = 0 To 7
                    Buffer((((XLine * 8) + X) + ((YLine * 8) + Y) * (m_lngGraphicWidthBlocks * 8))) = Colour15To24(PaletteData(EditBuffer(lngCurrentPosition + X)))
                Next X
                lngCurrentPosition = lngCurrentPosition + 8
            Next Y
        Next XLine
    Next YLine
    Blit32 Buffer, picViewSprite, m_lngGraphicWidthBlocks * 8, m_lngGraphicHeightBlocks * 8
    picViewSprite.Refresh
End Sub

Private Sub LoadSpriteStructure(Index As Integer, Frame As Integer)

    Dim PrimarySpriteHeader As SpriteHeader
    Dim SecondarySpriteHeader As SpriteHeader2
    Dim PalleteHeaders(0 To 26) As PalleteHeader
    Dim j As Integer
    On Error GoTo ErrLoadSpriteStructure

    iFreeFile = FreeFile
    
    'Make sure the ROM is Open
    If Len(m_strCurrentROM) <= 0 Or m_blnROMOpened = False Then Exit Sub
    
    Open m_strCurrentROM For Binary As #iFreeFile
        '217 sprites
        'Seek #iFreeFile, &H3718D5 + (36 * Index)
        Seek #iFreeFile, m_lngSpriteBank + (36 * Index)
        Get #iFreeFile, , PrimarySpriteHeader
        
        lblSpriteNumber.Caption = "Sprite #: " & Index
        lblStarterBytes.Caption = "Starter Bytes: " & Right("0000" & Hex(PrimarySpriteHeader.StarterBytes), 4)
        lblPalleteNumber.Caption = "Pallete #: " & Right("00" & Hex(PrimarySpriteHeader.PalleteModifier), 2)
        lblUnknownData.Caption = "Unknown Data: " & Right("00" & Hex(PrimarySpriteHeader.Unknown1(0)), 2) & " " & _
                                                    Right("00" & Hex(PrimarySpriteHeader.Unknown1(1)), 2) & " " & _
                                                    Right("00" & Hex(PrimarySpriteHeader.Unknown1(2)), 2)
        lblSpriteDataSize.Caption = "Sprite Data Size: " & Right("0000" & Hex(PrimarySpriteHeader.SpriteDataSize), 4)
        lblSpriteWidth.Caption = "Width: " & PrimarySpriteHeader.Width
        lblSpriteHeight.Caption = "Height: " & PrimarySpriteHeader.Height
        lblUnknownData2.Caption = "Unknown Data 2: " & Right("00" & Hex(PrimarySpriteHeader.Unknown2), 2) & " " & Right("00" & Hex(PrimarySpriteHeader.Unknown3), 2) & " " & Right("0000" & Hex(PrimarySpriteHeader.Unknown4), 4)
        lblSpritePointer.Caption = "Sprite Pointer: " & Right("00000000" & Hex(PrimarySpriteHeader.SpriteHeader2Pointer), 8)
        lblUnknownPointer1.Caption = "Unknown Pointer 1: " & Right("00000000" & Hex(PrimarySpriteHeader.Pointer1), 8)
        lblUnknownPointer2.Caption = "Unknown Pointer 2: " & Right("00000000" & Hex(PrimarySpriteHeader.Pointer2), 8)
        lblUnknownPointer3.Caption = "Unknown Pointer 3: " & Right("00000000" & Hex(PrimarySpriteHeader.Pointer3), 8)
        lblUnknownPointer4.Caption = "Unknown Pointer 4: " & Right("00000000" & Hex(PrimarySpriteHeader.Pointer5), 8)
        
        
        Seek #iFreeFile, (PrimarySpriteHeader.SpriteHeader2Pointer - &H8000000) + 1 + (8 * Frame)
        Get #iFreeFile, , SecondarySpriteHeader
        
        lblSpritePointerHdr2.Caption = "Sprite Pointer: " & Right("00000000" & Hex(SecondarySpriteHeader.SpritePointer), 8)
        lblSpriteDataSizeHdr2.Caption = "Data Size: " & Right("0000" & Hex(SecondarySpriteHeader.SpriteDataSize), 4)
        lblUnknownHdr2.Caption = "Unknown 1: " & Right("0000" & Hex(SecondarySpriteHeader.Unknown), 4)
        
        Seek #iFreeFile, m_lngSpritePalleteHeaders
        Get #iFreeFile, , PalleteHeaders
        
    Close #iFreeFile
    
    For i = 0 To UBound(PalleteHeaders)
        If PalleteHeaders(i).Index = PrimarySpriteHeader.PalleteModifier Then
            j = i
            Exit For
        End If
    Next i
    
    Call LoadSprite(m_strCurrentROM, (SecondarySpriteHeader.SpritePointer - &H8000000) + 1, (PalleteHeaders(j).DataPointer - &H8000000) + 1, PrimarySpriteHeader.Height / 8, PrimarySpriteHeader.Width / 8)
    Call DrawSpriteView
    Call DrawTileEdit
    Call DrawPalette

    On Error GoTo 0
    Exit Sub

ErrLoadSpriteStructure:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadSpriteStructure of frmOverworldEditor", vbCritical, "Error"

End Sub


Private Sub ToggleEditing(blnEnable As Boolean)
    picEditTile.Enabled = blnEnable
    picSelectPalette.Enabled = blnEnable
    picViewSprite.Enabled = blnEnable
    cmdGo.Enabled = blnEnable
    cmdIndexBack.Enabled = blnEnable
    cmdIndexForward.Enabled = blnEnable
    cmdSpriteFrameBackwards.Enabled = blnEnable
    cmdSpriteFrameForward.Enabled = blnEnable
    txtSpriteFrame.Enabled = blnEnable
    txtSpriteIndex.Enabled = blnEnable
End Sub
