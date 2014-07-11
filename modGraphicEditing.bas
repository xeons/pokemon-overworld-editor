Attribute VB_Name = "modGraphicEditing"
Option Explicit

Public Type SpriteHeader
    StarterBytes As Integer
    PalleteModifier As Byte
    Unknown1(2) As Byte
    SpriteDataSize As Integer
    Width As Integer
    Height As Integer
    Unknown2 As Byte
    Unknown3 As Byte
    Unknown4 As Integer
    Pointer1 As Long
    Pointer2 As Long
    Pointer3 As Long
    SpriteHeader2Pointer As Long
    Pointer5 As Long
End Type

Public Type SpriteHeader2
    SpritePointer As Long
    SpriteDataSize As Integer
    Unknown As Integer
End Type

Public Type PalleteHeader
    DataPointer As Long
    Index As Byte
    UnknownData(2) As Byte
End Type


Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColor0 As Long
    bmiColor1 As Long
    bmiColor2 As Long
End Type

Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Public Type BITMAPV4HEADER
    bV4Size As Long
    bV4Width As Long
    bV4Height As Long
    bV4Planes As Integer
    bV4BitCount As Integer
    bV4V4Compression As Long
    bV4SizeImage As Long
    bV4XPelsPerMeter As Long
    bV4YPelsPerMeter As Long
    bV4ClrUsed As Long
    bV4ClrImportant As Long
    bV4RedMask As Long
    bV4GreenMask As Long
    bV4BlueMask As Long
    bV4AlphaMask As Long
    bV4CSType As Long
    bV4Endpoints As Long
    bV4GammaRed As Long
    bV4GammaGreen As Long
    bV4GammaBlue As Long
End Type


Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Public Function ReadINI(ByVal Section As String, ByVal KeyName As String, ByVal FileName As String) As String
    Dim sRet As String
    sRet = String(256, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function

Public Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Call WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function


Public Sub Blit32(Buffer() As Long, ByRef pic As PictureBox, imagewidth As Long, imageheight As Long)
    If imagewidth = 0 Or imageheight = 0 Then Exit Sub
    Dim bi As BITMAPINFO
    With bi.bmiHeader
        .biWidth = imagewidth
        .biHeight = -imageheight
        .biSize = 40
        .biBitCount = 32
        .biPlanes = 1
    End With
    If pic.ScaleMode <> 3 Then pic.ScaleMode = 3
    StretchDIBits pic.hdc, 0, 0, pic.ScaleWidth, pic.ScaleHeight, 0, 0, imagewidth, imageheight, Buffer(0), bi, 0, vbSrcCopy
End Sub

Public Sub Blit15(Buffer() As Integer, ByRef pic As PictureBox, imagewidth As Long, imageheight As Long)
    If imagewidth = 0 Or imageheight = 0 Then Exit Sub
    Dim bi As BITMAPINFO
    With bi.bmiHeader
        .biWidth = imagewidth
        .biHeight = -imageheight
        .biSize = 40
        .biBitCount = 16
        .biPlanes = 1
    End With
    If pic.ScaleMode <> 3 Then pic.ScaleMode = 3
    StretchDIBits pic.hdc, 0, 0, pic.ScaleWidth, pic.ScaleHeight, 0, 0, imagewidth, imageheight, Buffer(0), bi, 0, vbSrcCopy
End Sub

Public Function Colour15To24(ByVal ColourData As Integer) As Long
    Dim R As Byte, G As Byte, B As Byte
    R = ((ColourData And 31) / 31) * &HFF
    G = (((ColourData \ 32) And 31) / 31) * &HFF
    B = (((ColourData \ 1024) And 31) / 31) * &HFF
    Colour15To24 = CLng(B) + (256 * CLng(G)) + (65536 * CLng(R))
End Function

Public Function Colour15To24RGB(ByVal ColourData As Integer) As Long
    Dim R As Byte, G As Byte, B As Byte
    R = ((ColourData And 31) / 31) * &HFF
    G = (((ColourData \ 32) And 31) / 31) * &HFF
    B = (((ColourData \ 1024) And 31) / 31) * &HFF
    Colour15To24RGB = CLng(R) + (256 * CLng(G)) + (65536 * CLng(B))
End Function

Public Function FileExists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
