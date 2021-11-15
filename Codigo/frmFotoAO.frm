VERSION 5.00
Begin VB.Form frmFotoAO 
   Caption         =   "Form1"
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   2730
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFoto 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmFotoAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Private Const picDesc = " ImperiumAO        (c)2005"
Private FN As String
Private IndiceFoto As Integer

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type
Private Type PicBmp
    size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type



Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcsrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (picDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long



Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
Dim r As Long
Dim Pic As PicBmp
         ' IPicture requires a reference to "Standard OLE Types."
Dim IPic As IPicture
Dim IID_IDispatch As GUID

         ' Fill in with IDispatch Interface ID.
With IID_IDispatch
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
End With

         ' Fill Pic with necessary parts.
         With Pic
            .size = Len(Pic)          ' Length of structure.
            .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
            .hBmp = hBmp              ' Handle to bitmap.
            .hPal = hPal              ' Handle to palette (may be null).
         End With

         ' Create Picture object.
         r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

         ' Return the new Picture object.
         Set CreateBitmapPicture = IPic
      End Function

Private Function CaptureWindow(ByVal hWndSrc As Long, _
    ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) As Picture

Dim hDCMemory As Long
Dim hBmp As Long
Dim hBmpPrev As Long
Dim r As Long
Dim hdcsrc As Long
Dim hPal As Long
Dim hPalPrev As Long
Dim RasterCapsScrn As Long
Dim HasPaletteScrn As Long
Dim PaletteSizeScrn As Long
Dim LogPal As LOGPALETTE

' Depending on the value of Client get the proper device context.
hdcsrc = GetDC(hWndSrc) ' Get device context for client area.
' Create a memory device context for the copy process.
hDCMemory = CreateCompatibleDC(hdcsrc)
' Create a bitmap and place it in the memory DC.
hBmp = CreateCompatibleBitmap(hdcsrc, WidthSrc, HeightSrc)
hBmpPrev = SelectObject(hDCMemory, hBmp)

' Get screen properties.
RasterCapsScrn = GetDeviceCaps(hdcsrc, RASTERCAPS) ' Raster
' capabilities.
HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                              ' support.
PaletteSizeScrn = GetDeviceCaps(hdcsrc, SIZEPALETTE) ' Size of
                                                              ' palette.
' If the screen has a palette make a copy and realize it.
If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    ' Create a copy of the system palette.
    LogPal.palVersion = &H300
    LogPal.palNumEntries = 256
    r = GetSystemPaletteEntries(hdcsrc, 0, 256, _
        LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
    ' Select the new palette into the memory DC and realize it.
    hPalPrev = SelectPalette(hDCMemory, hPal, 0)
    r = RealizePalette(hDCMemory)
End If

' Copy the on-screen image into the memory DC.
r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hdcsrc, _
    LeftSrc, TopSrc, vbSrcCopy)
' Remove the new copy of the  on-screen image.
hBmp = SelectObject(hDCMemory, hBmpPrev)

' If the screen has a palette get back the palette that was
' selected in previously.
If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    hPal = SelectPalette(hDCMemory, hPalPrev, 0)
End If

' Release the device context resources back to the system.
r = DeleteDC(hDCMemory)
r = ReleaseDC(hWndSrc, hdcsrc)

' Call CreateBitmapPicture to create a picture object from the
' bitmap and palette handles. Then return the resulting picture
' object.
Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function


'devuelve -1 en error, mayor que cero (indice de foto) si esta bien
Public Function Foto() As Long
On Local Error GoTo FotoErrorHandler
Dim t As String

Set picFoto = CaptureWindow(frmMain.hwnd, 0, 0, (frmMain.width \ Screen.TwipsPerPixelX) - 4, (frmMain.height \ Screen.TwipsPerPixelY) - 4 - 20) '-4 por los bordes - 20 por la barra de titulo
picFoto.CurrentX = 10
picFoto.CurrentY = Screen.TwipsPerPixelY * 565
picFoto.Print picDesc
picFoto.CurrentY = Screen.TwipsPerPixelY * 565
t = Date & " - " & Time
picFoto.CurrentX = (Screen.TwipsPerPixelY * 750) - TextWidth(t)
picFoto.Print t
SavePicture picFoto.Image, FN & Trim(Str(IndiceFoto)) & ".bmp"
Foto = IndiceFoto
IndiceFoto = IndiceFoto + 1

Call AddtoRichTextBox(frmMain.RecTxt, "Screenshot grabada correctamente como " & FN & Trim(Str(IndiceFoto)) & ".bmp", 65, 190, 156, False, True, False)

Exit Function
FotoErrorHandler:
Foto = -1
End Function

Private Sub Form_Load()
Dim i As Integer
With picFoto
    .FontBold = True
    .FontSize = 12
    .FontTransparent = False
    .ForeColor = vbRed
    .BackColor = vbBlack
End With
FN = App.Path & "\Fotos\ImpAO_Foto"
i = 1
If Dir(App.Path & "\Fotos", vbDirectory) = "" Then
    MkDir (App.Path & "\Fotos")
End If
Do While Dir(FN & Trim(Str(i)) & ".bmp") <> ""
    i = i + 1
    DoEvents
Loop
IndiceFoto = i
End Sub
