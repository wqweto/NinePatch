VERSION 5.00
Begin VB.UserControl ctxNineButton 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4044
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   KeyPreview      =   -1  'True
   ScaleHeight     =   105
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   Windowless      =   -1  'True
End
Attribute VB_Name = "ctxNineButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' Nine Patch PNGs for VB6 (c) 2018-2022 by wqweto@gmail.com
'
' ctxNineButton.ctl -- windowless 9-patch button control w/ state animation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "ctxNineButton"

#Const ImplUseShared = NPPNG_USE_SHARED <> 0
#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)
#Const ImplSelfContained = True

'=========================================================================
' Public events
'=========================================================================

Event Click()
Attribute Click.VB_UserMemId = -600
Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Event ContextMenu()
Event OwnerDraw(ByVal hGraphics As Long, ByVal hFont As Long, ByVal ButtonState As UcsNineButtonStateEnum, ClientLeft As Long, ClientTop As Long, ClientWidth As Long, ClientHeight As Long, Caption As String, ByVal hPicture As Long)
Event RegisterCancelMode(oCtl As Object, Handled As Boolean)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -604
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event AccessKeyPress(KeyAscii As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As Object, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragOver(Data As Object, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As Object, DataFormat As Integer)
Event OLEStartDrag(Data As Object, AllowedEffects As Long)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsNineButtonStateEnum
    ucsBstNormal = 0
    ucsBstHover = 1
    ucsBstPressed = 2
    ucsBstHoverPressed = 3
    ucsBstDisabled = 4
    ucsBstFocused = 8
End Enum
Private Const ucsBstLast = ucsBstFocused

Public Enum UcsNineButtonTextFlagsEnum
    ucsBflHorLeft = 0
    ucsBflHorCenter = 1
    ucsBflHorRight = 2
    ucsBflVertTop = 0
    ucsBflVertCenter = 4
    ucsBflVertBottom = 8
    ucsBflCenter = ucsBflHorCenter Or ucsBflVertCenter
    ucsBflDirectionRightToLeft = &H1 * 16
    ucsBflDirectionVertical = &H2 * 16
    ucsBflNoFitBlackBox = &H4 * 16
    ucsBflDisplayFormatControl = &H20 * 16
    ucsBflNoFontFallback = &H400 * 16
    ucsBflMeasureTrailingSpaces = &H800& * 16
    ucsBflNoWrap = &H1000& * 16
    ucsBflLineLimit = &H2000& * 16
    ucsBflNoClip = &H4000& * 16
End Enum

Public Enum UcsNineButtonStyleEnum
    ucsBtyNone
    ucsBtyButtonDefault
    ucsBtyButtonGreen
    ucsBtyButtonRed
    ucsBtyButtonTurnGreen
    ucsBtyButtonTurnRed
    ucsBtyFlatPrimary
    ucsBtyFlatSecondary
    ucsBtyFlatSuccess
    ucsBtyFlatDanger
    ucsBtyFlatWarning
    ucsBtyFlatInfo
    ucsBtyFlatLight
    ucsBtyFlatDark
    ucsBtyOutlinePrimary
    ucsBtyOutlineSecondary
    ucsBtyOutlineSuccess
    ucsBtyOutlineDanger
    ucsBtyOutlineWarning
    ucsBtyOutlineInfo
    ucsBtyOutlineLight
    ucsBtyOutlineDark
    ucsBtyCardDefault
    ucsBtyCardPrimary
    ucsBtyCardSuccess
    ucsBtyCardOrange
    ucsBtyCardDanger
    ucsBtyCardWarning
    ucsBtyCardPurple
    ucsBtyCardFocus
End Enum

Public Enum UcsNineButtonOleDropMode
    ucsModNone
    ucsModManual
End Enum

'=========================================================================
' API
'=========================================================================

'--- for GdipCreateBitmapFromScan0
Private Const PixelFormat32bppARGB          As Long = &H26200A
Private Const PixelFormat32bppPARGB         As Long = &HE200B
'--- for GdipDrawImageXxx
Private Const UnitPixel                     As Long = 2
Private Const UnitPoint                     As Long = 3
'--- for GdipSetTextRenderingHint
Private Const TextRenderingHintAntiAlias    As Long = 4
Private Const TextRenderingHintClearTypeGridFit As Long = 5
'--- DIB Section constants
Private Const DIB_RGB_COLORS                As Long = 0 '  color table in RGBs
'--- for GdipBitmapLockBits
Private Const ImageLockModeRead             As Long = &H1
Private Const ImageLockModeWrite            As Long = &H2
'--- for matrix order
'Private Const MatrixOrderPrepend            As Long = 0
Private Const MatrixOrderAppend             As Long = 1
'--- for thunks
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, pIconInfo As ICONINFO) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function ApiUpdateWindow Lib "user32" Alias "UpdateWindow" (ByVal hWnd As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal lX As Long, ByVal lY As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'--- gdi+
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, pInputBuf As Any, Optional ByVal pOutputBuf As Long = 0) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal lWidth As Long, ByVal lHeight As Long, ByVal lStride As Long, ByVal lPixelFormat As Long, ByVal Scan0 As Long, hBitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal hImage As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal srcUnit As Long = UnitPixel, Optional ByVal hImageAttributes As Long, Optional ByVal pfnCallback As Long, Optional ByVal lCallbackData As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (hImgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hImgAttr As Long, ByVal lAdjustType As Long, ByVal fAdjustEnabled As Long, clrMatrix As Any, grayMatrix As Any, ByVal lFlags As Long) As Long
Private Declare Function GdipSetImageAttributesColorKeys Lib "gdiplus" (ByVal hImgAttr As Long, ByVal lAdjustType As Long, ByVal fAdjustEnabled As Long, ByVal clrLow As Long, ByVal clrHigh As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImgAttr As Long) As Long
Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal hBitmap As Long, ByVal lX As Long, ByVal lY As Long, clrCurrent As Any) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal lNamePtr As Long, ByVal hFontCollection As Long, hFontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "gdiplus" (hFontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal hFontFamily As Long) As Long
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal hFontFamily As Long, ByVal emSize As Single, ByVal lStyle As Long, ByVal lUnit As Long, hFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal hFont As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, hBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal hGraphics As Long, ByVal lStrPtr As Long, ByVal lLength As Long, ByVal hFont As Long, uRect As RECTF, ByVal hStringFormat As Long, ByVal hBrush As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal hFormatAttributes As Long, ByVal nLanguage As Integer, hStringFormat As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal hStringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal hStringFormat As Long, ByVal lFlags As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal hStringFormat As Long, ByVal eAlign As StringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal hStringFormat As Long, ByVal eAlign As StringAlignment) As Long
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal hGraphics As Long, ByVal lMode As Long) As Long
Private Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal lPixelFormat As Long, ByVal srcBitmap As Long, dstBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBmp As Long, ByVal hPal As Long, hBtmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal hIcon As Long, hBitmap As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef nWidth As Single, ByRef nHeight As Single) As Long '
Private Declare Function GdipCloneImage Lib "gdiplus" (ByVal hImage As Long, hCloneImage As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, lpRect As Any, ByVal lFlags As Long, ByVal lPixelFormat As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal hGraphics As Long, ByVal nDx As Single, ByVal nDy As Single, ByVal lOrder As Long) As Long
Private Declare Function GdipScaleWorldTransform Lib "gdiplus" (ByVal hGraphics As Long, ByVal nSx As Single, ByVal nSy As Single, ByVal lOrder As Long) As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
#End If
#If Not ImplUseShared Then
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    '--- for thunks
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
#End If

Private Type RECTF
   Left                 As Single
   Top                  As Single
   Right                As Single
   Bottom               As Single
End Type

Private Enum FontStyle
   FontStyleRegular = 0
   FontStyleBold = 1
   FontStyleItalic = 2
   FontStyleBoldItalic = 3
   FontStyleUnderline = 4
   FontStyleStrikeout = 8
End Enum

Private Enum StringAlignment
   StringAlignmentNear = 0
   StringAlignmentCenter = 1
   StringAlignmentFar = 2
End Enum

Private Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    B                   As Byte
    A                   As Byte
End Type

Private Type ICONINFO
    fIcon               As Long
    xHotspot            As Long
    yHotspot            As Long
    hbmMask             As Long
    hbmColor            As Long
End Type

Private Type BitmapData
    Width               As Long
    Height              As Long
    Stride              As Long
    PixelFormat         As Long
    Scan0               As Long
    Reserved            As Long
End Type

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As Long
    cElements           As Long
    lLbound             As Long
End Type

'=========================================================================
' Constants and variables
'=========================================================================

Private Const DBL_EPLISON           As Double = 0.000001
Private Const DEF_STYLE             As Long = ucsBtyNone
Private Const DEF_ENABLED           As Boolean = True
Private Const DEF_OPACITY           As Single = 1
Private Const DEF_ANIMATIONDURATION As Single = 0
Private Const DEF_FORECOLOR         As Long = vbButtonText
Private Const DEF_MANUALFOCUS       As Boolean = False
Private Const DEF_MASKCOLOR         As Long = vbMagenta
Private Const DEF_AUTOREDRAW        As Boolean = False
Private Const DEF_TEXTOPACITY       As Single = 1
Private Const DEF_TEXTCOLOR         As Long = -1  '--- none
Private Const DEF_TEXTFLAGS         As Long = ucsBflCenter
Private Const DEF_IMAGEOPACITY      As Single = 1
Private Const DEF_IMAGEZOOM         As Single = 1
Private Const DEF_SHADOWOPACITY     As Single = 0.5
Private Const DEF_SHADOWCOLOR       As Long = vbButtonShadow
Private Const STR_RES_PNG1          As String = "iVBORw0KGgoAAAANSUhEUgAAAOcAAACfCAYAAAAChc6MAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAA3S0lEQVR4Xu2dCXwUVbr2vYx8cxf93fEOKkoWliRAAggkgWyQgCDigsjijsvoeBUUt1mU63cNGFlUYARDQEEUgglLQHABdCSQIAgJRBwhCZpOoqDIruNEEubjfO9T3ZWcnH4rpyrpADFVv98/TZ33eZ9uu/N4qrtOVy4QQtiCNrrha43F9eRrjcX15GuN5Vx7en8EcFPvwEVc8OCCwrZ3pO0OHTtld/yY5wtSVDCOOnRcv0vLZMGCBW1ffnVh6Euz58dPm52RooJx1KHj+r0/Aripd9CaSRWizS1TihJvmbJr0pjJBak6oLstrSgBfZyfS4uhzYuzMxKnz5o3" & _
                                                "adrM9FQd0E1/JT0BfbKP90cAN9m8NTMudc9lo58vmMCFUAf6bkv7WzDn63J+MzM9PXj67HkTuBDqQB/6TS/vD2W7KXV35KgphemjJxcWjZ5SKPygcdRvTi3o4Wup3UxjldFTPovy9hTmj5pcIFTIK9/05Po5En+/KirpoXfSB45fWzTg4XeEStLD7xShnvDQatue/e/LjIy7f9m8+PvfLiIEQxHqMfcvadDzlud3DzHDdkva7tFj0z7rYHXYinHUxz5fMKa2Z3LhIE4LaPuXim8O3VHxzZGCygNHqioPHhV+0Djq5V8fuR16zkdl6svzI6fOfHXeCzPT8+hW+PHyq/l0O++Fl16x9Xzifr8oLrujpNSzs2S/p4oQDFWoFxeXOX6cU2elF02blS5UMO70cW7fvvPOHTsKCnbsKKzauXOXUME46tu27bijocf50qx5Q8ywvfiXjNGzZmV0sDpsxTjqM2anjzF70G/WvT+kzQjm5MLy0VMK" & _
                                                "athg1lJQM2pKgUcNqGksA8/RqTvzb55c8MioqQXthqd+2lYF46gjpHYCmjhhXdSA8Ws9Ax/5oCZ54oci5bGP/MA46gPGv+OxE9DY378dFfe7rPLEB1fWJD20xi/sRuBpHPW4+7M8Cfdns56pueLCMVMKnzWC9vyuGziNFdCjb+zzu/6H9tlfgvKvv3/4wKHjournanHmzBka8t8wjjp00NOQn48MfuFfeHnuVmLCn6dPb0e0ZWiH+tSX5+bb+cX/Yt+X478qqxTHjh0X//jHP0RVVZUfGEcdOug5HxkjmLPSyyl8p7lgmqBOeOw8zu3bCx/evesz4SnziIMHDohvDx70A+OoQwc950P8Cx2iPmsEc2aGo9cdevTNmJ1R+7obBXkbPWXX0tHP7zrDB1LB0O1a6ms1NvPOZDAjjkotmMCFUgU66DkfmQEPrc1MfmQ9G0oV6KDnfGTiHng7M/G/V7GhVIEOes7nzhl7gsb4ZsCxqTvbcxor" & _
                                                "oK/tffELtrf8m8N7TlXX0D/1G3TQ0z/9fGSMmcYbTC6U9TACSnrOR6a4pGzPiRMn2VCqQAc95yND95vJhdEK6DkfmU8/LdxTUV7OhlIFOug5nxfTF7c3Z8AX09Mdve7Q1/UuNnqNgrzRrFnMBtGKybv2+lqNTb5DE8yGFLzfqkG04LcjSc/5yCQ9vLYkeSIfRhXoSL+X85GhQ9YSOhRmw6gCXdzv3mY9b03bFTfGF7B7U3P/ldNYAb3ZCx9OU3Hg6EmrGVPdoIOe/unnI4NDVgqe1Yyp0u6Fma/mcT4yJfvLTlrNmCrQQc/5yFDYirkQWjF15jzt675zZ+FJqxlTBTroOZ8Zc+bFmQFLXbzY0esOvdkLH4wZBXljA6jB12ps8h2a4D0lE0JLoFc9VBAQLohWQK96qOA9pRxAHdCrHsA4ReILGFfXYfbCh6vjPaWTzaf385GhX3pBoeOCyAK96qGC95RcEK2AXvVQ4QKoQ/VQwXtKLohW" & _
                                                "QK96AOMUiS9gXF2H2Qsf7BuD8saFT4ev1djkOzNxw+kMs9cNpz9c+HSoHirnbTiRJ+Onb+PCp8PXam7YJ++6O21qODlPhIMLoRXQy/2cZ1PDaXoGMpzc42xqODnPpoaT82xqODlPLnw65H7Os6nhND0DGU7DE4PyxoVPh6/V2NQ7BO7M6Qyz1505/eHCp0P1UDlvZ07jh7Rx4dPhazU2+c5M3HA6w+x1w+kPFz4dqodKiwnnqMmFpVwArYDe12ps8p2ZOP20FnrORyZp/NpSR5/Wkp7zkUl4IKvUyae10HM+ZyOcTj6ttRVOh5/WQs/5yCBszj6ttRHOmemlXAAtIT3nI4OwOfu09tyFM5MLoRXQ+1qNTb4zk9Z2nnPstF1hY3wBa9J5TvLhNJ6vv//cyXlO6Omffj4yNBMG/DznvuKvPndynhN6zkeG7jfg5zm3by/43Ml5Tug5n1mzFoaZAWvKeU74YMwoyBtW51A4Krgg+jG5sBJ6" & _
                                                "X6uxyXdoAg1mQzsrhEZOLsyDnvORwYofmr0qBz7ygbESiA2ld4UQhWltpZ0VQljxE/9AVmXigyuNlUBcKH0rhAR0ViuExqZ/cdEYX8Aau0IIjJ217d84TfH+ykedrBCCnob8fGSwksZY+eMNaEBWCBXu/ttEJyuEoOd8ZIzHOSu9gguiP69W2nmcmzblT3SyQgh6zic9fcVFteFs/Aqh515asuQ/MGYU1A3hGD2lIIdmxXIulBReD+pqMLHJdygDLWZEhBTvKVWMNbfeuvbJNPEGdG3OgPHrPFyQMI66k7W1WDMbd39WTvz92R68p/Tjgewy1PVrawtHjjFD9nzBGDtra7EG1+xBP6cFtP266PP995ZVfr+robW1qH/+xf57oed8VIxffJoRfWtovetpJXxrbp2sWf319u277t2778tdxaVlP+OwVQXjqG/fWeTscc6am0OPxcOFcurMdA/qTh7nhx9+fO/27Tt3NbS2FvWPPvr4voYe" & _
                                                "54zZ80aaAcWaWTtra7EGt65nXu3r7v0RwM00bu0gcLc8v6vR30qxCrLL+Q0CN2PW/PFm2JyAb6XIQfb+COBmGrt4v885dsru5NpF8BrwfU58/xN9nJ9Li6HNi7PmJZuL4HXg+5z4/if6ZB/vjwBusrmLF8yC7pUQWh+YBZt8JQQ7eHPH1xqL68nXGovrydcay7n2ZAc5zvUDtYvrydcai+vJ1xqLE0920CWwPFgY3fapj7uHPr6+WzyRwhCPOnRcv0vzUxgd3faTTt1DN4d2i8/t2C1FBeOoQ8f1NwfsoEtgSBWpbR7/sGviExu6TXp8fUSqDuie3NgtAX2cn0szkJraJi+kayKFb1JuSESqDui2hHZLQB/rF0DYQZem8+TGHsGPb4iYwIVQC/Whn/N1CRxbuvQIzg2NmMCFUAv1oZ/zDRTs4KMfhEU+tj4i/bGN4UWPbYgQ/tA41Se+18X2yX3T89EN4XkT14cLFarnO/W8cUr7qJvS" & _
                                                "2qePfOGKopEvtBf+XF6EOnRcP8ewP/9X1LVPX5o+fNKlRcOfvlT48cylRahf84dLGnycj2/oPsQM25Mbuo1+4uOIDlaHrRhH/fGNXcfUBbR77YWeODaFhEVtDo5I3xISXrQ5JEL4Q+NUzw22/3zWegaF5+cGhws/gsLynXrO/6/2kemXdUifd1lQ0bzLg4UfNI76q/9l/zWaTtqX2nWY93K7oKKXLw0SftA46i9dcnmDjzM3pPsQM2x5od1GfxwR0cHqsBXjqOeGdB1TF9KGX6PvNl0VdXhL3/QjW6KLCKFyeHPfz1A/lNuLfZx+A94QhZdP3BBewwfTC+qEx06Y4Dnxgy5bH9nQZcKDKzq3I9oytEMdIbXjicCNnNq+fMzLV9bc+koHcducID8wjjp0dgKKYFL4ym/438tqbkq7nAl7e4Fx1K975lKPVUBTc1MufGJ912eNYG7s5mgZF/RGODd2tbzAF0JEYSmnANb4h1KG6sHhHjth" & _
                                                "guemoC5bc4O6TPjoks7tiLYM7VDfRCG144lgZlweXP765aE1b7YPFW9d0dEPjKNOQfXYCSiCOfPSDuWvXBpck35ZkH/YCYyjPvPSII9VQHNTUi7MDe36rDdkzl4j6NG3JdT6NTKCublv+ZHNfWu4YNZC9cNboj1cQOvtAJrFMrkwWgG96qGCGdEXTC6U9YAOes5H5qapl2dS8NhQqkAHPecjM/zpdpk3PncZG0oV6KDnfP6U2yvInAH/+H6kowXQ0Nf25vK9uSHhmXwYeaDnfGSMGdEbTC6U9YAOes5HJuOyoMyF7UPYUKpABz3nI0OzYuYrFqFUgQ56zic3rFdQ7QzY0dlrBL2u98iWvplsGC3p6/c46+0Amg2LuRBaMXFDmPYCSpgNKXhWM6ZKu0fXh2kvHkUBKbGaMVWgI/0+zkeGwlZiNWOqQDf8mXbsf/uTH3WPMwNGs6izCz2R3uyFD6ehGbGYC6EVW4L1rxEOWSl4VjOmCs2g" & _
                                                "YdqvjGVcHlRsNWOqQDfvsmDt45zZrkOx1YypAh2Fk/XM7dg9ri5gzl4j6Ot6+deIZs0SPoRW9PV7nPV2ABdAHaqHCt5TMiG0BHrVQwUB4YJoBfSqhwreU8oB1AG96gG8p0e8AePqOsxe+HB1LoA6VA8VvKdkQmgJ9KqHCgLCBdEK6FUPFbynlAOoA3rVAxinSHwB4+o66sLJv0Z8ABtG9ai3A7jw6VA9VNxwOsMNpzWtKpzqigUufDrkfvipnk0NJ+eJcHAhtAJ6uZ/zbGo4Tc9AhpN7nFz4dMj9nGdTw8l5IhxcCK2AXu7nPJsaTtMzkOHkHicXPh1yv+EpDwAufDpUDxV35nSGHE6uzoVPh+qh4s6czpDDydW58OlQPertAC58OlQPFTecznDDaU2rDufEDeH7uQBaAb3qoeL401rScz4yI6e2L3X0aS3pOR+Z4c+0K3X4aS3r2ezhDI4o5QJoCek5H5nm+LR23mVBpc4+rQ3SPs6Z" & _
                                                "7TqUOvm0FnrO5yyEs1QNnwa/x1lvB9Cs5Z7nZMKo0tB5zic2RIaZAWvKeU74cBr3PKe9cDZ4nrNLZFhdwJpwnpN8OE3znOd8r0uPiesjKrggqtCsWWlnNQ80mA19AQ3kCqFKBM9qBsU46tDZWSGEFT80G1YgeFYzKMZRv+6ZdpXWK4QiLzID1ugVQusjnvvDxl7GhZ5UsDpnc3B4JRdEP0hnZzUPNMbKH29AA7JCCCt+Mi4PqkTwrGZQjBvBJJ2dFUJY8UOzYSWCZzWDYhx16CxXCEVGXlQbsEauECKe29OLf42w4ufI5r4VfBBV+lbaWiEEEI7H1ofnTNwQ4VEDCSZuDPcYdRshMvF6RqQb4aP3lCpYc4u6E09fQHMIDxekURinup1gmiBw1z1zac51ky714D2lCsZR166t3Rgx0gwo1szaWVuLNbh1PRGWF/gCvoDm0Azm4UK5JTjcg7qdEJl4PTGDInz+a2ux5taoO/BE4OhwNWf+" & _
                                                "ZcEeLkjGONWdrK1F4F6+tEMOlufhPaWKd7xDjnZtbceIkXUB7TrGztparMGt7aF+TmuCwB3e3DfHWMbHhPLwlr7lqNteW+sSGBYgcBsjxteGzQkbIiagn/N1CRwI3ObQiPF1AXVAaMQEqyAHCnbQJTDge5lPfNgt+YmN3kXwOozvfX7YNdH9PudZBN/n7NgtuW4RfMPg+5z4/qf7fc5fCDhsda+EcH6DWbDFXglBXQERCFxPvtZYXE++1ljOtSc7yHGuH6hdXE++1lhcT77WWJx4soOtGjq0PPFJQuixT/rFH8mPSVE5trlfPOrQsf0uzc6DhaLt/e//HHr36h/j711zIkXl7tXH4lGHjutvKbCDrZPUNsfzYhKP5sVMOrIlJlUHdMe2xCagj/dzCTSpQrS5Z+2PiXfnnJx0d86JVD0nJ933zo8J6OP8znfYwdYGhSz46JaYCVwIdaAP/ZyvS+C4752q4HtWH5/Ah7Bh0Id+zvd8hh2s" & _
                                                "vTBRfp+8I1v6YgVDPQ5v7pPf0IWJODa8Fhb119ci0nNfjyj6+PWuQmUTjaO+McP+Ce5Vs8Ki3nklLH3d3LDda+eECz/mhhehvvqVhj2P5scMMcN2LC9m9NFt8R0sD1tpHPXDeTFjagNK/azWx+2rjkXdver4vLtXnyiiXxbhxyoap/q47OO2/9tNz7tyTuaNW3VcqNy16ni+U89Br34XNWj+0fSrFxwrGjz/qFAZROOop2Qcsu2ZOKsyauArB9IHzDn42cA5B4TKgDkHPkM94ZWvG/S8+52TQ+i58oZtzcnRd797tIPVYSvGUaeZc4zZg35Oa1JdXR156lRNenXN6SL8TVOV6uqaIqNeXW17sYTp+XN1df7Pp6qFStXP1fkNefoNIJhH8vrkH93cd8KxT8N+S7Rl+C3qCKmdgCKYmxZ2Lc9/s1vNp5ndxY5lkX5gHHXSeewEFMFcNzfcsyEjooZC7Rd2gHHUobMOaMqFh7fEPmsELb+f" & _
                                                "o2Vc0PvCaXmhJ4Ro3KqTHgpgjRFEK6g+btUJj50wwfPOlce3EhPGrjjejmjL0A51hNSOJ4JJASwf8vrRmmGLjopr3zjmB8ZRH7TgiMdOQI1gzjlYnpz+fc2gjCN+YQcYR510HquApuaKC+/JOfEsPU+pd68+7ug1gh5996w+afkaIUQUlnIKYY0aSoUa6OwEFJ5VP/+8lYI58fjx40HEfzIEoU66fM6z3g7AjOgLJhfKehgBJb3qoZK7MGLp1je7saFUgQ56zkdm7ZywTAoeG0oV6KDnfI592j/InAEP58Y6WgANvdn7006+d1zO8cx6IdQAPecjgxnRF0wulPWADnrOR2bwgqNLh77Oh1IFOug5HxkKXGZK+iE2lCrQQc/5PPBBVRA9N8YMeO+Knxy9RtDX9r7P91LmljJBtAR6zkfGN2MimFwo6wEd9KpHvR2A2ZCCZzVjqvwWetVDhQ5Zi61mTBXoNr0Wrr3Q07o54cVWM6YKdOvm" & _
                                                "8Be5Ora1f5wZsHKHF+OC3uyFD6ehQ8wShM4upNf+t2M2pOBZzZgq7e5aeVx7wTQKSLHVjKkC3eD5x7SPk8JWbDVjqkBHh76s573rfoij58YbsFzh6DWCvraXfDgNBa5YDWBDUDi1/+04ZKXgWc2YKkE0c/q9RvV2AN5TMiG0BHrVQwUB4YJoBfSqhwreU8oB1AG96gGMUyS+gHF1HWYvfLg6AucU1UMF7ymZEFoCveqhgoBwQbQCetVDBe8p5QDqgF71AMbpEV/AuLoOsxc+XJ0LoA7VQwXvKZkQWgK96lFvB7jhdIYbTmvccPJB5GDDqa5YaGo44ad6IhxcCK2AXu7nPJsaTtMzkOHkHifC5hS5n/Nsajg5T4SDC6EV0Mv9nGdTw2l6BjKc3OPkwqdD7uc8mxpOw1MeAO7M6Qw5nFwdYXOK6qHizpzOMHtb3MypDrjhdIYbTmvccPJB5LAVzub5tLZrqaNPa0nP+cismxO23+GnteyF" & _
                                                "yJo7nONyTpQicHaBnvORcfxpLek5H5nBC46WOvq0lvScj8yAOQdLHX5ay3o2dzgpGPu5AFoBPecj4/TTWuhVj3o7oLWd5/xhW3yYGbCmnOeED6dxz3M2/Tzn/Wt/CKPnxhuwJpznhA+naTHnOb3XPQnsCiGs+KGQVCJ4VjMoxlHftLBrpZ0VQljxQ7NhJYJnNYNiHPV1c8MqLVcIfZFykRmwxq4QIp4Te65hL/SE1Tl0WFmJ4OmAzs5qHmgwG/oCGpAVQljxM3j+kUoEz2oGxbgvmJV2Vghhxc+AOQcqETyrGRTjqNOsWWm1Qmh8rriInh8jYI1dIUQ894eNgn2NsDrn1KnqCi6IKjRrVtpcIRSFlT++gAZmhRAwAkozIsKH95R+5PfJQ93J2loEbtPCiBws4+OCRLOlB3Una2uNgM7tkvPu3HAP3lOqYBx17dravH4jzYBizaydtbVYg2v2oJ/V+vAG9EQOHbJ6EEIGj1G3ESITaDEj" & _
                                                "GuGj95QqWHOLuhNPb0CP5hDlcoBMBr123EO3OU7W1iJw9F4yZ+Dcgx68p/Rj7rcYz9GtraVD0pH0PHkDijWzNtbWGmtwfT3o57QmRkCrq3MogB41kKDaGK/OsRNME8OTZkQcsuI9pR/VDtfWtkoKH2x7ND92vBk2J+BbKehnfV0ChhG4d34YXxdQ++BbKVZBPp9hB1snqW1O5MUmHzUXwWvA9znx/U/3+5xnj1Qh2ty3+u/J96z5wbsIXsvJSfj+J/o4v/MddrBVQ4etxpUQNltcCeET90oI5xrMgt4rIRyzuBLCj63rSgjqCohA4HrytcbievK1xnKuPdlBjnP9QO3ievK1xtJSPH+JsIMtCBy2hJ4+fTr+59OnU1SqaBx1n47rd2lmoh98sG3C4MGhCQlXx8clpaSo9KNx1KHj+lsz7GALoA0FMrH69OlJP1dXp+qAjvQJ6FN8XJqJ1NTUNnEDBiX2Sxw0qV9CSqoW0sUPHJyAPs6v" & _
                                                "NcIOnucEn6qpmcCFUAf60K/4uQSYgQOvCe6flDKBDaEG9KGf821tsIN7v6yM2l9+YN6Xld8WfVlxUDAUob7vq4ZPHMt8tvfLqC9KPen79lcUfVFaLlT27i+ncU/6nn1fNeh56vTpIWbYqv/5z9FVVVUdaNzqkKgt6lXV1WNqA0r9jK6WPv0HRMbGJafHxKfsjk1IEQyfoR4bO8j2yehaz7jkvJj4ZKESHZ+c79QzJCQsKrhTWHpIp7CikM5hwg8aN+qk4/o52rdvH9n+yg7pV1wZUtT+ymChcvmVQTTeIZ10DXrS4eqQ2rAlDBodP3hwB6vDVoyj3j8hZYzZg35O29rwG0Awv6o4WF7+9aGarw9+L7759rAfGEedQuqxE1AEc29peXlJWWUNhZoLu8A46hRUTwMBvZAC9qwRslOnHC3jgt4I9OnT/0P77IWejBDFJ5dTYGqYUNaCOuGxEyZ4Rsclb42OHzghOjq6HdGWoR3qMXED8+x4" & _
                                                "InAhnbuUdwrvWtOla6QI6xblB8ZRp6B67ATUG8zg8g4hHWuCO3XxDzuBcdRJ57EKaEpKyoVxCSnPGkFLSnH0GkHvC6jla9Sa8BugoGSWf/MdG0oV6KBXPVT2llYsLSn72i+QHNBBz/kQWItozIA//eRsATT0Zi/ts70xCSlLuTBaAT3nI4MZ0RdMLpT1IR30nI8MzYpLO4V3Y0OpAh30nI/M5VeEZHYI6cSGUgU66Dmf/ikpQeYMGJuS4ug1gt7spZA76v0l4jfwZcW3xVYzpgp0dOirvdgRHbIWW82YKtCRnvWsqamJkwLm6EJPxL+avfBh6hfQoWwxF0IraEbU/rfjkJWCZzVjqrSL7p+s/XoXha3YasZUgY70+zgfmfZXBhVbzZgq0LW/Ipj17J+YEicFzNFrBL3ZCx9O05rwG0BAuCBaAb3qoYL3lHIAdUCvegDjFEldOP3qOsxe+HB1LoA6VA8V4z0lH0QW6FUPFQSEC6IV0Kse"
Private Const STR_RES_PNG2           As String = "KnhPKQdQB/SqB/CeHvEGjKvrMHvhw9VbE34DCAcXQiugVz1U3HDyQeRww+mG08RvtQbCwYXQCujlfvipnk0Np+kZyHByj5MLnw65n/Nsajg5T4SDC6EV0Mv9nGdTw2l6BjKc6mNsbfgNIBxcCK2AXvVQcWdOPogc7szpzpwmfgMIBxdCK6BXPVTccPJB5HDD6YbTxG/gq4qDpU4+rYVe9VDZW1pe6ujTWtJzPs0ezrjkUi6AlpCe85Fx/Gkt6TkfmeDOYaVOPq2FnvORaX9lUKmjT2tJz/m44QwcfgMUkPP2POcpIcLMgDXlPCd8OI17nrPp5znjU1LCzIA15TwnfDhNa8JvACt+vqr4tgLBs5pBMY46zZqVdlYIYcUPzYYVCJ7VDIpxI5j7yysbWCF0UW3AGrlCiHiO9tkLPWF1Dh1WVnJBVDF0NlbzGJ5xA/N8AbWaQRuxQiisAsGzmkExjnpo57BKmyuEouhQtRLBs5pBMY46dA2s" & _
                                                "ELrIDFgTVgg9d801/AXTWhPsoC+gOV9WfuvhglT29bce1J2srUXg6L1kzhf7Kzx4T6myzzueo1tbW/3Pf440A4o1s3bW1mINrtmDfkZXC8IRG5+cQ3jYUCYMKjPqNkJkYngaM2hyPt5T+hGXnIe6E08ELrhzlxwKjUcNkZfwcqNuI5gmRkA7BOdQ+Dx4T8ngMeqatbX9BqSMNAOKNbP21tYOGm32oJ/TtjbYwfOctqdqasabYXOC71sp7vcGm5kHKXD9kwaNrw2bA/CtFPRzvq0NdrAFgO9zJlefPm0sgtfh+z5nIvoUH5dmAt/LjB8wKDkuybcIXkfioEn4/qf7fc462MEWBP4PG4orHhif5CrgCgmo+3Rcv0szg8NWXOkAVzzAJ7AquEKCeyUEHnaQozlWa7iefK2xtGbPXyLsoEogn0zXyxmulzPOV6/GwA66uLRWrnlJ/EfS9KqQxOkiKmGqiMftwDQRjHFO35ywgy4urY1UIdok" & _
                                                "TROd46eK5LjnT6eoYDxp2qnOqaln70NFdtDFpTWRki4uik8T/bhQqiSm1cQkzhAXcz6Bhh1cs2bNFe+8+8He5TlrTr+9fJUwwT7GUef6OBYtWnTxsuWrMpYtX+mRvbCPcdS5Pg5ol2RmZSzJzPYsWZYtasE+jTvx2rp168VbdxRmbN1R4PlkR6Ewwb4xTnWuj2PEohEXj1h+c8aIFaM8N60YJUywb4xTnevj6Pq7xIvjnxqQEf/UQE/CUwOFCfYxjjrXx9F1RNeL+4yLyeg9LtbT5+5YYWLs0zjqXB/Hn9p1vTi1Q/eM1KDuntSgSFEH7dM46lwfx4Tgrlc83T167x979Dv9x97xohbaxzjqXB/H/Vd2CxvfucfxBzp2P/O70O7CBPsYR53rMxm7QvwqaarozwXRCujPxgzqN7Bhw4f/vXrtezVFn+8TB787LE7++I9asF+0Z69Ytfbdmvc2brxd7VXJylqZsCJnbVlDXtk5a8qWZmfj" & _
                                                "b9ezHiZvZmYmLstaWfZxbr7YUbBLFH2+txbsf7wpT6D+VlZWH65fJm97QeK2nbvKPJUHxfdHT4qTf6+qBftlFQcE6tsKCrReN2XdlDBy5aiyWz++Xdyx7S4xrvDuWrB/619vF6jfkD1K+98Y9/iAxKSnB5Zd/9aNYvTGMeK2bbfXgv3r3rpBJP55YFn8k/Hax9X7zpjE6Pv6lSVNShYpUweLwbOG1oL9pEkDBeq97orWPq7/e2XXxMlBkWV/CYoSGR16iNeD6sA+xlFPvbK71utPEX0emtQzvmZ53DUib8ANoihlZC1baD8rfqh4pmdczdPhfbS/X4907vHSQx2jzjwf2lO80rGXyOh4VS1/6dhTTOnYQ6A+sXPPSVw/iJ8mwtTwxU8RPSNTxUUXUHBxi31Vg0Ngzi+Q+A2se2/9ll2ffV4vSCqFRXsEdGqvytvZqybb8YKO65eh2XHyR5u2iM/+tk98/kWx+HxvsfjbvhLjFvsY/yh3" & _
                                                "C2ZRrde2nYWTvyr/ul4oVb70VFJAC7VeN2XfPHnsh7eKuwrGWTJ2460COq5fJv6JxMnD37he3PYJBdKC4YtuENBx/TJXjYuenPD0ACOMVyOUs4m/eG+NfSLhmYECOq5fZnKH7lNmd4gSCymMVszuECmg4/plnu7WN29Z3NUUxpssyYwbIqDj+mVoZqx8LjRKzO90lVhAgVTBeGpoDwEd1z9kuvhPOXCxaWJgdKoI4bR90n4ORV3Wo5/TBgq/ATp0rT5As9oJCo4V33z7vaAZsUrtVclcvrLEjldm9qrPuH4ZCl1Jwa4iI5D7Sr8UJaVfiZL9ZcYt9jG+c9duhFPrRYeuJYeP1Z8xVTCD0uGt1mtE9siS2z+5Q9y1k4Jowe1b7xQjsm/WesU/nlQyasMYcevW2ywZtX60iHsiSevVe1xMSfL0q41AXj3nGjFEAvsYT5k2WPQZF6v1er5D95IFFL5FFEIr5lN4p3SI1Hr9MSq2ekvSdWJ3" & _
                                                "8ghLNlP9j1Ex2t8vHLr+hWZNBPH1jr3FIro1wT7GUf99p8j/x/UPmCoi5bD1myKu4nQm/V8QvWR9/FTRjdMFCr8BvB888QMFR0PWilVn1F6VZdmrTnG9KsuyV57k+mXeWpZ9am/JfiOM+8sqBGa+MgK32Mc46vQeVOtFoTvFBVKF3oNqvW7MGnnqzh13CR0jsm7WesU/NuDUrXkUQg10+Kv16n13zKmrX/EF8tVhYmh6Hdg3Qkr1PnfHaL3SgiJPvRHUU+ggndbrqZ6xYvfAG7X8oWes9vfrd6HdKITeML5BYXwThF5l3GLfG9KrxP2kU3tBYpqIkcPWc5q4hNOZoC7r414Q2sP4puA3QIFiA6QCndqrEkgvfPCDAOJws7zyG/H1N9+KygPfGbfYxzjq0Km9KvjghwujCnRqr8qIt28Wd35KAdQAndqrQjOnuGXLrVqgU3tV8MGPEUwEct4wcc38a8XQBdd6b2kf46hDp/aqpNF7ysUU" & _
                                                "Ph3Qqb0qf+jVX+wacKMW6NReFXzwgwAupiAuITKJpb5b7GMcdejUXpCUJpLksEUvaHiZJ+qyHv2cLlD4rYJobKDg05xeCN3+r8qFp+IbI5QH6HD420NHjFvsYxx1NZycV2PDyXnduGykuGP7nVqgk/s4r/jHKJybKYAaoJP7OC+EbihmTAqiEcrXrhXDXh9u3GLfGKe6Gk7OKy2YwhlMAdQAndzHeRnhTLpBixpOzguhQwDfogAilMuILN8t9jGOuhpO00sOGpA1Vqg96mMKJMYDlQcys1eK4z/8pAU6uQ8+zemF0GF2rPzmoBHI774/Kg4dPmbcGgGlcdTPSTi3UQA12AlnHIVubO4tWqCT+zgvY+ak2REzpRnMYQsJX0AxjrqdcL5AoXuTwqcDOrmP83qqVz9RmHS9FujkPs7LDGddMPuIbAK3ZkDdcBLwaU4vhA7vL3EYi9MwCObhoyeMW+xjHPWzHc4bMm8S+EBIB3RyH+cVN5HC" & _
                                                "uYkCqAE6uY/zMmZOXziHveYN5rWLrvMGlPaNw1ub4ZxKoXuLwqcDOrmP83qqJ4UzkQKoATq5j/NC6PD+EoexbxMI5nJfQLGPcdR/WeE8SaHRYDucTK+K3XCW0aHr13QIi8NZOZzYxzjqZz2cSymcWymAGqCT+zivuImJYszHY7VAJ/dxXkY459EMSYew5qxZG07Mnsah7bW2wjktuIdYEtxLC3RyH+eF0BUkXKfFaThxOCuHE/tuOH3Apzm9zHDi/SXC+P2R40Y4cYt9jJ+LcF6/ZIS4Lf92LdDJfZxX3KMUzr9SADVAJ/dxXmY48f4SYUQwTbDvfd9pL5zTKXRLKXw6oJP7OK+nesSKgvjhWqCT+zgvM5w4fEUYEUwT7GP8FxfOYxQYHXbDyfWqtOhwvkXhzKMAaoBO7uO8jHB+RAHUcNbDGdJDZIb00gKd3Md5GeGMu1aLG07jceK2bsANZx32wnmjuHXLbVqgk/s4r/6PJIrRH47R" & _
                                                "Ap3cx3kFMpwzKHTLKHw6oJP7OK8nKXQ7KXw6oJP7OC83nBa44fRy/ZsUzs0UQA3QyX2cV/8JFM4NFEAN0Ml9nFcgw/liSE/xdshVWqCT+zivJ6NixM7+w7RAJ/dxXq0vnFkUqBMUGg3QyX3waU6v8zWc171xo7hl061aoJP7OK/+ExKM5Xk6oJP7OK9AhvMlCl0WhU8HdHIf54XQ7eh3jRY3nMbjxG3dwFIKytETf9cCndwHn+b0Om/DuYjC+TEFUAN0ch/n1X88hfMDCqAG6OQ+ziuQ4XyZQpdN4dMBndzHeT0ZGS12xA7VAp3cx3m54bTADacXfEtk7F9v0QKd3Md59aPQ3fz+KC3QyX2cV6DDuZzCp8NOOJ+g0H0aM1QLdHIf5+WG0wI3nF6GL6RwfkQB1ACd3Md59XuYwvkeBVADdHIf5xXQcIZSOEMpgBqgk/s4L284h2hxw2k8TtzWDbjhrMNOOK99/Xox5sOxWqCT+zivfg/H" & _
                                                "i5Hv3qwFOrmP8wpkOGdS6FaE9tYCndzHeT3Rva/YHn21FujkPs6r1YVzydsrxZHjf9cCndwHn2b1Ol/D+RqFcyMFUAN0ch/n1e+heDFyHQVQA3RyH+cVyHDOotCtpPDpgE7u47yMcPYdrMUNp/E4cVs34IazDjvhHLbgOvZ0hwp0ch/nFUuhu2ntSC3QyX2cVyDDOZtCt4rCpwM6uY/zerxbX7Gtz2At0Ml9nJcbTgvccHoxwrmeAqjBVjj/m8L5DgVQA3RyH+cVyHDiagI5FD4d0Ml9nNfj3fqIbb0HaYFO7uO83HBa4IbTy7D5w9nTHSrQyX2cV+yDcWLE6pu0QCf3cV6BDWcvsTq0jxbo5D7OC6H7pHeKFjecxuPEbd3AkrdXiMPHf9QCndwHn2b1Ol/DmUHhfJ8CqAE6uY/zMsKZQwHUcLbDiSvbraFfeB3QyX2c1+NdKZxXUQA1QCf3cV6tM5zHKDQabIeT6VVpyeG8Zt5w9nSH" & _
                                                "CnRyH+cV8/s4ceOqEVqgk/s4r0CGcw6F7h36hdcBndzHeT3WtbfY2itZC3RyH+flhtMCN5xerkmncL5LAdQAndzHeRnhXEkB1HC2wzmXQreWfuF1QCf3cV6PRVA4ew7UAp3cx3m54bTADaeXoenXsqc7VKCT+zivmAf6ixtW3KgFOrmP8wpkOF/teJVY17GvFujkPs7rsYirRH6PAVqgk/s4r1YXzrcoKN9TYHRAJ/fBpzm9zttwvkrhXEsB1ACd3Md5GeFcTgHUcC7C+S6FT4etcIZTOKMogBqgk/s4LzecFrjh9DJk7jD2dIcKdHIf5xV9f39xffYNWqCT+zivgIezEwVQg51wTqTQ5UUlaYFO7uO8Wl84l1GgjlJoNEAn98GnOb3O23DOoXCuoQBqgE7u47yMcGZRADWc7XCmU+jeo/DpgE7u47wmhvcSeZGJWqCT+zgvN5wWuOH0gqumc+ciVaCT+ziv6N/1E9e9fb0W6OQ+ziuQ" & _
                                                "4ZxHoXufwqcDOrmP85oY1kts6Z6gBTq5j/NqdeF8c9lycejoD1qgk/vg05xe53U4mXORKrbCeR+FM5MCqAE6uY/zCmQ4M+gX/INO0Vqgk/s4LyOc3SiAGtxwGo8Tt3UDbjjrsBXOvwxlz0WqQCf3cV59KXTDl16nBTq5j/MKZDjn0y/4egqfDujkPs7r0S49xeau8Vqgk/s4LzecFrjh9HL1bAoncy5SBTq5j/Pqey+FcwkFUAN0ch/nFehwbqDw6bAfzjgtbjiNx4nbuoE3MylQRyg0GqCT++DTnF7nazjxdy65c5Eq0Ml9nFffe2PFtW9RkDRAJ/dxXoEM54JOvcXGztFaoJP7OC+ELjciTosbTuNx4rZuAEH5jgKjw244uV6VFh3OmRRO5lykCnRyH+fV9x4K55sUJA3QyX2cVyDD+VqnPuLDzjFaoJP7OK9HO/cQueH9tUAn93FebjgtcMPpZdDMIey5SBXo5D7OywjnYgqShrMd" & _
                                                "ztcpdB9R+HRAJ/dxXkY4w/ppccNpPE7c1g244azDVjhfpnAy5yJVoJP7OC8EZdgbFCYNdgIVyHAupND9lcKnAzq5j/N6pFOU2NQlVgt0ch/n9YsPpzrQ2EBxBNLLDKf5h4zUcFr9ISOOxoaTI+Wlq9lzkSrQqb0qRjgXUZg0qIHiMMNp/iEjNZxWf8iIA6H7mMKnQw0nR2PDyWGG0/xDRmo4rf6QkUl8mkiQg+b0j+ein9MFCr+Bpdk51Z5vvmdDZOL5+pCATu1VWZy5vMSOF3Rcv8ySzOySfaVfGSFk/wQgje8r+VJAx/XLbN1RUHLk+A9sIE1Qh47rl0mZMbgEH9Jct4xCaAHq0HH9Mr3HxZQMmTeUDaTJkPShAjquXwaawbOHeMPJ/QlAGh88a4gtr4Ude5ds7NSXDaQJ6tBx/TJPdIysfr9zNBtIk/eoDh3XL/NAaLczC0J7GiG0+hOAqEPH9Q94QfSSw+b0z86jn9MFCr+BFTnr" & _
                                                "Sgv37GODZFLw2V4Bndqr8ubS7IV2vKDj+mXeysxa9OnOQuPvcJp/PBezpvnHczG+fUeBgI7rl9m6o3BhOem5UJqgDh3XL5My4+qF17w+jA2lCerQcf0yfcbFLEyeOogNpcnAtBQBHdcvA83AKYPq/fHc2lnT98dzB6Ta81rUqc/CtZpAoQ4d1y/zdMeo0sWd+7AeJm9QHTquX+ah0O4nX+zoDaf5x3O9s2YfYx/jM6gOHdefNE10lsPWXxO2flPEVbIe/ZwuUPgNrFz5bqfslWtPf178FRsmjGevWlu9evXqULVXZd68ZZdkZq8ub8grMzunePHixb/h+mUWL17zm2XZq8p37fmb+PrgIWO2RDCNWZP2MU51W175+fmXfLqzqPwA9XLBPPDdEbFt567i3NxcrVfStKRLUl4eXI7gcMEctohmzZcGF6ekpmi9et7R85K+9/UrJ71fKMGgFweL6Ptii3uP7G3LK/q+/uUDpw2mWdI7WyKY" & _
                                                "3llzuEimcarb8poX0vOSxZ16l1vNeO/TzPlGxz7Fszvqvf43JLzTnztGnn67c1/WK8sIZo/qP3Xsqv39eiKkc++HaFacbQTUO1simN5Zs49xwbGHOkaeefzKcHy65Nd/wwLx7wlpYqAcuOjUqhBOG5f6c0dZFz/1dPINqeLfOW2gYAfXrFlzxfKcdXszs1adxvtBE+xjHHWuj2PRokUXL16aPf/Npcs9shf2MY4618cB7ZLMrAw6dPXgvWUt2KdxJ15bt269mGbGDDp09eC9pQn2jXGqc30ciTMSL6bD1ozkF6/2pLx4tTAx9mkcda6Po+uIrhfTbJbRe1ysB+8HTYx9Gked6+MIpFf6pZEXLezcZ/7CTr09iyhAJtjHOOpcH8cz7bq1f6Zjj7106Hoa7y1NsI9x1Lk+jkfadesyPiTy+P0UUry3NME+xlHn+kzipol6oQOxU2piY1NF+7A54te4xb6qSZou2BAHEnbQxaW1kJoq2sSn" & _
                                                "iX5q+BoiMa0Gf2XpX1SvQMMOuri0JlJmi99wIeSInyqSE2cI20ccTYEddHFpbWAGTZp2qjPCZxVKfAAEHdffHLCDLi6tlWteEv8xME0EJ04XUQlTRTxuk6ZXhWCc0zcn7KBKIFdBuF7OcL2ccb56NQZ2kKM5HqjrydcaS2v2/CXCDrq4BBIcEuLQ0DxUxHs485ARh5Dn4pCxJcAOurgEglSBD1mE5YcsJt4PW06d1Q9bWgL1dy64IODnblqKZ3PQUv7bm8MzJV1c1Jjzh2frNEVLoO4f9AKZVHxz6I6Kb44UVB44UlV58Kjwg8ZRL6s4dIfcJxurnjOzi+dOmLPnp9teKBSjp/iDcdRnkE7ua8jzjR1Pzn1+07U/PbWxh3hsQ4QfGEf9te1/aNDTBLXj77009/hbD//0w7xR4sdXb/QD46gffvclw5PzAaiZ7H7kybnbkq/9aUtYD7E5JMIPjKNe8EjDj1OuHdr9zNzvt4/+6XBevDiy" & _
                                                "JdoPjKN+cPf/2Pb8MOO7uSsmfffTWw8fFIt//60fGEd9A+nkPtVz7Arxq6Spoj8XQB3oc2dQL94f0hNdVvHd+AOHjouqn6vFmTNYzO+/YRx16KCX+6ns5zl12b4sLpBWQC/3c54Z2ydmcYG0Anq53/SUOfzO1CwukFZAz/nI97PzwYlZXCCtgF7u5zwP7vpTFhdIK6CX+znPD+Z+l8UF0gro5X7TE8RPE2Fq6OKniJ6RqeKiCyi4xn3TLfYxrmqbe0F5S8H7o+5JblP+zeE9p6praFi/QQc9+kwPGvbzfPiVPVVcCK2AHn2mB+c5OXdoFRdCK6BHn+lhesqcePPBKi6EVkDP+Zj3QbT5JGloFRdCK6BHn+nBeX6/bWQVF0IroEef6cF5Ln/6uyouhFZAjz7Tw/QcMl38pxy02DQxMDq14XWofdJ+DoVO7oMPp21NeH/UvUi/qjhw9KTVjKlu0EGPPtODhv08b5+66wwXQiugR5/pwXk+" & _
                                                "9WGvM1wIrYAefaaH6Snzw/zRZ7gQWgE952PeB/GrvIheZ7gQWgE9+kwPzvNwXsIZLoRWQI8+04PzXDL+2zNcCK2AHn2mh+k5YKqIlEOGr1mZtYbA17Xkvvipohuna014f3ifYPxf8EK8p3SyQY8+X7/6wp/3njKN8aTNzwf+vvtpdc9nYpqIkUOm+wKzifpF5rgXRF9O15rw/mj6i9TW1x/IF/6seMo0xpM2Px/4++6n1T2fSWkiSQ6Z7tIfJuolQODD6VoT2PDTfJHaNvVFIrDfYjxpq30yGuNJm/xkNtvj9P37vPeUAwZIWvv86FB74adqWhPeHwF8kWioRXnKNMaTNj8f+Pvup9U9n04vmmVyti+e1RLw/mghLzz+7RsLmKdMYzxp8/OBv+9+Wt3z6fSiWSZn++JZLQHvD+VFcvJprd0X/nz1lGmMJ21+PvD33U+rez6dXjTL5GxfPKsl4P1R9yJd2MjznA1+Eng+e8o0xpP+6ecD" & _
                                                "f9/9tLrn08lFs0zOxcWzWgLeH94XCVxYvL/yUScrhKBHn+lB5RblKdMYTxry8zHvg2iVz6fuolnQnOuLZ7UEvD/qXiScVP4/e7746h7P14cLG1pbi/rnX+y/F3pfX70XydwnzmtPGWicelr4NOlxNoenr6/ef7u5TwT0+WzMRbNMztbFs1oC3h91L5JxiEP8mvg34t8bAHXooDcObUCtcQvxbA7M+yBsP07OR4a2gP+3m/tEwJ9PJxfNMsFXx9xvpdRR94/6LxT+j4gnH2/4rUAdOssXyBwnzmvP5sC8L8LW4+Q8VGgL+H+7OU4E/PnUXTTLxPt9zrN78ayWQN0/6l4k+cXSUa9HNm5Jns2Bep8E97hq4TxUaHPk6aNez9nwVFEvmoVAnuuLZ7UE/AeUJ94OqoeKqrcD5yOj6u3A+TQ36mOwguu1Qu21A+cjo+rtwPm4BA520ER9MWQ4vR1UHxlObwfVR4bTnyvUxybD6e2g+shwejuo" & _
                                                "PjKc3qV5YAc5aKMbvtZYXE++1lhas+cvD3HB/wepoa0nOlnfLQAAAABJRU5ErkJggg=="
Private Const LNG_RES_WIDTH         As Long = 231
Private Const LNG_FLAT_TOP          As Long = 0
Private Const LNG_FLAT_WIDTH        As Long = 21
Private Const LNG_FLAT_HEIGHT       As Long = 21
Private Const LNG_BUTTON_TOP        As Long = 84
Private Const LNG_BUTTON_WIDTH      As Long = 19
Private Const LNG_BUTTON_HEIGHT     As Long = 52
Private Const LNG_CARD_TOP          As Long = 136
Private Const LNG_CARD_WIDTH        As Long = 21
Private Const LNG_CARD_HEIGHT       As Long = 23

'--- design-time
Private m_eStyle                As UcsNineButtonStyleEnum
Private m_sngOpacity            As Single
Private m_sngAnimationDuration  As Single
Private m_uButton(0 To ucsBstLast) As UcsNineButtonStateType
Private m_sCaption              As String
Private WithEvents m_oFont      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_clrFore               As Long
Private m_bManualFocus          As Boolean
Private m_oPicture              As StdPicture
Private m_clrMask               As OLE_COLOR
Private m_bAutoRedraw           As Boolean
'--- run-time
Private m_eState                As UcsNineButtonStateEnum
Private m_hPrevBitmap           As Long
Private m_hBitmap               As Long
Private m_hAttributes           As Long
Private m_sngBitmapAlpha        As Single
Private m_hFocusBitmap          As Long
Private m_hFocusAttributes      As Long
Private m_nDownButton           As Integer
Private m_nDownShift            As Integer
Private m_sngDownX              As Single
Private m_sngDownY              As Single
Private m_dblAnimationStart     As Double
Private m_dblAnimationEnd       As Double
Private m_sngAnimationOpacity1  As Single
Private m_sngAnimationOpacity2  As Single
Private m_hFont                 As Long
Private m_bShown                As Boolean
Private m_hPictureBitmap        As Long
Private m_hPictureAttributes    As Long
Private m_eContainerScaleMode   As ScaleModeConstants
Private m_pTimer                As IUnknown
Private m_hRedrawDib            As Long
'--- debug
Private m_sInstanceName         As String
#If DebugMode Then
    Private m_sDebugID          As String
#End If

Private Type UcsNineButtonStateType
    ImageArray()        As Byte
    ImagePatch          As cNinePatch
    ImageOpacity        As Single
    ImageZoom           As Single
    TextFont            As StdFont
    TextFlags           As UcsNineButtonTextFlagsEnum
    TextColor           As OLE_COLOR
    TextOpacity         As Single
    TextOffsetX         As Single
    TextOffsetY         As Single
    ShadowColor         As OLE_COLOR
    ShadowOpacity       As Single
    ShadowOffsetX       As Single
    ShadowOffsetY       As Single
End Type

Private Enum UcsNineButtonResIndex
    ucsIdxFlatPrimaryNormal = 0
    ucsIdxFlatPrimaryOutline
    ucsIdxFlatPrimaryHover
    ucsIdxFlatPrimaryPressed
    ucsIdxFlatPrimaryFocus
    ucsIdxFlatSecondaryNormal
    ucsIdxFlatSecondaryOutline
    ucsIdxFlatSecondaryHover
    ucsIdxFlatSecondaryHoverOutline
    ucsIdxFlatSecondaryPressed
    ucsIdxFlatSecondaryFocus
    ucsIdxFlatSuccessNormal
    ucsIdxFlatSuccessOutline
    ucsIdxFlatSuccessHover
    ucsIdxFlatSuccessPressed
    ucsIdxFlatSuccessFocus
    ucsIdxFlatDangerNormal
    ucsIdxFlatDangerOutline
    ucsIdxFlatDangerHover
    ucsIdxFlatDangerPressed
    ucsIdxFlatDangerFocus
    ucsIdxFlatWarningNormal
    ucsIdxFlatWarningOutline
    ucsIdxFlatWarningHover
    ucsIdxFlatWarningPressed
    ucsIdxFlatWarningFocus
    ucsIdxFlatInfoNormal
    ucsIdxFlatInfoOutline
    ucsIdxFlatInfoHover
    ucsIdxFlatInfoPressed
    ucsIdxFlatInfoFocus
    ucsIdxFlatLightNormal
    ucsIdxFlatLightOutline
    ucsIdxFlatLightHover
    ucsIdxFlatLightPressed
    ucsIdxFlatLightFocus
    ucsIdxFlatDarkNormal
    ucsIdxFlatDarkOutline
    ucsIdxFlatDarkHover
    ucsIdxFlatDarkPressed
    ucsIdxFlatDarkFocus
    
    ucsIdxButtonDefNormal = 0
    ucsIdxButtonDefHover
    ucsIdxButtonDefPressed
    ucsIdxButtonDisabled
    ucsIdxButtonGreenNormal
    ucsIdxButtonGreenHover
    ucsIdxButtonGreenPressed
    ucsIdxButtonRedNormal
    ucsIdxButtonRedHover
    ucsIdxButtonRedPressed
    ucsIdxButtonFocus
    
    ucsIdxCardDefault = 0
    ucsIdxCardPrimary
    ucsIdxCardSuccess
    ucsIdxCardOrange
    ucsIdxCardDanger
    ucsIdxCardWarning
    ucsIdxCardPurple
    ucsIdxCardFocus
End Enum

'=========================================================================
' Error handling
'=========================================================================

Friend Function frInstanceName() As String
    frInstanceName = m_sInstanceName
End Function

Private Property Get MODULE_NAME() As String
#If ImplUseShared Then
    #If DebugMode Then
        MODULE_NAME = GetModuleInstance(STR_MODULE_NAME, frInstanceName, m_sDebugID)
    #Else
        MODULE_NAME = GetModuleInstance(STR_MODULE_NAME, frInstanceName)
    #End If
#Else
    MODULE_NAME = STR_MODULE_NAME
#End If
End Property

Private Function PrintError(sFunction As String) As VbMsgBoxResult
#If ImplUseShared Then
    PopPrintError sFunction, MODULE_NAME, PushError
#Else
    Debug.Print "Critical error: " & Err.Description & " [" & STR_MODULE_NAME & "." & sFunction & "]", Timer
#End If
End Function

'Private Function RaiseError(sFunction As String) As VbMsgBoxResult
'    Err.Raise Err.Number, STR_MODULE_NAME & "." & sFunction & vbCrLf & Err.Source, Err.Description
'End Function

'=========================================================================
' Properties
'=========================================================================

'== design-time ==========================================================

Property Get Style() As UcsNineButtonStyleEnum
    Style = m_eStyle
End Property

Property Let Style(ByVal eValue As UcsNineButtonStyleEnum)
    If m_eStyle <> eValue Then
        m_eStyle = eValue
        pvSetStyle eValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Property Let Enabled(ByVal bValue As Boolean)
    If UserControl.Enabled <> bValue Then
        UserControl.Enabled = bValue
        pvState(ucsBstDisabled) = Not bValue
    End If
    PropertyChanged
End Property

Property Get Opacity() As Single
    Opacity = m_sngOpacity
End Property

Property Let Opacity(ByVal sngValue As Single)
    If m_sngOpacity <> sngValue Then
        m_sngOpacity = sngValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get AnimationDuration() As Single
    AnimationDuration = m_sngAnimationDuration
End Property

Property Let AnimationDuration(sngValue As Single)
    m_sngAnimationDuration = sngValue
    PropertyChanged
End Property

Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = m_sCaption
End Property

Property Let Caption(sValue As String)
    If m_sCaption <> sValue Then
        m_sCaption = sValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
    Set Font = m_oFont
End Property

Property Set Font(oValue As StdFont)
    If Not m_oFont Is oValue Then
        Set m_oFont = oValue
        pvPrepareFont m_oFont, m_hFont
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_clrFore
End Property

Property Let ForeColor(ByVal clrValue As OLE_COLOR)
    If m_clrFore <> clrValue Then
        m_clrFore = clrValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get ManualFocus() As Boolean
    ManualFocus = m_bManualFocus
End Property

Property Let ManualFocus(ByVal bValue As Boolean)
    m_bManualFocus = bValue
    PropertyChanged
End Property

Property Get Picture() As StdPicture
    Set Picture = m_oPicture
End Property

Property Set Picture(oValue As StdPicture)
    If Not m_oPicture Is oValue Then
        Set m_oPicture = oValue
        pvPreparePicture m_oPicture, m_clrMask, m_hPictureBitmap, m_hPictureAttributes
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get MaskColor() As OLE_COLOR
    MaskColor = m_clrMask
End Property

Property Let MaskColor(ByVal clrValue As OLE_COLOR)
    If m_clrMask <> clrValue Then
        m_clrMask = clrValue
        pvPreparePicture m_oPicture, m_clrMask, m_hPictureBitmap, m_hPictureAttributes
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get AutoRedraw() As Boolean
    AutoRedraw = m_bAutoRedraw
End Property

Property Let AutoRedraw(ByVal bValue As Boolean)
    If m_bAutoRedraw <> bValue Then
        m_bAutoRedraw = bValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get ButtonState() As UcsNineButtonStateEnum
    ButtonState = m_eState
End Property

Property Let ButtonState(ByVal eState As UcsNineButtonStateEnum)
    pvState(m_eState And Not eState) = False
    pvState(eState And Not m_eState) = True
    PropertyChanged
End Property

Property Get OLEDropMode() As UcsNineButtonOleDropMode
    OLEDropMode = UserControl.OLEDropMode
End Property

Property Let OLEDropMode(ByVal eValue As UcsNineButtonOleDropMode)
    UserControl.OLEDropMode = eValue
    PropertyChanged
End Property

'== run-time =============================================================

Property Get Value() As Boolean
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "400"
    '--- do nothing
End Property

Property Let Value(ByVal bValue As Boolean)
    If bValue Then
        pvHandleClick
    End If
End Property

Property Get ButtonImageArray(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Byte()
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonImageArray = m_uButton(eState).ImageArray
End Property

Property Let ButtonImageArray(Optional ByVal eState As UcsNineButtonStateEnum = -1, baValue() As Byte)
    Dim oPatch          As cNinePatch
    
    If eState < 0 Then
        eState = m_eState
    End If
    m_uButton(eState).ImageArray = baValue
    Set oPatch = New cNinePatch
    If oPatch.LoadFromByteArray(baValue) Then
        Set m_uButton(eState).ImagePatch = oPatch
    Else
        Set m_uButton(eState).ImagePatch = Nothing
    End If
    pvRefresh
End Property

Property Get ButtonImageBitmap(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Long
    If eState < 0 Then
        eState = m_eState
    End If
    If Not m_uButton(eState).ImagePatch Is Nothing Then
        ButtonImageBitmap = m_uButton(eState).ImagePatch.Bitmap
    End If
End Property

Property Let ButtonImageBitmap(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal hBitmap As Long)
    Dim oPatch          As cNinePatch
    Dim hNewBitmap      As Long
    
    If eState < 0 Then
        eState = m_eState
    End If
    Set oPatch = New cNinePatch
    Call GdipCloneImage(hBitmap, hNewBitmap)
    If hNewBitmap = 0 Then
        Set m_uButton(eState).ImagePatch = Nothing
    ElseIf oPatch.LoadFromBitmap(hNewBitmap) Then
        Set m_uButton(eState).ImagePatch = oPatch
    Else
        Set m_uButton(eState).ImagePatch = Nothing
    End If
    pvRefresh
End Property

Property Get ButtonImageOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonImageOpacity = m_uButton(eState).ImageOpacity
End Property

Property Let ButtonImageOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).ImageOpacity <> sngValue Then
        m_uButton(eState).ImageOpacity = sngValue
        pvRefresh
    End If
End Property

Property Get ButtonImageZoom(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonImageZoom = m_uButton(eState).ImageZoom
End Property

Property Let ButtonImageZoom(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).ImageZoom <> sngValue Then
        m_uButton(eState).ImageZoom = sngValue
        pvRefresh
    End If
End Property

Property Get ButtonTextFont(Optional ByVal eState As UcsNineButtonStateEnum = -1) As StdFont
    If eState < 0 Then
        eState = m_eState
    End If
    Set ButtonTextFont = m_uButton(eState).TextFont
End Property

Property Set ButtonTextFont(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal oValue As StdFont)
    If eState < 0 Then
        eState = m_eState
    End If
    If Not m_uButton(eState).TextFont Is oValue Then
        Set m_uButton(eState).TextFont = oValue
        pvRefresh
    End If
End Property

Property Get ButtonTextFlags(Optional ByVal eState As UcsNineButtonStateEnum = -1) As UcsNineButtonTextFlagsEnum
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonTextFlags = m_uButton(eState).TextFlags
End Property

Property Let ButtonTextFlags(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal eValue As UcsNineButtonTextFlagsEnum)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).TextFlags <> eValue Then
        m_uButton(eState).TextFlags = eValue
        pvRefresh
    End If
End Property

Property Get ButtonTextColor(Optional ByVal eState As UcsNineButtonStateEnum = -1) As OLE_COLOR
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonTextColor = m_uButton(eState).TextColor
End Property

Property Let ButtonTextColor(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal clrValue As OLE_COLOR)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).TextColor <> clrValue Then
        m_uButton(eState).TextColor = clrValue
        pvRefresh
    End If
End Property

Property Get ButtonTextOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonTextOpacity = m_uButton(eState).TextOpacity
End Property

Property Let ButtonTextOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).TextOpacity <> sngValue Then
        m_uButton(eState).TextOpacity = sngValue
        pvRefresh
    End If
End Property

Property Get ButtonTextOffsetX(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonTextOffsetX = m_uButton(eState).TextOffsetX
End Property

Property Let ButtonTextOffsetX(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).TextOffsetX <> sngValue Then
        m_uButton(eState).TextOffsetX = sngValue
        pvRefresh
    End If
End Property

Property Get ButtonTextOffsetY(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonTextOffsetY = m_uButton(eState).TextOffsetY
End Property

Property Let ButtonTextOffsetY(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).TextOffsetY <> sngValue Then
        m_uButton(eState).TextOffsetY = sngValue
        pvRefresh
    End If
End Property

Property Get ButtonShadowColor(Optional ByVal eState As UcsNineButtonStateEnum = -1) As OLE_COLOR
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonShadowColor = m_uButton(eState).ShadowColor
End Property

Property Let ButtonShadowColor(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal clrValue As OLE_COLOR)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).ShadowColor <> clrValue Then
        m_uButton(eState).ShadowColor = clrValue
        pvRefresh
    End If
End Property

Property Get ButtonShadowOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonShadowOpacity = m_uButton(eState).ShadowOpacity
End Property

Property Let ButtonShadowOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).ShadowOpacity <> sngValue Then
        m_uButton(eState).ShadowOpacity = sngValue
        pvRefresh
    End If
End Property

Property Get ButtonShadowOffsetX(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonShadowOffsetX = m_uButton(eState).ShadowOffsetX
End Property

Property Let ButtonShadowOffsetX(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).ShadowOffsetX <> sngValue Then
        m_uButton(eState).ShadowOffsetX = sngValue
        pvRefresh
    End If
End Property

Property Get ButtonShadowOffsetY(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonShadowOffsetY = m_uButton(eState).ShadowOffsetY
End Property

Property Let ButtonShadowOffsetY(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).ShadowOffsetY <> sngValue Then
        m_uButton(eState).ShadowOffsetY = sngValue
        pvRefresh
    End If
End Property

Property Get DownButton() As Integer
Attribute DownButton.VB_MemberFlags = "400"
    DownButton = m_nDownButton
End Property

'== private ==============================================================

Private Property Get pvState(ByVal eState As UcsNineButtonStateEnum) As Boolean
    pvState = (m_eState And eState) <> 0
End Property

Private Property Let pvState(ByVal eState As UcsNineButtonStateEnum, ByVal bValue As Boolean)
    Dim ePrevState      As UcsNineButtonStateEnum
    
    ePrevState = m_eState
    If bValue Then
        If (m_eState And eState) <> eState Then
            m_eState = m_eState Or eState
        End If
    Else
        If (m_eState And eState) <> 0 Then
            m_eState = m_eState And Not eState
        End If
    End If
    If ePrevState <> m_eState And m_bShown Then
        pvStartAnimation m_sngAnimationDuration, _
            m_sngOpacity * m_uButton(pvGetEffectiveState(ePrevState)).ImageOpacity, _
            m_sngOpacity * m_uButton(pvGetEffectiveState(m_eState)).ImageOpacity
    End If
End Property

Private Property Get pvAddressOfTimerProc() As ctxNineButton
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Refresh()
    Const FUNC_NAME     As String = "Refresh"
    Dim hMemDC          As Long
    Dim hPrevDib        As Long
    
    On Error GoTo EH
    If m_hRedrawDib <> 0 Then
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
    If AutoRedraw Then
        hMemDC = CreateCompatibleDC(0)
        If hMemDC = 0 Then
            GoTo QH
        End If
        If Not pvCreateDib(hMemDC, ScaleWidth, ScaleHeight, m_hRedrawDib) Then
            GoTo QH
        End If
        hPrevDib = SelectObject(hMemDC, m_hRedrawDib)
        pvPaintControl hMemDC
    End If
    UserControl.Refresh
QH:
    On Error Resume Next
    If hMemDC <> 0 Then
        Call SelectObject(hMemDC, hPrevDib)
        Call DeleteDC(hMemDC)
        hMemDC = 0
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Sub Repaint()
    Const FUNC_NAME     As String = "Repaint"
    
    On Error GoTo EH
    If m_bShown Then
        pvPrepareBitmap m_eState, m_hFocusBitmap, m_hBitmap
        pvPrepareAttribs m_sngOpacity * m_uButton(pvGetEffectiveState(m_eState)).ImageOpacity, m_hAttributes
        Refresh
'        Call ApiUpdateWindow(ContainerHwnd) '--- pump WM_PAINT
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Public Sub CancelMode()
    Const FUNC_NAME     As String = "CancelMode"
    
    On Error GoTo EH
    pvState(ucsBstHoverPressed) = False
    m_nDownButton = 0
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Public Function TimerProc() As Long
Attribute TimerProc.VB_MemberFlags = "40"
    Const FUNC_NAME     As String = "TimerProc"
    
    On Error GoTo EH
    pvAnimateState TimerEx - m_dblAnimationStart, m_sngAnimationOpacity1, m_sngAnimationOpacity2
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'== private ==============================================================

Private Function pvGetEffectiveState(ByVal eState As UcsNineButtonStateEnum) As UcsNineButtonStateEnum
    If (eState And ucsBstDisabled) <> 0 Then
        If Not m_uButton(ucsBstDisabled).ImagePatch Is Nothing Then
            pvGetEffectiveState = ucsBstDisabled
            Exit Function
        End If
    End If
    eState = eState And Not ucsBstFocused
    If m_uButton(eState).ImagePatch Is Nothing Then
        eState = eState And Not ucsBstHover
    End If
    If m_uButton(eState).ImagePatch Is Nothing Then
        eState = eState And Not ucsBstPressed
    End If
    If m_uButton(eState).ImagePatch Is Nothing Then
        eState = ucsBstNormal
    End If
    pvGetEffectiveState = eState
End Function

Private Function pvPrepareBitmap(ByVal eState As UcsNineButtonStateEnum, hFocusBitmap As Long, hBitmap As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareBitmap"
    Dim hGraphics       As Long
    Dim hNewFocusBitmap As Long
    Dim hNewBitmap      As Long
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim hBrush          As Long
    Dim hShadowBrush    As Long
    Dim lOffset         As Long
    Dim uRect           As RECTF
    Dim hStringFormat   As Long
    Dim hFont           As Long
    Dim sngPicWidth     As Single
    Dim sngPicHeight    As Single
    Dim sCaption        As String
    
    On Error GoTo EH
    If (eState And ucsBstFocused) <> 0 And ((eState And ucsBstHoverPressed) <> ucsBstHoverPressed Or m_bManualFocus) Then
        If hFocusBitmap = 0 Then
            With m_uButton(ucsBstFocused)
                If Not .ImagePatch Is Nothing Then
                    If GdipCreateBitmapFromScan0(ScaleWidth, ScaleHeight, ScaleWidth * 4, PixelFormat32bppARGB, 0, hNewFocusBitmap) <> 0 Then
                        GoTo QH
                    End If
                    If GdipGetImageGraphicsContext(hNewFocusBitmap, hGraphics) <> 0 Then
                        GoTo QH
                    End If
                    If Not .ImagePatch.DrawToGraphics(hGraphics, 0, 0, ScaleWidth, ScaleHeight) Then
                        GoTo QH
                    End If
                    Call GdipDeleteGraphics(hGraphics)
                    hGraphics = 0
                End If
            End With
        Else
            hNewFocusBitmap = hFocusBitmap
        End If
    End If
    With m_uButton(pvGetEffectiveState(eState))
        If Not .TextFont Is Nothing Then
            If Not pvPrepareFont(.TextFont, hFont) Then
                GoTo QH
            End If
        Else
            hFont = m_hFont
        End If
        If GdipCreateBitmapFromScan0(ScaleWidth, ScaleHeight, ScaleWidth * 4, PixelFormat32bppARGB, 0, hNewBitmap) <> 0 Then
            GoTo QH
        End If
        If GdipGetImageGraphicsContext(hNewBitmap, hGraphics) <> 0 Then
            GoTo QH
        End If
        If GdipSetTextRenderingHint(hGraphics, TextRenderingHintClearTypeGridFit) <> 0 Then
            GoTo QH
        End If
        If .ImageZoom <> 1 Then
            If GdipTranslateWorldTransform(hGraphics, -ScaleWidth / 2, -ScaleHeight / 2, MatrixOrderAppend) <> 0 Then
                GoTo QH
            End If
            If GdipScaleWorldTransform(hGraphics, .ImageZoom, .ImageZoom, MatrixOrderAppend) <> 0 Then
                GoTo QH
            End If
            If GdipTranslateWorldTransform(hGraphics, ScaleWidth / 2, ScaleHeight / 2, MatrixOrderAppend) <> 0 Then
                GoTo QH
            End If
        End If
        If Not .ImagePatch Is Nothing Then
            If Not .ImagePatch.DrawToGraphics(hGraphics, 0, 0, ScaleWidth, ScaleHeight) Then
                GoTo QH
            End If
            .ImagePatch.CalcClientRect ScaleWidth, ScaleHeight, lLeft, lTop, lWidth, lHeight
        Else
            lWidth = ScaleWidth
            lHeight = ScaleHeight
        End If
        sCaption = m_sCaption
        RaiseEvent OwnerDraw(hGraphics, hFont, eState, lLeft, lTop, lWidth, lHeight, sCaption, m_hPictureBitmap)
        If lWidth > 0 And lHeight > 0 Then
            If m_hPictureBitmap <> 0 Then
                If GdipGetImageDimension(m_hPictureBitmap, sngPicWidth, sngPicHeight) <> 0 Then
                    GoTo QH
                End If
                If GdipDrawImageRectRect(hGraphics, m_hPictureBitmap, lLeft + (lWidth - sngPicWidth) / 2, lTop + (lHeight - sngPicHeight) / 2, sngPicWidth, sngPicHeight, 0, 0, sngPicWidth, sngPicHeight, , m_hPictureAttributes) <> 0 Then
                    GoTo QH
                End If
            ElseIf hFont <> 0 And LenB(sCaption) <> 0 Then
                If GdipCreateSolidFill(pvTranslateColor(IIf(.TextColor = DEF_TEXTCOLOR, m_clrFore, .TextColor), .TextOpacity), hBrush) <> 0 Then
                    GoTo QH
                End If
                If Not pvPrepareStringFormat(.TextFlags, hStringFormat) Then
                    GoTo QH
                End If
                lOffset = .TextOffsetX * -((eState And ucsBstHoverPressed) = ucsBstHoverPressed)
                uRect.Left = lLeft + lOffset
                lOffset = .TextOffsetY * -((eState And ucsBstHoverPressed) = ucsBstHoverPressed)
                uRect.Top = lTop + lOffset
                uRect.Right = lWidth
                uRect.Bottom = lHeight
                If .ShadowOffsetX <> 0 Or .ShadowOffsetY <> 0 Or .ImagePatch Is Nothing Then
                    If GdipCreateSolidFill(pvTranslateColor(.ShadowColor, .ShadowOpacity), hShadowBrush) <> 0 Then
                        GoTo QH
                    End If
                    If GdipSetTextRenderingHint(hGraphics, TextRenderingHintAntiAlias) <> 0 Then
                        GoTo QH
                    End If
                    uRect.Left = uRect.Left + .ShadowOffsetX
                    uRect.Top = uRect.Top + .ShadowOffsetY
                    If GdipDrawString(hGraphics, StrPtr(sCaption), -1, hFont, uRect, hStringFormat, hShadowBrush) <> 0 Then
                        GoTo QH
                    End If
                    uRect.Left = uRect.Left - .ShadowOffsetX
                    uRect.Top = uRect.Top - .ShadowOffsetY
                End If
                If GdipDrawString(hGraphics, StrPtr(sCaption), -1, hFont, uRect, hStringFormat, hBrush) <> 0 Then
                    GoTo QH
                End If
            End If
        End If
    End With
    '--- commit
    If hNewFocusBitmap <> hFocusBitmap Then
        If hFocusBitmap <> 0 Then
            Call GdipDisposeImage(hFocusBitmap)
            hFocusBitmap = 0
        End If
        hFocusBitmap = hNewFocusBitmap
    End If
    hNewFocusBitmap = 0
    If hNewBitmap <> hBitmap Then
        If hBitmap <> 0 Then
            Call GdipDisposeImage(hBitmap)
            hBitmap = 0
        End If
        hBitmap = hNewBitmap
    End If
    hNewBitmap = 0
    '-- success
    pvPrepareBitmap = True
QH:
    On Error Resume Next
    If hFont <> 0 And hFont <> m_hFont Then
        Call GdipDeleteFont(hFont)
        hFont = 0
    End If
    If hStringFormat <> 0 Then
        Call GdipDeleteStringFormat(hStringFormat)
        hStringFormat = 0
    End If
    If hShadowBrush <> 0 Then
        Call GdipDeleteBrush(hShadowBrush)
        hShadowBrush = 0
    End If
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
        hBrush = 0
    End If
    If hNewFocusBitmap <> 0 Then
        Call GdipDisposeImage(hNewFocusBitmap)
        hNewFocusBitmap = 0
    End If
    If hNewBitmap <> 0 Then
        Call GdipDisposeImage(hNewBitmap)
        hNewBitmap = 0
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPrepareAttribs(ByVal sngAlpha As Single, hAttributes As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareAttribs"
    Dim clrMatrix(0 To 4, 0 To 4) As Single
    Dim hNewAttributes  As Long
    
    On Error GoTo EH
    If GdipCreateImageAttributes(hNewAttributes) <> 0 Then
        GoTo QH
    End If
    clrMatrix(0, 0) = 1
    clrMatrix(1, 1) = 1
    clrMatrix(2, 2) = 1
    clrMatrix(3, 3) = sngAlpha
    clrMatrix(4, 4) = 1
    If GdipSetImageAttributesColorMatrix(hNewAttributes, 0, 1, clrMatrix(0, 0), clrMatrix(0, 0), 0) <> 0 Then
        GoTo QH
    End If
    '--- commit
    If hAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hAttributes)
        hAttributes = 0
    End If
    hAttributes = hNewAttributes
    hNewAttributes = 0
    '--- success
    pvPrepareAttribs = True
QH:
    On Error Resume Next
    If hNewAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hNewAttributes)
        hNewAttributes = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPrepareFont(oFont As StdFont, hFont As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareFont"
    Dim hFamily         As Long
    Dim hNewFont        As Long
    Dim eStyle          As FontStyle

    On Error GoTo EH
    If oFont Is Nothing Then
        GoTo QH
    End If
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFamily) <> 0 Then
        If GdipGetGenericFontFamilySansSerif(hFamily) <> 0 Then
            GoTo QH
        End If
    End If
    eStyle = FontStyleBold * -oFont.Bold _
        Or FontStyleItalic * -oFont.Italic _
        Or FontStyleUnderline * -oFont.Underline _
        Or FontStyleStrikeout * -oFont.Strikethrough
    If GdipCreateFont(hFamily, oFont.Size, eStyle, UnitPoint, hNewFont) <> 0 Then
        GoTo QH
    End If
    '--- commit
    If hFont <> 0 Then
        Call GdipDeleteFont(hFont)
    End If
    hFont = hNewFont
    hNewFont = 0
    '--- success
    pvPrepareFont = True
QH:
    On Error Resume Next
    If hFamily <> 0 Then
        Call GdipDeleteFontFamily(hFamily)
        hFamily = 0
    End If
    If hNewFont <> 0 Then
        Call GdipDeleteFont(hNewFont)
        hNewFont = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPrepareStringFormat(ByVal lFlags As UcsNineButtonTextFlagsEnum, hStringFormat As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareStringFormat"
    Dim hNewFormat      As Long
    
    On Error GoTo EH
    If GdipCreateStringFormat(0, 0, hNewFormat) <> 0 Then
        GoTo QH
    End If
    If GdipSetStringFormatAlign(hNewFormat, lFlags And 3) <> 0 Then
        GoTo QH
    End If
    If GdipSetStringFormatLineAlign(hNewFormat, (lFlags \ 4) And 3) <> 0 Then
        GoTo QH
    End If
    If GdipSetStringFormatFlags(hNewFormat, lFlags \ 16) <> 0 Then
        GoTo QH
    End If
    '--- commit
    If hStringFormat <> 0 Then
        Call GdipDeleteStringFormat(hStringFormat)
    End If
    hStringFormat = hNewFormat
    hNewFormat = 0
    '--- success
    pvPrepareStringFormat = True
QH:
    On Error Resume Next
    If hNewFormat <> 0 Then
        Call GdipDeleteStringFormat(hNewFormat)
        hNewFormat = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvPreparePicture(oPicture As StdPicture, ByVal clrMask As OLE_COLOR, hPictureBitmap As Long, hPictureAttributes As Long) As Boolean
    Const FUNC_NAME     As String = "pvPreparePicture"
    Dim hTempBitmap     As Long
    Dim hNewBitmap      As Long
    Dim hNewAttributes  As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim uHdr            As BITMAPINFOHEADER
    Dim hMemDC          As Long
    Dim uInfo           As ICONINFO
    Dim baColorBits()   As Byte
    Dim bHasAlpha       As Boolean
    Dim hDib            As Long
    Dim lpBits          As Long
    Dim hPrevDib        As Long
    Dim lIdx            As Long
    
    On Error GoTo EH
    If Not oPicture Is Nothing Then
        If oPicture.Handle <> 0 Then
            Select Case oPicture.Type
            Case vbPicTypeBitmap
                If GdipCreateBitmapFromHBITMAP(oPicture.Handle, 0, hNewBitmap) <> 0 Then
                    GoTo QH
                End If
                If clrMask <> -1 Then
                    If GdipCreateImageAttributes(hNewAttributes) <> 0 Then
                        GoTo QH
                    End If
                    If GdipSetImageAttributesColorKeys(hNewAttributes, 0, 1, pvTranslateColor(clrMask), pvTranslateColor(clrMask)) <> 0 Then
                        GoTo QH
                    End If
                End If
            Case Else
                lWidth = HM2Pix(oPicture.Width)
                lHeight = HM2Pix(oPicture.Height)
                hMemDC = CreateCompatibleDC(0)
                If hMemDC = 0 Then
                    GoTo QH
                End If
                With uHdr
                    .biSize = Len(uHdr)
                    .biPlanes = 1
                    .biBitCount = 32
                    .biWidth = lWidth
                    .biHeight = -lHeight
                    .biSizeImage = (4 * lWidth) * lHeight
                End With
                If oPicture.Type = vbPicTypeIcon Then
                    If GetIconInfo(oPicture.Handle, uInfo) = 0 Then
                        GoTo QH
                    End If
                    ReDim baColorBits(0 To uHdr.biSizeImage - 1) As Byte
                    If GetDIBits(hMemDC, uInfo.hbmColor, 0, lHeight, baColorBits(0), uHdr, DIB_RGB_COLORS) = 0 Then
                        GoTo QH
                    End If
                    For lIdx = 3 To UBound(baColorBits) Step 4
                        If baColorBits(lIdx) <> 0 Then
                            bHasAlpha = True
                            Exit For
                        End If
                    Next
                    If Not bHasAlpha Then
                        '--- note: GdipCreateBitmapFromHICON working ok for old-style (single-bit) transparent icons only
                        If GdipCreateBitmapFromHICON(oPicture.Handle, hNewBitmap) <> 0 Then
                            GoTo QH
                        End If
                    Else
                        If GdipCreateBitmapFromScan0(lWidth, lHeight, 4 * lWidth, PixelFormat32bppARGB, VarPtr(baColorBits(0)), hTempBitmap) <> 0 Then
                            GoTo QH
                        End If
                        '--- note: pixel format (or size) *must* differ from hTempBitmap's one for actual
                        '---   memcpy to happen (PixelFormat32bppARGB -> PixelFormat32bppPARGB)
                        If GdipCloneBitmapAreaI(0, 0, lWidth, lHeight, PixelFormat32bppPARGB, hTempBitmap, hNewBitmap) <> 0 Then
                            GoTo QH
                        End If
                    End If
                Else
                    hDib = CreateDIBSection(hMemDC, uHdr, DIB_RGB_COLORS, lpBits, 0, 0)
                    If hDib = 0 Then
                        GoTo QH
                    End If
                    hPrevDib = SelectObject(hMemDC, hDib)
                    pvRenderPicture oPicture, hMemDC, 0, 0, lWidth, lHeight, 0, oPicture.Height, oPicture.Width, -oPicture.Height
                    If GdipCreateBitmapFromScan0(lWidth, lHeight, 4 * lWidth, PixelFormat32bppARGB, lpBits, hTempBitmap) <> 0 Then
                        GoTo QH
                    End If
                    If GdipCloneBitmapAreaI(0, 0, lWidth, lHeight, PixelFormat32bppPARGB, hTempBitmap, hNewBitmap) <> 0 Then
                        GoTo QH
                    End If
                End If
            End Select
        End If
    End If
    '--- commit
    If hPictureBitmap <> 0 Then
        Call GdipDisposeImage(hPictureBitmap)
        hPictureBitmap = 0
    End If
    hPictureBitmap = hNewBitmap
    hNewBitmap = 0
    If hPictureAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hPictureAttributes)
        hPictureAttributes = 0
    End If
    hPictureAttributes = hNewAttributes
    hNewAttributes = 0
    '--- success
    pvPreparePicture = True
QH:
    On Error Resume Next
    If hNewBitmap <> 0 Then
        Call GdipDisposeImage(hNewBitmap)
        hNewBitmap = 0
    End If
    If hNewAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hNewAttributes)
        hNewAttributes = 0
    End If
    If hTempBitmap <> 0 Then
        Call GdipDisposeImage(hTempBitmap)
        hTempBitmap = 0
    End If
    If hPrevDib <> 0 Then
        Call SelectObject(hMemDC, hPrevDib)
        hPrevDib = 0
    End If
    If hDib <> 0 Then
        Call DeleteObject(hDib)
        hDib = 0
    End If
    If uInfo.hbmColor <> 0 Then
        Call DeleteObject(uInfo.hbmColor)
        uInfo.hbmColor = 0
    End If
    If uInfo.hbmMask <> 0 Then
        Call DeleteObject(uInfo.hbmMask)
        uInfo.hbmMask = 0
    End If
    If hMemDC <> 0 Then
        Call DeleteDC(hMemDC)
        hMemDC = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvTranslateColor(ByVal clrValue As OLE_COLOR, Optional ByVal Alpha As Single = 1) As Long
    Dim uQuad           As UcsRgbQuad
    Dim lTemp           As Long
    
    Call OleTranslateColor(clrValue, 0, VarPtr(uQuad))
    lTemp = uQuad.R
    uQuad.R = uQuad.B
    uQuad.B = lTemp
    lTemp = Alpha * &HFF
    If lTemp > 255 Then
        uQuad.A = 255
    ElseIf lTemp < 0 Then
        uQuad.A = 0
    Else
        uQuad.A = lTemp
    End If
    Call CopyMemory(pvTranslateColor, uQuad, 4)
End Function

Private Function pvParentRegisterCancelMode(oCtl As Object) As Boolean
    Dim bHandled        As Boolean
    
    RaiseEvent RegisterCancelMode(oCtl, bHandled)
    If Not bHandled Then
        On Error GoTo QH
        Parent.RegisterCancelMode oCtl
        On Error GoTo 0
    End If
    '--- success
    pvParentRegisterCancelMode = True
QH:
End Function

Private Function pvHitTest(ByVal X As Single, ByVal Y As Single) As HitResultConstants
    Const FUNC_NAME     As String = "pvHitTest"
    Dim uQuad           As UcsRgbQuad
    
    On Error GoTo EH
    pvHitTest = vbHitResultHit
    If GdipBitmapGetPixel(m_hBitmap, X, Y, uQuad) <> 0 Then
        GoTo QH
    End If
    If uQuad.A < 255 Then
        If uQuad.A > 0 Then
            pvHitTest = vbHitResultTransparent
        Else
            pvHitTest = vbHitResultOutside
        End If
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvStartAnimation(ByVal sngDuration As Single, ByVal sngOpacity1 As Single, ByVal sngOpacity2 As Single) As Boolean
    Const FUNC_NAME     As String = "pvStartAnimation"
    Dim hNewBitmap      As Long
    
    On Error GoTo EH
    If Not pvPrepareBitmap(m_eState, m_hFocusBitmap, hNewBitmap) Then
        GoTo QH
    End If
    If m_hPrevBitmap <> 0 Then
        Call GdipDisposeImage(m_hPrevBitmap)
        m_hPrevBitmap = 0
    End If
    m_hPrevBitmap = m_hBitmap
    m_hBitmap = hNewBitmap
    hNewBitmap = 0
    m_dblAnimationStart = TimerEx
    m_dblAnimationEnd = m_dblAnimationStart + sngDuration
    m_sngAnimationOpacity1 = sngOpacity1
    m_sngAnimationOpacity2 = sngOpacity2
    pvAnimateState 0, m_sngAnimationOpacity1, m_sngAnimationOpacity2
QH:
    If hNewBitmap <> 0 Then
        Call GdipDisposeImage(hNewBitmap)
        hNewBitmap = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvAnimateState(ByVal dblElapsed As Double, ByVal sngOpacity1 As Single, ByVal sngOpacity2 As Single) As Boolean
    Const FUNC_NAME     As String = "pvAnimateState"
    Dim sngOpacity      As Single
    Dim dblFull         As Double

    On Error GoTo EH
    sngOpacity = sngOpacity2
    m_sngBitmapAlpha = 1
    dblFull = (m_dblAnimationEnd - m_dblAnimationStart)
    If dblFull > DBL_EPLISON And dblElapsed <= dblFull Then
        sngOpacity = sngOpacity1 + (sngOpacity2 - sngOpacity1) * dblElapsed / dblFull
        m_sngBitmapAlpha = dblElapsed / dblFull
    End If
    If Not pvPrepareAttribs(sngOpacity, m_hAttributes) Then
        GoTo QH
    End If
    Refresh
    If m_sngBitmapAlpha < 1 Then
        Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc)
    End If
    '--- success
    pvAnimateState = True
    Exit Function
QH:
    On Error Resume Next
    If m_hPrevBitmap <> 0 Then
        Call GdipDisposeImage(m_hPrevBitmap)
        m_hPrevBitmap = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Property Get pvNppGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_NPP_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvNppGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvNppGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_NPP_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property

Private Sub pvSetStyle(ByVal eStyle As UcsNineButtonStyleEnum)
    Const FUNC_NAME     As String = "pvSetStyle"
    Static hResBitmap   As Long
    
    On Error GoTo EH
    pvSetEmptyStyle
    If eStyle <> ucsBtyNone Then
        If hResBitmap = 0 Then
            hResBitmap = pvNppGlobalData("hResBitmap")
        End If
        If hResBitmap = 0 Then
            With New cNinePatch
                If Not .frBitmapFromByteArray(FromBase64Array(STR_RES_PNG1 & STR_RES_PNG2), hResBitmap) Then
                    GoTo QH
                End If
            End With
            pvNppGlobalData("hResBitmap") = hResBitmap
        End If
        Select Case eStyle
        '--- buttons
        Case ucsBtyButtonDefault
            pvSetButtonStyle hResBitmap, ucsIdxButtonDefNormal, vbBlack, _
                ShadowOpacity:=0.8, ShadowColor:=vbWhite, ShadowOffsetY:=1
        Case ucsBtyButtonGreen
            pvSetButtonStyle hResBitmap, ucsIdxButtonGreenNormal, vbWhite, _
                ShadowOpacity:=0.2
        Case ucsBtyButtonRed
            pvSetButtonStyle hResBitmap, ucsIdxButtonRedNormal, vbWhite, _
                ShadowOpacity:=0.2
        Case ucsBtyButtonTurnGreen
            pvSetButtonStyle hResBitmap, ucsIdxButtonGreenNormal, vbWhite, _
                ShadowOpacity:=0.2, NormalTextColor:=&H7C3F&
        Case ucsBtyButtonTurnRed
            pvSetButtonStyle hResBitmap, ucsIdxButtonRedNormal, vbWhite, _
                ShadowOpacity:=0.2, NormalTextColor:=&H3124CB
        '--- flat buttons
        Case ucsBtyFlatPrimary
            pvSetFlatStyle hResBitmap, ucsIdxFlatPrimaryNormal, vbWhite, _
                ShadowOpacity:=0.2
        Case ucsBtyFlatSecondary
            pvSetFlatStyle hResBitmap, ucsIdxFlatSecondaryNormal, &H575049, _
                ShadowOpacity:=0.8, ShadowColor:=vbWhite, ShadowOffsetY:=1, PressedOffset:=1
        Case ucsBtyFlatSuccess
            pvSetFlatStyle hResBitmap, ucsIdxFlatSuccessNormal, vbWhite, _
                ShadowOpacity:=0.2
        Case ucsBtyFlatDanger
            pvSetFlatStyle hResBitmap, ucsIdxFlatDangerNormal, vbWhite, _
                ShadowOpacity:=0.2
        Case ucsBtyFlatWarning
            pvSetFlatStyle hResBitmap, ucsIdxFlatWarningNormal, vbWhite, _
                ShadowOpacity:=0.2
        Case ucsBtyFlatInfo
            pvSetFlatStyle hResBitmap, ucsIdxFlatInfoNormal, vbWhite, _
                ShadowOpacity:=0.2
        Case ucsBtyFlatLight
            pvSetFlatStyle hResBitmap, ucsIdxFlatLightNormal, &H575049, _
                ShadowOpacity:=0.8, ShadowColor:=vbWhite, ShadowOffsetY:=1
        Case ucsBtyFlatDark
            pvSetFlatStyle hResBitmap, ucsIdxFlatDarkNormal, vbWhite, _
                ShadowOpacity:=0.2
        '--- outline buttons
        Case ucsBtyOutlinePrimary
            pvSetOutlineStyle hResBitmap, ucsIdxFlatPrimaryOutline, &HCF7F46, _
                ShadowOpacity:=0.2
        Case ucsBtyOutlineSecondary
            pvSetOutlineStyle hResBitmap, ucsIdxFlatSecondaryOutline, &H575049, _
                ShadowOpacity:=0.2, HoverOffset:=2
        Case ucsBtyOutlineSuccess
            pvSetOutlineStyle hResBitmap, ucsIdxFlatSuccessOutline, &HBA5E&, _
                ShadowOpacity:=0.2
        Case ucsBtyOutlineDanger
            pvSetOutlineStyle hResBitmap, ucsIdxFlatDangerOutline, &H1F20CD, _
                ShadowOpacity:=0.2
        Case ucsBtyOutlineWarning
            pvSetOutlineStyle hResBitmap, ucsIdxFlatWarningOutline, &HFC4F1, _
                ShadowOpacity:=0.2
        Case ucsBtyOutlineInfo
            pvSetOutlineStyle hResBitmap, ucsIdxFlatInfoOutline, &HF2AA45, _
                ShadowOpacity:=0.2
        Case ucsBtyOutlineLight
            pvSetOutlineStyle hResBitmap, ucsIdxFlatLightOutline, &H575049, _
                ShadowOpacity:=0.8, ShadowColor:=vbWhite, ShadowOffsetY:=1, TextColor:=DEF_TEXTCOLOR, HoverOffset:=1
        Case ucsBtyOutlineDark
            pvSetOutlineStyle hResBitmap, ucsIdxFlatDarkOutline, &H403A34, _
                ShadowOpacity:=0.2
        '--- cards
        Case ucsBtyCardDefault
            pvSetCardStyle hResBitmap, ucsIdxCardDefault
        Case ucsBtyCardPrimary
            pvSetCardStyle hResBitmap, ucsIdxCardPrimary
        Case ucsBtyCardSuccess
            pvSetCardStyle hResBitmap, ucsIdxCardSuccess
        Case ucsBtyCardOrange
            pvSetCardStyle hResBitmap, ucsIdxCardOrange
        Case ucsBtyCardDanger
            pvSetCardStyle hResBitmap, ucsIdxCardDanger
        Case ucsBtyCardWarning
            pvSetCardStyle hResBitmap, ucsIdxCardWarning
        Case ucsBtyCardPurple
            pvSetCardStyle hResBitmap, ucsIdxCardPurple
        End Select
    End If
QH:
'    If hResBitmap <> 0 Then
'        Call GdipDisposeImage(hResBitmap)
'        hResBitmap = 0
'    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub pvSetButtonStyle( _
            ByVal hResBitmap As Long, _
            ByVal eIdx As UcsNineButtonResIndex, _
            ByVal clrFore As OLE_COLOR, _
            Optional ByVal ShadowOpacity As Single = 1, _
            Optional ByVal ShadowColor As OLE_COLOR = vbBlack, _
            Optional ByVal ShadowOffsetY As Long = -1, _
            Optional ByVal NormalTextColor As OLE_COLOR = DEF_TEXTCOLOR)
    With m_uButton(ucsBstNormal)
        If NormalTextColor = DEF_TEXTCOLOR Then
            Set .ImagePatch = pvResExtract(hResBitmap, eIdx, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
            .ShadowOpacity = ShadowOpacity
            .ShadowColor = ShadowColor
            .ShadowOffsetY = ShadowOffsetY
        Else
            Set .ImagePatch = pvResExtract(hResBitmap, ucsIdxButtonDefNormal, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
            .TextColor = NormalTextColor
            .ShadowOpacity = 0.8
            .ShadowColor = vbWhite
            .ShadowOffsetY = 1
        End If
    End With
    With m_uButton(ucsBstHover)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 1, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstPressed)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 2, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
        .TextOffsetY = 1
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstDisabled)
        Set .ImagePatch = pvResExtract(hResBitmap, ucsIdxButtonDisabled, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
        .TextOpacity = 0.4
        .TextColor = &H2E2924
        .ShadowOpacity = 0.8
        .ShadowColor = vbWhite
        .ShadowOffsetY = 1
    End With
    With m_uButton(ucsBstFocused)
        Set .ImagePatch = pvResExtract(hResBitmap, ucsIdxButtonFocus, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
    End With
    AnimationDuration = 0.2
    ForeColor = clrFore
End Sub

Private Sub pvSetFlatStyle( _
            ByVal hResBitmap As Long, _
            ByVal eIdx As UcsNineButtonResIndex, _
            ByVal clrFore As OLE_COLOR, _
            Optional ByVal ShadowOpacity As Single = 1, _
            Optional ByVal ShadowColor As OLE_COLOR = vbBlack, _
            Optional ByVal ShadowOffsetY As Long = -1, _
            Optional ByVal PressedOffset As Long = 0)
    With m_uButton(ucsBstNormal)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstHover)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 2, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstPressed)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 3 + PressedOffset, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .TextOffsetY = 1
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstDisabled)
        Set .ImagePatch = m_uButton(ucsBstNormal).ImagePatch
        .ImageOpacity = 0.65
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstFocused)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 4 + PressedOffset, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
    End With
    AnimationDuration = 0.2
    ForeColor = clrFore
End Sub

Private Sub pvSetOutlineStyle( _
            ByVal hResBitmap As Long, _
            ByVal eIdx As UcsNineButtonResIndex, _
            ByVal clrFore As OLE_COLOR, _
            Optional ByVal ShadowOpacity As Single = DEF_SHADOWOPACITY, _
            Optional ByVal ShadowColor As OLE_COLOR = vbBlack, _
            Optional ByVal ShadowOffsetY As Long = -1, _
            Optional ByVal TextColor As OLE_COLOR = vbWhite, _
            Optional ByVal HoverOffset As Long = -1)
    With m_uButton(ucsBstNormal)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .ShadowOpacity = 0
        .ShadowOffsetY = 1
    End With
    With m_uButton(ucsBstHover)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + HoverOffset, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .TextColor = TextColor
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstPressed)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + HoverOffset, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .TextColor = TextColor
        .TextOffsetY = 1
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstDisabled)
        Set .ImagePatch = m_uButton(ucsBstNormal).ImagePatch
        .ImageOpacity = 0.65
        .ShadowOpacity = 0
        .ShadowOffsetY = 1
    End With
    With m_uButton(ucsBstFocused)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 3, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
    End With
    AnimationDuration = 0.2
    ForeColor = clrFore
End Sub

Private Sub pvSetCardStyle( _
            ByVal hResBitmap As Long, _
            ByVal eIdx As UcsNineButtonResIndex)
    With m_uButton(ucsBstNormal)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx, LNG_CARD_TOP, LNG_CARD_WIDTH, LNG_CARD_HEIGHT)
    End With
    With m_uButton(ucsBstDisabled)
        Set .ImagePatch = m_uButton(ucsBstNormal).ImagePatch
        .ImageOpacity = 0.65
    End With
    With m_uButton(ucsBstFocused)
        Set .ImagePatch = pvResExtract(hResBitmap, ucsIdxCardFocus, LNG_CARD_TOP, LNG_CARD_WIDTH, LNG_CARD_HEIGHT)
    End With
    ForeColor = vbBlack
End Sub

Private Function pvResExtract( _
            ByVal hResBitmap As Long, _
            ByVal eIdx As UcsNineButtonResIndex, _
            ByVal lTop As Long, _
            ByVal lWidth As Long, _
            ByVal lHeight As Long) As cNinePatch
    Const FUNC_NAME     As String = "pvResExtract"
    Dim hNewBitmap      As Long
    Dim lLeft           As Long
    Dim oRetVal         As cNinePatch
    
    On Error GoTo EH
    lLeft = eIdx * lWidth
    lTop = lTop + lHeight * (lLeft \ LNG_RES_WIDTH)
    lLeft = lLeft Mod LNG_RES_WIDTH
    If GdipCloneBitmapAreaI(lLeft, lTop, lWidth, lHeight, 0, hResBitmap, hNewBitmap) <> 0 Then
        GoTo QH
    End If
    Set oRetVal = New cNinePatch
    If Not oRetVal.LoadFromBitmap(hNewBitmap) Then
        GoTo QH
    End If
    hNewBitmap = 0
    '--- success
    Set pvResExtract = oRetVal
QH:
    On Error Resume Next
    If hNewBitmap <> 0 Then
        Call GdipDisposeImage(hNewBitmap)
        hNewBitmap = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Sub pvSetEmptyStyle()
    Dim lIdx            As Long
    Dim uEmpty          As UcsNineButtonStateType

    With uEmpty
        .ImageOpacity = DEF_IMAGEOPACITY
        .ImageZoom = DEF_IMAGEZOOM
        .TextOpacity = DEF_TEXTOPACITY
        .TextColor = DEF_TEXTCOLOR
        .TextFlags = DEF_TEXTFLAGS
        .ShadowOpacity = DEF_SHADOWOPACITY
        .ShadowColor = DEF_SHADOWCOLOR
    End With
    For lIdx = 0 To UBound(m_uButton)
        m_uButton(lIdx) = uEmpty
    Next
End Sub

Private Sub pvHandleMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "pvHandleMouseDown"
    
    On Error GoTo EH
    m_nDownButton = Button
    m_nDownShift = Shift
    m_sngDownX = X
    m_sngDownY = Y
    If (Button And vbLeftButton) <> 0 Then
        If pvHitTest(X, Y) <> vbHitResultOutside Then
            pvParentRegisterCancelMode Me
            pvState(ucsBstPressed Or ucsBstFocused * (1 + m_bManualFocus)) = True
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvHandleClick()
    Const FUNC_NAME     As String = "pvHandleClick"
    
    On Error GoTo EH
    pvState(ucsBstPressed) = True
    pvState(ucsBstPressed) = False
    Call ApiUpdateWindow(ContainerHwnd) '--- pump WM_PAINT
    RaiseEvent Click
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvRenderPicture(pPicture As IPicture, ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal xSrc As OLE_XPOS_HIMETRIC, ByVal ySrc As OLE_YPOS_HIMETRIC, ByVal cxSrc As OLE_XSIZE_HIMETRIC, ByVal cySrc As OLE_YSIZE_HIMETRIC)
    If Not pPicture Is Nothing Then
        If pPicture.Handle <> 0 Then
            pPicture.Render hDC, X, Y, cx, cy, xSrc, ySrc, cxSrc, cySrc, ByVal 0
        End If
    End If
End Sub

Private Function pvMergeBitmap(ByVal hDstBitmap As Long, ByVal hSrcBitmap As Long, ByVal lDstAlpha As Long, ByVal lSrcAlpha As Long) As Boolean
    Const FUNC_NAME     As String = "pvMergeBitmap"
    Dim uDstData        As BitmapData
    Dim uSrcData        As BitmapData
    Dim uDstArray       As SAFEARRAY1D
    Dim uSrcArray       As SAFEARRAY1D
    Dim baDst()         As Byte
    Dim baSrc()         As Byte
    Dim lIdx            As Long
    Dim lY              As Long
    Dim lX              As Long
    Dim lG              As Long

    On Error GoTo EH
    If GdipBitmapLockBits(hDstBitmap, ByVal 0, ImageLockModeRead Or ImageLockModeWrite, PixelFormat32bppARGB, uDstData) <> 0 Then
        GoTo QH
    End If
    If GdipBitmapLockBits(hSrcBitmap, ByVal 0, ImageLockModeRead, PixelFormat32bppARGB, uSrcData) <> 0 Then
        GoTo QH
    End If
    With uDstArray
        .cDims = 1
        .fFeatures = 1 ' FADF_AUTO
        .cbElements = 1
        .cLocks = 1
        .pvData = uDstData.Scan0
        .cElements = uDstData.Stride * uDstData.Height
    End With
    Call CopyMemory(ByVal ArrPtr(baDst), VarPtr(uDstArray), 4)
    With uSrcArray
        .cDims = 1
        .fFeatures = 1 ' FADF_AUTO
        .cbElements = 1
        .cLocks = 1
        .pvData = uSrcData.Scan0
        .cElements = uSrcData.Stride * uSrcData.Height
    End With
    Call CopyMemory(ByVal ArrPtr(baSrc), VarPtr(uSrcArray), 4)
    For lIdx = 0 To UBound(baDst)
        lY = lIdx \ uDstData.Stride
        If lY < uSrcData.Height Then
            lX = lIdx - (lY * uDstData.Stride)
            If lX < uSrcData.Stride Then
                lG = (baDst(lIdx) * lDstAlpha + baSrc(lY * uSrcData.Stride + lX) * lSrcAlpha) \ 255
                If lG > 255 Then
                    lG = 255
                ElseIf lG < 0 Then
                    lG = 0
                End If
                baDst(lIdx) = lG
            End If
        End If
    Next
    '--- success
    pvMergeBitmap = True
QH:
    On Error Resume Next
    If uDstData.Scan0 <> 0 Then
        Call GdipBitmapUnlockBits(hDstBitmap, uDstData)
    End If
    If uSrcData.Scan0 <> 0 Then
        Call GdipBitmapUnlockBits(hSrcBitmap, uSrcData)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Sub pvRefresh()
    m_bShown = False
    If m_hRedrawDib <> 0 Then
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
    UserControl.Refresh
End Sub

Private Function pvPaintControl(ByVal hDC As Long) As Boolean
    Const FUNC_NAME     As String = "pvPaintControl"
    Dim hGraphics       As Long
    Dim hMergeBitmap    As Long
    Dim lSrcAlpha       As Long
    
    On Error GoTo EH
    If Not m_bShown Then
        m_bShown = True
        pvPrepareBitmap m_eState, m_hFocusBitmap, m_hBitmap
        pvPrepareAttribs m_sngOpacity * m_uButton(pvGetEffectiveState(m_eState)).ImageOpacity, m_hAttributes
    End If
    If m_hBitmap <> 0 Then
        If m_hPrevBitmap <> 0 Then
            lSrcAlpha = Int(m_sngBitmapAlpha * 255 + 0.5!)
            If lSrcAlpha < 255 Then
                If GdipCloneImage(m_hPrevBitmap, hMergeBitmap) <> 0 Then
                    GoTo QH
                End If
                If lSrcAlpha > 0 Then
                    If Not pvMergeBitmap(hMergeBitmap, m_hBitmap, 255 - lSrcAlpha, lSrcAlpha) Then
                        GoTo QH
                    End If
                End If
            End If
        End If
        If GdipCreateFromHDC(hDC, hGraphics) <> 0 Then
            GoTo QH
        End If
        If m_hFocusBitmap <> 0 Then
            If GdipDrawImageRectRect(hGraphics, m_hFocusBitmap, 0, 0, ScaleWidth, ScaleHeight, 0, 0, ScaleWidth, ScaleHeight, , m_hFocusAttributes) <> 0 Then
                GoTo QH
            End If
        End If
        If GdipDrawImageRectRect(hGraphics, IIf(hMergeBitmap <> 0, hMergeBitmap, m_hBitmap), 0, 0, ScaleWidth, ScaleHeight, 0, 0, ScaleWidth, ScaleHeight, , m_hAttributes) <> 0 Then
            GoTo QH
        End If
        '--- success
        pvPaintControl = True
    End If
QH:
    On Error Resume Next
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    If hMergeBitmap <> 0 Then
        Call GdipDisposeImage(hMergeBitmap)
        hMergeBitmap = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvCreateDib(ByVal hMemDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, hDib As Long) As Boolean
    Const FUNC_NAME     As String = "pvCreateDib"
    Dim uHdr            As BITMAPINFOHEADER
    Dim lpBits          As Long
    
    On Error GoTo EH
    With uHdr
        .biSize = Len(uHdr)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = lWidth
        .biHeight = -lHeight
        .biSizeImage = 4 * lWidth * lHeight
    End With
    hDib = CreateDIBSection(hMemDC, uHdr, DIB_RGB_COLORS, lpBits, 0, 0)
    If hDib = 0 Then
        GoTo QH
    End If
    '--- success
    pvCreateDib = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

#If Not ImplUseShared Then
Private Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Private Function FromBase64Array(sText As String) As Byte()
    Dim lSize           As Long
    Dim dwDummy         As Long
    Dim baOutput()      As Byte
    
    lSize = Len(sText) + 1
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(sText, Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, dwDummy)
    If lSize > 0 Then
        ReDim Preserve baOutput(0 To lSize - 1) As Byte
        FromBase64Array = baOutput
    Else
        FromBase64Array = vbNullString
    End If
End Function

Private Function HM2Pix(ByVal Value As Single) As Long
   HM2Pix = Int(Value * 1440 / 2540 / Screen.TwipsPerPixelX + 0.5!)
End Function

Private Function ToScaleMode(sScaleUnits As String) As ScaleModeConstants
    Select Case sScaleUnits
    Case "Twip"
        ToScaleMode = vbTwips
    Case "Point"
        ToScaleMode = vbPoints
    Case "Pixel"
        ToScaleMode = vbPixels
    Case "Character"
        ToScaleMode = vbCharacters
    Case "Centimeter"
        ToScaleMode = vbCentimeters
    Case "Millimeter"
        ToScaleMode = vbMillimeters
    Case "Inch"
        ToScaleMode = vbInches
    Case Else
        ToScaleMode = vbTwips
    End Select
End Function

Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If hThunk = 0 Then
        Exit Function
    End If
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitFireOnceTimerThunk(pObj As Object, ByVal pfnCallback As Long, Optional Delay As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgeogERkAV1aLdCQUg8YIgz4AdCqL+oHHBBMZAIvCBSgSGQCri8IFZBIZAKuLwgV0EhkAqzPAq7kIAAAA86WBwgQTGQBSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0IMSCX/AAAAUItKDDsMJHULWIsPV/9RFDP/62P/QgyBYgz/AAAAjQTKjQTIjUyIMIB5EwB101jHAf80JLiJeQTHQQiJRCQEi8ItBBMZAAWgEhkAUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQiU8MUf90JBRqAGoAiw//URiJRwiLRCQYiTheX7g0ExkALSARGQAFABQAAMIQAGaQi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1HYtCDMZAEwCLCv9yCGoA/1Eci1QkBIsKUv9RFDPAwgQAi1QkBIsKi0EohcB0J1L/0FqD+AF3SYsKUv9RLFqFwHU+iwpSavD/cSD/USRaqQAAAAh1K4sKUv9yCGoA/1EcWv9CBDPAUFT/chD/UhSLVCQIx0IIAAAAAFLodv///1jCFABmkA==" ' 27.3.2019 9:14:57
    Const THUNK_SIZE    As Long = 5652
    Static hThunk       As Long
    Dim aParams(0 To 9) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitFireOnceTimerThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        If hThunk = 0 Then
            Exit Function
        End If
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetTimer")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "KillTimer")
        '--- for IDE protection
        Debug.Assert pvGetIdeOwner(aParams(6))
        If aParams(6) <> 0 Then
            aParams(7) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(8) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitFireOnceTimerThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, 0, Delay, VarPtr(aParams(0)), VarPtr(InitFireOnceTimerThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function pvGetIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvGetIdeOwner = True
End Function

Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property
#End If

'=========================================================================
' Event handlers
'=========================================================================

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    pvPrepareFont m_oFont, m_hFont
    pvRefresh
    PropertyChanged
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "ScaleUnits" Then
        m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
    End If
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    If Ambient.UserMode Then
        HitResult = pvHitTest(X, Y)
    Else
        HitResult = vbHitResultHit
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii = 32 Or KeyAscii = 13 Then
        pvHandleClick
        KeyAscii = 0
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent AccessKeyPress(KeyAscii)
    If KeyAscii <> 0 Then
        pvHandleClick
        KeyAscii = 0
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, ScaleX(X, ScaleMode, m_eContainerScaleMode), ScaleY(Y, ScaleMode, m_eContainerScaleMode))
    pvHandleMouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "UserControl_MouseMove"
    
    On Error GoTo EH
    RaiseEvent MouseMove(Button, Shift, ScaleX(X, ScaleMode, m_eContainerScaleMode), ScaleY(Y, ScaleMode, m_eContainerScaleMode))
    If Button = -1 Then
        GoTo QH
    End If
    If X >= 0 And X < ScaleWidth And Y >= 0 And Y < ScaleHeight Then
        If Not pvState(ucsBstHover) Then
            If pvParentRegisterCancelMode(Me) Then
                pvState(ucsBstHover) = True
            End If
        End If
        pvState(ucsBstPressed) = (Button And vbLeftButton) <> 0
    Else
        pvState(ucsBstPressed) = False
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "UserControl_MouseUp"
    
    On Error GoTo EH
    RaiseEvent MouseUp(Button, Shift, ScaleX(X, ScaleMode, m_eContainerScaleMode), ScaleY(Y, ScaleMode, m_eContainerScaleMode))
    If Button = -1 Then
        GoTo QH
    End If
    If (Button And vbLeftButton) <> 0 Then
        pvState(ucsBstPressed) = False
    End If
    If Button <> 0 And X >= 0 And X < ScaleWidth And Y >= 0 And Y < ScaleHeight Then
        If (m_nDownButton And Button And vbLeftButton) <> 0 Then
            Call ApiUpdateWindow(ContainerHwnd) '--- pump WM_PAINT
            RaiseEvent Click
        ElseIf (m_nDownButton And Button And vbRightButton) <> 0 Then
            RaiseEvent ContextMenu
        End If
    Else
        pvState(ucsBstHover) = False
    End If
    m_nDownButton = 0
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_DblClick()
    pvHandleMouseDown vbLeftButton, m_nDownShift, m_sngDownX, m_sngDownY
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Paint()
    Const FUNC_NAME     As String = "UserControl_Paint"
    Const AC_SRC_ALPHA  As Long = 1
    Const Opacity       As Long = &HFF
    Const CLR_YELLOW    As Long = &HE0FFFF
    Dim hMemDC          As Long
    Dim hPrevDib        As Long
    
    On Error GoTo EH
    If AutoRedraw Then
        hMemDC = CreateCompatibleDC(hDC)
        If hMemDC = 0 Then
            GoTo DefPaint
        End If
        If m_hRedrawDib = 0 Then
            If Not pvCreateDib(hMemDC, ScaleWidth, ScaleHeight, m_hRedrawDib) Then
                GoTo DefPaint
            End If
            hPrevDib = SelectObject(hMemDC, m_hRedrawDib)
            If Not pvPaintControl(hMemDC) Then
                GoTo DefPaint
            End If
        Else
            hPrevDib = SelectObject(hMemDC, m_hRedrawDib)
        End If
        If AlphaBlend(hDC, 0, 0, ScaleWidth, ScaleHeight, hMemDC, 0, 0, ScaleWidth, ScaleHeight, AC_SRC_ALPHA * &H1000000 + Opacity * &H10000) = 0 Then
            GoTo DefPaint
        End If
    Else
        If Not pvPaintControl(hDC) Then
            GoTo DefPaint
        End If
    End If
    If False Then
DefPaint:
        If m_hRedrawDib <> 0 Then
            '--- note: before deleting DIB try de-selecting from dc
            Call SelectObject(hMemDC, hPrevDib)
            Call DeleteObject(m_hRedrawDib)
            m_hRedrawDib = 0
        End If
        Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), CLR_YELLOW, BF
        Line (2, 2)-(ScaleWidth - 3, ScaleHeight - 3), vbButtonShadow, B
    End If
QH:
    On Error Resume Next
    If hMemDC <> 0 Then
        Call SelectObject(hMemDC, hPrevDib)
        Call DeleteDC(hMemDC)
        hMemDC = 0
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub UserControl_EnterFocus()
    Const FUNC_NAME     As String = "UserControl_EnterFocus"
    
    On Error GoTo EH
    If Not m_bManualFocus Then
        pvState(ucsBstFocused) = True
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_ExitFocus()
    Const FUNC_NAME     As String = "UserControl_ExitFocus"
    
    On Error GoTo EH
    If Not m_bManualFocus Then
        pvState(ucsBstFocused) = False
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_InitProperties()
    Const FUNC_NAME     As String = "UserControl_InitProperties"
    
    On Error GoTo EH
    m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
    Style = DEF_STYLE
    Enabled = DEF_ENABLED
    Opacity = DEF_OPACITY
    AnimationDuration = DEF_ANIMATIONDURATION
    Caption = Ambient.DisplayName
    Set Font = Ambient.Font
    ForeColor = DEF_FORECOLOR
    ManualFocus = DEF_MANUALFOCUS
    MaskColor = DEF_MASKCOLOR
    AutoRedraw = DEF_AUTOREDRAW
    On Error GoTo QH
    m_sInstanceName = TypeName(Extender.Parent) & "." & Extender.Name
    #If DebugMode Then
        DebugInstanceName m_sInstanceName, m_sDebugID
    #End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
    With PropBag
        Style = .ReadProperty("Style", DEF_STYLE)
        Enabled = .ReadProperty("Enabled", DEF_ENABLED)
        Opacity = .ReadProperty("Opacity", DEF_OPACITY)
        AnimationDuration = .ReadProperty("AnimationDuration", DEF_ANIMATIONDURATION)
        Caption = .ReadProperty("Caption", vbNullString)
        Set Font = .ReadProperty("Font", Ambient.Font)
        ForeColor = .ReadProperty("ForeColor", DEF_FORECOLOR)
        ManualFocus = .ReadProperty("ManualFocus", DEF_MANUALFOCUS)
        MaskColor = .ReadProperty("MaskColor", DEF_MASKCOLOR)
        AutoRedraw = .ReadProperty("AutoRedraw", DEF_AUTOREDRAW)
    End With
    On Error GoTo QH
    m_sInstanceName = TypeName(Extender.Parent) & "." & Extender.Name
    #If DebugMode Then
        DebugInstanceName m_sInstanceName, m_sDebugID
    #End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_WriteProperties"
    
    On Error GoTo EH
    With PropBag
        .WriteProperty "Style", Style, DEF_STYLE
        .WriteProperty "Enabled", Enabled, DEF_ENABLED
        .WriteProperty "Opacity", Opacity, DEF_OPACITY
        .WriteProperty "AnimationDuration", AnimationDuration, DEF_ANIMATIONDURATION
        .WriteProperty "Caption", Caption, vbNullString
        .WriteProperty "Font", Font, Ambient.Font
        .WriteProperty "ForeColor", ForeColor, DEF_FORECOLOR
        .WriteProperty "ManualFocus", ManualFocus, DEF_MANUALFOCUS
        .WriteProperty "MaskColor", MaskColor, DEF_MASKCOLOR
        .WriteProperty "AutoRedraw", AutoRedraw, DEF_AUTOREDRAW
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_Resize()
    Const FUNC_NAME     As String = "UserControl_Resize"
    
    On Error GoTo EH
    If m_hFocusBitmap <> 0 Then
        Call GdipDisposeImage(m_hFocusBitmap)
        m_hFocusBitmap = 0
    End If
    pvRefresh
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_Hide()
    Const FUNC_NAME     As String = "UserControl_Hide"
    
    On Error GoTo EH
    m_bShown = False
    If m_hPrevBitmap <> 0 Then
        Call GdipDisposeImage(m_hPrevBitmap)
        m_hPrevBitmap = 0
    End If
    CancelMode
    Set m_pTimer = Nothing
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub UserControl_Initialize()
    Dim aInput(0 To 3)  As Long
    
    #If DebugMode Then
        DebugInstanceInit MODULE_NAME, m_sDebugID, Me
    #End If
    If GetModuleHandle("gdiplus") = 0 Then
        aInput(0) = 1
        Call GdiplusStartup(0, aInput(0))
    End If
    m_eContainerScaleMode = vbTwips
    pvSetEmptyStyle
End Sub

Private Sub UserControl_Terminate()
    If m_hAttributes <> 0 Then
        Call GdipDisposeImageAttributes(m_hAttributes)
        m_hAttributes = 0
    End If
    If m_hBitmap <> 0 Then
        Call GdipDisposeImage(m_hBitmap)
        m_hBitmap = 0
    End If
    If m_hPrevBitmap <> 0 Then
        Call GdipDisposeImage(m_hPrevBitmap)
        m_hPrevBitmap = 0
    End If
    If m_hFocusBitmap <> 0 Then
        Call GdipDisposeImage(m_hFocusBitmap)
        m_hFocusBitmap = 0
    End If
    If m_hFocusAttributes <> 0 Then
        Call GdipDisposeImageAttributes(m_hFocusAttributes)
        m_hFocusAttributes = 0
    End If
    If m_hFont <> 0 Then
        Call GdipDeleteFont(m_hFont)
        m_hFont = 0
    End If
    If m_hPictureBitmap <> 0 Then
        Call GdipDisposeImage(m_hPictureBitmap)
        m_hPictureBitmap = 0
    End If
    If m_hPictureAttributes <> 0 Then
        Call GdipDisposeImageAttributes(m_hPictureAttributes)
        m_hPictureAttributes = 0
    End If
    If m_hRedrawDib <> 0 Then
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
    Set m_pTimer = Nothing
    #If DebugMode Then
        DebugInstanceTerm MODULE_NAME, m_sDebugID
    #End If
End Sub
