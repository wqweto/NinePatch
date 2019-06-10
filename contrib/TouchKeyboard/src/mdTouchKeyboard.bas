Attribute VB_Name = "mdTouchKeyboard"
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdTouchKeyboard"

#Const ImplUseShared = NPPNG_USE_SHARED <> 0

'=========================================================================
' Public enums
'=========================================================================

#If Not ImplUseShared Then
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
#End If

'==============================================================================
' API
'==============================================================================

'--- for VirtualProtect
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const MEM_COMMIT                    As Long = &H1000
'--- for CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1
'--- for gdi+
Private Const ImageLockModeRead             As Long = &H1
Private Const ImageLockModeWrite            As Long = &H2
Private Const PixelFormat32bppARGB          As Long = &H26200A
'--- for GdipCreateFont
Private Const UnitPixel                     As Long = 2
Private Const UnitPoint                     As Long = 3
'--- for GdipSetPenDashStyle
Private Const DashStyleSolid                As Long = 0
'--- GDI+ colors
Private Const Transparent                   As Long = &HFFFFFF
'--- for GdipSetSmoothingMode
Private Const SmoothingModeAntiAlias        As Long = 4
'--- for GdipCreatePath
Private Const FillModeAlternate             As Long = 0

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'--- gdi+
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, lpRect As Any, ByVal lFlags As Long, ByVal lPixelFormat As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal lColor As Long, hBrush As Long) As Long
Private Declare Function GdipNewPrivateFontCollection Lib "gdiplus" (hFontCollection As Long) As Long
Private Declare Function GdipPrivateAddFontFile Lib "gdiplus" (ByVal hFontCollection As Long, ByVal lpFileName As Long) As Long
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal hFontFamily As Long, ByVal emSize As Single, ByVal lStyle As Long, ByVal lUnit As Long, hFont As Long) As Long
Private Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal hFontCollection As Long, ByVal lNumSought As Long, aFamilies As Any, lNumFound As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal hImage As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal lWidth As Long, ByVal lHeight As Long, ByVal lStride As Long, ByVal lPixelFormat As Long, ByVal Scan0 As Long, hBitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As Long
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal srcUnit As Long = UnitPixel, Optional ByVal hImageAttributes As Long, Optional ByVal pfnCallback As Long, Optional ByVal lCallbackData As Long) As Long
Private Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal hBitmap As Long, ByVal lX As Long, ByVal lY As Long, ByVal lColor As Long) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal hGraphics As Long, ByVal lColor As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal lSmoothingMd As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal lColor As Long, ByVal sngWidth As Single, ByVal lUnit As Long, hPen As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal hPen As Long, ByVal dStyle As Long) As Long
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal lBrushmode As Long, hPath As Long) As Long
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal hPath As Long, ByVal sngX As Single, ByVal sngY As Single, ByVal sngWidth As Single, ByVal sngHeight As Single, ByVal sngStartAngle As Single, ByVal sngSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal hPath As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal hPath As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal pen As Long, ByVal hPath As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal hPen As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal hPath As Long) As Long
Private Declare Function GdipBeginContainer2 Lib "gdiplus" (ByVal hGraphics As Long, hState As Long) As Long
Private Declare Function GdipEndContainer Lib "gdiplus" (ByVal hGraphics As Long, ByVal hState As Long) As Long
'--- public
Public Declare Function GdipDeletePrivateFontCollection Lib "gdiplus" (hFontCollection As Long) As Long
#If Not ImplUseShared Then
    Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
    Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal hFormatAttributes As Long, ByVal nLanguage As Integer, hStringFormat As Long) As Long
    Private Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal hStringFormat As Long, ByVal lFlags As Long) As Long
    Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal hStringFormat As Long, ByVal eAlign As StringAlignment) As Long
    Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal hStringFormat As Long, ByVal eAlign As StringAlignment) As Long
    Private Declare Function GdipCreateFontFromDC Lib "gdiplus" (ByVal hDC As Long, hCreatedFont As Long) As Long
    '--- public
    Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal hStringFormat As Long) As Long
    Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal hFont As Long) As Long
    Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As Long
#End If

Private Type BitmapData
    Width               As Long
    Height              As Long
    Stride              As Long
    PixelFormat         As Long
    Scan0               As Long
    Reserved            As Long
End Type

Private Enum StringAlignment
   StringAlignmentNear = 0
   StringAlignmentCenter = 1
   StringAlignmentFar = 2
End Enum

Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    B                   As Byte
    A                   As Byte
End Type

Private Type UcsHsbColor
    Hue                 As Single
    Sat                 As Single
    Bri                 As Single
    A                   As Byte
End Type

#If Not ImplUseShared Then
    Private Type ThunkBytes
        Thunk(5)            As Long
    End Type
    
    Private Type PushParamThunk
        pfn                 As Long
        Code                As ThunkBytes
    End Type
    
    Public Type FireOnceTimerData
        TimerID             As Long
        TimerProcThunkData  As PushParamThunk
        TimerProcThunkThis  As PushParamThunk
    End Type
#End If

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
#If ImplUseShared Then
    PopPrintError sFunction, MODULE_NAME, PushError
#Else
    Debug.Print "Ciritical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]", Timer
#End If
End Sub

'==============================================================================
' Functions
'==============================================================================

Public Function GdipBlurBitmap( _
            ByVal hBitmap As Long, _
            ByVal sngRadius As Single, _
            Optional ByVal AffectChannels As Long = 15) As Boolean
    Const FUNC_NAME     As String = "BlurBitmap"
    Dim uData           As BitmapData
    Dim dblBuffer()     As Double

    On Error GoTo EH
    If GdipBitmapLockBits(hBitmap, ByVal 0, ImageLockModeRead Or ImageLockModeWrite, PixelFormat32bppARGB, uData) <> 0 Then
        GoTo QH
    End If
    ReDim dblBuffer(0 To uData.Width - 1, 0 To uData.Height - 1) As Double
    If (AffectChannels And 1) <> 0 Then
        If Not pvBlurChannel(uData.Scan0, uData.Stride \ 4, 0, 0, uData.Width, uData.Height, sngRadius, 0, dblBuffer) Then
            GoTo QH
        End If
    End If
    If (AffectChannels And 2) <> 0 Then
        If Not pvBlurChannel(uData.Scan0, uData.Stride \ 4, 0, 0, uData.Width, uData.Height, sngRadius, 1, dblBuffer) Then
            GoTo QH
        End If
    End If
    If (AffectChannels And 4) <> 0 Then
        If Not pvBlurChannel(uData.Scan0, uData.Stride \ 4, 0, 0, uData.Width, uData.Height, sngRadius, 2, dblBuffer) Then
            GoTo QH
        End If
    End If
    If (AffectChannels And 8) <> 0 Then
        If Not pvBlurChannel(uData.Scan0, uData.Stride \ 4, 0, 0, uData.Width, uData.Height, sngRadius, 3, dblBuffer) Then
            GoTo QH
        End If
    End If
    '--- success
    GdipBlurBitmap = True
QH:
    On Error Resume Next
    If uData.Scan0 <> 0 Then
        Call GdipBitmapUnlockBits(hBitmap, uData)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvBlurChannel( _
            ByVal lpBits As Long, _
            ByVal lStride As Long, _
            ByVal lLeft As Long, _
            ByVal lTop As Long, _
            ByVal lWidth As Long, _
            ByVal lHeight As Long, _
            ByVal dblRadius As Double, _
            ByVal lChannel As Long, _
            dblBuffer() As Double) As Boolean
'--- Gaussian blur filter, using an IIR (Infininte Impulse Response) approach
'--- based on https://github.com/tannerhelland/PhotoDemon/blob/master/Modules/Filters_ByteArray.bas#L40
    Const NUM_ITERS     As Long = 3
    Dim lIdx            As Long
    Dim lIter           As Long
    Dim dblTemp         As Double
    Dim dblNu           As Double
    Dim dblBndryScale   As Double
    Dim dblPostScale    As Double

    ' Prep some IIR-specific values
    dblTemp = Sqr(-(dblRadius * dblRadius) / (2 * Log(1 / 255)))
    If dblTemp <= 0 Then
        dblTemp = 0.01
    End If
    dblTemp = dblTemp * (1 + (0.3165 * NUM_ITERS + 0.5695) / ((NUM_ITERS + 0.7818) * (NUM_ITERS + 0.7818)))
    dblTemp = (dblTemp * dblTemp) / (2 * NUM_ITERS)
    dblNu = (1 + 2 * dblTemp - Sqr(1 + 4 * dblTemp)) / (2 * dblTemp)
    dblBndryScale = (1 / (1 - dblNu))
    dblPostScale = ((dblNu / dblTemp) ^ (2 * NUM_ITERS)) * 255
    ' Copy the contents of the incoming byte array into the double array buffer
    LoadSave dblBuffer(0, 0), 1 / 255, lpBits + (lTop * lStride + lLeft) * 4 + lChannel, lStride, lWidth, lHeight, 0
    ' Filter horizontally along each row
    For lIdx = 0 To lHeight - 1
        For lIter = 1 To NUM_ITERS
            ProcessRow dblBuffer(0, lIdx), dblBndryScale, dblNu, 1, lWidth
            ProcessRow dblBuffer(lWidth - 1, lIdx), dblBndryScale, dblNu, -1, lWidth
        Next
    Next
    ' Now repeat all the above steps, but filtering vertically along each column, instead
    For lIdx = 0 To lWidth - 1
        For lIter = 1 To NUM_ITERS
            ProcessRow dblBuffer(lIdx, 0), dblBndryScale, dblNu, lWidth, lHeight
            ProcessRow dblBuffer(lIdx, lHeight - 1), dblBndryScale, dblNu, -lWidth, lHeight
        Next
    Next
    ' Apply final post-scaling and copy back to byte array
    LoadSave dblBuffer(0, 0), dblPostScale, lpBits + (lTop * lStride + lLeft) * 4 + lChannel, lStride, lWidth, lHeight, 1
    '--- success
    pvBlurChannel = True
End Function

Private Sub LoadSave(dblPtr As Double, ByVal dblScale As Double, ByVal srcPtr As Long, ByVal lStride As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal fSave As Long)
    'void __stdcall LoadSave(double *ptr, double scale, unsigned char *src, int stride, int w, int h, int fsave) {
    '    for(int j = 0; j < h; j++, src += 4*(stride - w)) {
    '        for(int i = 0; i < w; i++, src += 4, ptr++) {
    '            if (!fsave)
    '                *ptr = *src * scale;
    '            else {
    '                int v = *ptr * scale;
    '                *src = v > 0xFF ? 0xFF : v < 0 ? 0 : v;
    '            }
    '        }
    '    }
    '}
    Const STR_THUNK     As String = "VYvsU4tdIIXbD46DAAAAi00Yi0UcK8jyDxBNDItVCMHhAlaJTRiLTRSLdRhXi30khcB+VYvwhf91FQ" & _
                                    "+2AWYPbsDzD+bA8g9ZwfIPEQLrKfIPEALyD1nB8g8swD3/AAAAfge4/wAAAOsNhcDHRSAAAAAAD0hF" & _
                                    "IIgBg8EEg8IIg+4BdbOLRRyLdRgDzoPrAXWgX15bXcIgAA=="
    pvPatchThunk AddressOf mdTouchKeyboard.LoadSave, STR_THUNK
    LoadSave dblPtr, dblScale, srcPtr, lStride, lWidth, lHeight, fSave
End Sub

Private Sub ProcessRow(dblPtr As Double, ByVal dblBndryScale As Double, ByVal dblNu As Double, ByVal lStep As Long, ByVal lSize As Long)
    'void __stdcall ProcessRow(double *ptr, double bndry, double nu, int step, int size) {
    '    double temp = (*ptr *= bndry);
    '    ptr += step;
    '    for(int i = 1; i < size; i++) {
    '        temp = (*ptr += nu * temp);
    '        ptr += step;
    '    }
    '}
    Const STR_THUNK     As String = "VYvsi00Ii0UcVlfyDxABvwEAAADyD1lFDI0UxQAAAACLRSDyDxEBA8o7x3558g8QTRSD+AR+Vo1w+" & _
                                    "8HuAkaNPLUBAAAAZmZmDx+EAAAAAADyD1nB8g9YAfIPEQEDyvIPWcHyD1gB8g8RAQPK8g9ZwfIPWA" & _
                                    "HyDxEBA8ryD1nB8g9YAfIPEQEDyoPuAXXDO/h9FSvH8g9ZwfIPWAHyDxEBA8qD6AF17V9eXcIcAA=="
    pvPatchThunk AddressOf mdTouchKeyboard.ProcessRow, STR_THUNK
    ProcessRow dblPtr, dblBndryScale, dblNu, lStep, lSize
End Sub

Private Sub pvPatchThunk(ByVal pfn As Long, sThunkStr As String)
    Dim lThunkSize      As Long
    Dim lThunkPtr       As Long
    Dim bInIDE          As Boolean

    '--- decode thunk
    Call CryptStringToBinary(sThunkStr, Len(sThunkStr), CRYPT_STRING_BASE64, 0, lThunkSize, 0, 0)
    lThunkPtr = VirtualAlloc(0, lThunkSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CryptStringToBinary(sThunkStr, Len(sThunkStr), CRYPT_STRING_BASE64, lThunkPtr, lThunkSize, 0, 0)
    '--- patch func
    Debug.Assert pvSetTrue(bInIDE)
    If bInIDE Then
        Call CopyMemory(pfn, ByVal pfn + &H16, 4)
    Else
        Call VirtualProtect(pfn, 8, PAGE_EXECUTE_READWRITE, 0)
    End If
    ' B8 00 00 00 00       mov         eax,00000000h
    ' FF E0                jmp         eax
    Call CopyMemory(ByVal pfn, 6333077358968.8504@, 8)
    Call CopyMemory(ByVal (pfn Xor &H80000000) + 1 Xor &H80000000, lThunkPtr, 4)
End Sub

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Public Function GdipPreparePrivateFont(sFileName As String, ByVal lFontSize As Long, hFont As Long, hFontCollection As Long) As Boolean
    Dim hNewFontCol     As Long
    Dim hFamily         As Long
    Dim lNumFamilies    As Long
    Dim hNewFont        As Long
    
    If hFontCollection = 0 Then
        If GdipNewPrivateFontCollection(hNewFontCol) <> 0 Then
            GoTo QH
        End If
        If GdipPrivateAddFontFile(hNewFontCol, StrPtr(sFileName)) <> 0 Then
            GoTo QH
        End If
    Else
        hNewFontCol = hFontCollection
    End If
    If GdipGetFontCollectionFamilyList(hNewFontCol, 1, hFamily, lNumFamilies) <> 0 Or lNumFamilies = 0 Then
        GoTo QH
    End If
    If GdipCreateFont(hFamily, lFontSize, 0, UnitPoint, hNewFont) <> 0 Then
        GoTo QH
    End If
    '--- commit
    If hFont <> 0 Then
        Call GdipDeleteFont(hFont)
    End If
    hFont = hNewFont
    hNewFont = 0
    If hFontCollection <> 0 And hFontCollection <> hNewFontCol Then
        Call GdipDeletePrivateFontCollection(hFontCollection)
    End If
    hFontCollection = hNewFontCol
    hNewFontCol = 0
    '--- success
    GdipPreparePrivateFont = True
QH:
    If hNewFont <> 0 Then
        Call GdipDeleteFont(hNewFont)
        hNewFont = 0
    End If
    If hNewFontCol <> 0 And hFontCollection <> hNewFontCol Then
        Call GdipDeletePrivateFontCollection(hNewFontCol)
        hNewFontCol = 0
    End If
End Function

Public Function GdipPrepareButtonBitmap( _
            ByVal sngRadius As Single, _
            ByVal sngBlur As Single, _
            ByVal clrPen As Long, _
            ByVal clrBack As Long, _
            ByVal clrShadow As Long, _
            hBitmap As Long) As Boolean
    Const FUNC_NAME     As String = "GdipPrepareButtonBitmap"
    Const SHADOW_OFFSET As Single = 1
    Const CLR_WHITE     As Long = &HFFFFFFFF
    Const CLR_BLACK     As Long = &HFF000000
    Dim lIdx            As Long
    Dim lRoundWidth     As Long
    Dim lWidth          As Long
    Dim hNewBitmap      As Long
    Dim hDropShadow     As Long
    Dim hGraphics       As Long
    Dim hBrush          As Long
    
    On Error GoTo EH
    lRoundWidth = Ceil(sngRadius) + 1 + Ceil(sngRadius)
    lWidth = 1 + Ceil(sngBlur) + lRoundWidth + Ceil(SHADOW_OFFSET) + Ceil(sngBlur) + 1
    If GdipCreateBitmapFromScan0(lWidth, lWidth, lWidth * 4, PixelFormat32bppARGB, 0, hNewBitmap) <> 0 Then
        GoTo QH
    End If
    If GdipCreateBitmapFromScan0(lWidth, lWidth, lWidth * 4, PixelFormat32bppARGB, 0, hDropShadow) <> 0 Then
        GoTo QH
    End If
    If GdipGetImageGraphicsContext(hDropShadow, hGraphics) <> 0 Then
        GoTo QH
    End If
    If GdipGraphicsClear(hGraphics, &H1000000 Or (clrShadow And &HFFFFFF)) <> 0 Then
        GoTo QH
    End If
    If Not GdipDrawRoundedRectangle(hGraphics, 1 + Ceil(sngBlur) + SHADOW_OFFSET / 2, 1 + Ceil(sngBlur) + SHADOW_OFFSET, _
            lRoundWidth, lRoundWidth, sngRadius, clrShadow, clrBack:=clrShadow) Then
        GoTo QH
    End If
    If Not GdipBlurBitmap(hDropShadow, sngBlur) Then
        GoTo QH
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    If GdipGetImageGraphicsContext(hNewBitmap, hGraphics) <> 0 Then
        GoTo QH
    End If
    For lIdx = 1 To 2
        If GdipDrawImageRectRect(hGraphics, hDropShadow, 0, 0, lWidth, lWidth, 0, 0, lWidth, lWidth) <> 0 Then
            GoTo QH
        End If
    Next
    If Not GdipDrawRoundedRectangle(hGraphics, 1 + Ceil(sngBlur), 1 + Ceil(sngBlur), lRoundWidth, lRoundWidth, sngRadius, clrPen, clrBack:=clrBack) Then
        GoTo QH
    End If
    '--- draw nine-patch markers
    If GdipCreateSolidFill(CLR_WHITE, hBrush) <> 0 Then
        GoTo QH
    End If
    Call GdipFillRectangleI(hGraphics, hBrush, 0, 0, lWidth, 1)
    Call GdipFillRectangleI(hGraphics, hBrush, 0, lWidth - 1, lWidth, 1)
    Call GdipFillRectangleI(hGraphics, hBrush, lWidth - 1, 0, 1, lWidth)
    Call GdipFillRectangleI(hGraphics, hBrush, 0, 0, 1, lWidth)
    lIdx = 1 + Ceil(sngBlur) + Ceil(sngRadius)
    Call GdipDeleteBrush(hBrush)
    If GdipCreateSolidFill(CLR_BLACK, hBrush) <> 0 Then
        GoTo QH
    End If
    Call GdipBitmapSetPixel(hNewBitmap, lIdx, 0, CLR_BLACK)
    Call GdipBitmapSetPixel(hNewBitmap, 0, lIdx, CLR_BLACK)
    Call GdipFillRectangleI(hGraphics, hBrush, lWidth - 1, 1, 1, lWidth - 2)
    Call GdipFillRectangleI(hGraphics, hBrush, 1, lWidth - 1, lWidth - 2, 1)
    '--- commit
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    hBitmap = hNewBitmap
    hNewBitmap = 0
    '--- success
    GdipPrepareButtonBitmap = True
QH:
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    If hNewBitmap <> 0 Then
        Call GdipDisposeImage(hNewBitmap)
    End If
    If hDropShadow <> 0 Then
        Call GdipDisposeImage(hDropShadow)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function GdipDrawRoundedRectangle( _
            ByVal hGraphics As Long, _
            ByVal sngLeft As Single, _
            ByVal sngTop As Single, _
            ByVal sngWidth As Single, _
            ByVal sngHeight As Single, _
            ByVal sngRadius As Single, _
            Optional ByVal clrPen As Long = Transparent, _
            Optional ByVal PenWidth As Long = 1, _
            Optional ByVal DashStyle As Long = DashStyleSolid, _
            Optional ByVal clrBack As Long = Transparent) As Boolean
    Const FUNC_NAME     As String = "GdipDrawRoundedRectangle"
    Dim hState          As Long
    Dim hPen            As Long
    Dim hPath           As Long
    Dim hBrush          As Long
    Dim sngDiam         As Single
    
    On Error GoTo EH
    sngDiam = sngRadius + sngRadius
    '--- setup graphics
    If GdipBeginContainer2(hGraphics, hState) <> 0 Then
        GoTo QH
    End If
    If GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias) <> 0 Then
        GoTo QH
    End If
    '--- setup path
    If GdipCreatePath(FillModeAlternate, hPath) <> 0 Then
        GoTo QH
    End If
    If GdipAddPathArc(hPath, sngLeft + sngWidth - sngDiam, sngTop, sngDiam, sngDiam, 270, 90) <> 0 Then
        GoTo QH
    End If
    If GdipAddPathArc(hPath, sngLeft + sngWidth - sngDiam, sngTop + sngHeight - sngDiam, sngDiam, sngDiam, 0, 90) <> 0 Then
        GoTo QH
    End If
    If GdipAddPathArc(hPath, sngLeft, sngTop + sngHeight - sngDiam, sngDiam, sngDiam, 90, 90) <> 0 Then
        GoTo QH
    End If
    If GdipAddPathArc(hPath, sngLeft, sngTop, sngDiam, sngDiam, 180, 90) <> 0 Then
        GoTo QH
    End If
    If GdipClosePathFigure(hPath) <> 0 Then
        GoTo QH
    End If
    '--- setup brush
    If clrBack <> Transparent Then
        If GdipCreateSolidFill(clrBack, hBrush) <> 0 Then
            GoTo QH
        End If
        Call GdipFillPath(hGraphics, hBrush, hPath)
    End If
    '--- setup pen
    If clrPen <> Transparent Then
        If GdipCreatePen1(clrPen, PenWidth, UnitPixel, hPen) <> 0 Then
            GoTo QH
        End If
        If GdipSetPenDashStyle(hPen, DashStyle) <> 0 Then
            GoTo QH
        End If
        Call GdipDrawPath(hGraphics, hPen, hPath)
    End If
    If GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias) <> 0 Then
        GoTo QH
    End If
    '--- success
    GdipDrawRoundedRectangle = True
QH:
    On Error Resume Next
    '--- cleanup
    If hPen <> 0 Then
        Call GdipDeletePen(hPen)
    End If
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
    End If
    If hPath <> 0 Then
        Call GdipDeletePath(hPath)
    End If
    If hState <> 0 Then
        Call GdipEndContainer(hGraphics, hState)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function GdipAdjustColor( _
            ByVal clrValue As Long, _
            Optional ByVal AdjustBri As Single, _
            Optional ByVal AdjustSat As Single, _
            Optional ByVal AdjustAlpha As Single) As Long
    Dim hsbColor        As UcsHsbColor
    
    hsbColor = pvRGBToHSB(clrValue)
    If AdjustBri > 0 Then
        hsbColor.Bri = 1 - (1 - hsbColor.Bri) * (1 - AdjustBri)
    Else
        hsbColor.Bri = hsbColor.Bri * (1 + AdjustBri)
    End If
    If AdjustSat > 0 Then
        hsbColor.Sat = 1 - (1 - hsbColor.Sat) * (1 - AdjustSat)
    Else
        hsbColor.Sat = hsbColor.Sat * (1 + AdjustSat)
    End If
    If AdjustAlpha > 0 Then
        hsbColor.A = 1 - (1 - hsbColor.A) * (1 - AdjustAlpha)
    Else
        hsbColor.A = hsbColor.A * (1 + AdjustAlpha)
    End If
    GdipAdjustColor = pvHSBToRGB(hsbColor)
End Function

Private Function pvHSBToRGB(hsbColor As UcsHsbColor) As Long
    Dim lMax            As Long
    Dim lMid            As Long
    Dim lMin            As Long
    Dim sngDelta        As Single
    Dim rgbColor        As UcsRgbQuad
    Dim lHue            As Long

    lMax = hsbColor.Bri * 255
    lMin = (1 - hsbColor.Sat) * CSng(lMax)
    sngDelta = CSng(lMax - lMin) / 60
    With rgbColor
        lHue = hsbColor.Hue * 360
        Select Case lHue
        Case 0 To 60
            lMid = (lHue - 0) * sngDelta + lMin
            .R = lMax: .G = lMid: .B = lMin
        Case 60 To 120
            lMid = -(lHue - 120) * sngDelta + lMin
            .R = lMid: .G = lMax: .B = lMin
        Case 120 To 180
            lMid = (lHue - 120) * sngDelta + lMin
            .R = lMin: .G = lMax: .B = lMid
        Case 180 To 240
            lMid = -(lHue - 240) * sngDelta + lMin
            .R = lMin: .G = lMid: .B = lMax
        Case 240 To 300
            lMid = (lHue - 240) * sngDelta + lMin
            .R = lMid: .G = lMin: .B = lMax
        Case 300 To 360
            lMid = -(lHue - 360) * sngDelta + lMin
            .R = lMax: .G = lMin: .B = lMid
        End Select
        .A = hsbColor.A
    End With
    Call CopyMemory(pvHSBToRGB, rgbColor, 4)
End Function

Private Function pvRGBToHSB(ByVal clrValue As Long) As UcsHsbColor
    Dim rgbColor        As UcsRgbQuad
    Dim lMin            As Long
    Dim lMax            As Long
    Dim sngDelta        As Single
  
    Call CopyMemory(rgbColor, clrValue, 4)
    With rgbColor
        If .R > .G Then
            lMax = .R: lMin = .G
        Else
            lMax = .G: lMin = .R
        End If
        If .B > lMax Then
            lMax = .B
        ElseIf .B < lMin Then
            lMin = .B
        End If
        pvRGBToHSB.Bri = CSng(lMax) / 255
        sngDelta = lMax - lMin
        If sngDelta > 0 Then
            '--- note: sngDelta > 0 => lMax > 0
            pvRGBToHSB.Sat = sngDelta / lMax
            Select Case lMax
            Case .R
                pvRGBToHSB.Hue = (0 + (CSng(.G) - .B) / sngDelta) / 6
            Case .G
                pvRGBToHSB.Hue = (2 + (CSng(.B) - .R) / sngDelta) / 6
            Case Else
                pvRGBToHSB.Hue = (4 + (CSng(.R) - .G) / sngDelta) / 6
            End Select
            If pvRGBToHSB.Hue < 0 Then
                pvRGBToHSB.Hue = pvRGBToHSB.Hue + 1
            End If
        End If
        pvRGBToHSB.A = .A
'        Debug.Assert pvHSBToRGB(pvRGBToHSB) = clrValue
    End With
End Function

Private Function Ceil(ByVal Value As Single) As Single
    Ceil = -Int(CStr(-Value))
End Function

#If Not ImplUseShared Then
Private Sub PatchMethodProto(ByVal pfn As Long, ByVal lMethodIdx As Long)
    Dim bInIDE          As Boolean
    
    Debug.Assert pvSetTrue(bInIDE)
    If bInIDE Then
        '--- note: IDE is not large-address aware
        Call CopyMemory(pfn, ByVal pfn + &H16, 4)
    Else
        Call VirtualProtect(pfn, 12, PAGE_EXECUTE_READWRITE, 0)
    End If
    ' 0: 8B 44 24 04          mov         eax,dword ptr [esp+4]
    ' 4: 8B 00                mov         eax,dword ptr [eax]
    ' 6: FF A0 00 00 00 00    jmp         dword ptr [eax+lMethodIdx*4]
    Call CopyMemory(ByVal pfn, -684575231150992.4725@, 8)
    Call CopyMemory(ByVal (pfn Xor &H80000000) + 8 Xor &H80000000, lMethodIdx * 4, 4)
End Sub
 
Private Function TryGetValue(ByVal oCol As Collection, Index As Variant, RetVal As Variant) As Long
    Const IDX_COLLECTION_ITEM   As Long = 7
    PatchMethodProto AddressOf mdTouchKeyboard.TryGetValue, IDX_COLLECTION_ITEM
    TryGetValue = TryGetValue(oCol, Index, RetVal)
End Function

Public Function SearchCollection(oCol As Collection, Index As Variant, Optional RetVal As Variant) As Boolean
    If Not oCol Is Nothing Then
        SearchCollection = TryGetValue(oCol, Index, RetVal) = 0 ' S_OK
    End If
End Function

Public Function GdipPrepareFont(oFont As StdFont, hFont As Long) As Boolean
    Const FUNC_NAME     As String = "GdipPrepareFont"
    Dim hDC             As Long
    Dim pFont           As IFont
    Dim hPrevFont       As Long
    Dim hNewFont        As Long
    
    On Error GoTo EH
    Set pFont = oFont
    If pFont Is Nothing Then
        GoTo QH
    End If
    hDC = GetDC(0)
    If hDC = 0 Then
        GoTo QH
    End If
    hPrevFont = SelectObject(hDC, pFont.hFont)
    If hPrevFont = 0 Then
        GoTo QH
    End If
    If GdipCreateFontFromDC(hDC, hNewFont) <> 0 Then
        GoTo QH
    End If
    '--- commit
    If hFont <> 0 Then
        Call GdipDeleteFont(hFont)
    End If
    hFont = hNewFont
    hNewFont = 0
    '--- success
    GdipPrepareFont = True
QH:
    If hNewFont <> 0 Then
        Call GdipDeleteFont(hNewFont)
        hNewFont = 0
    End If
    If hPrevFont <> 0 Then
        Call SelectObject(hDC, hPrevFont)
        hPrevFont = 0
    End If
    If hDC <> 0 Then
        Call ReleaseDC(0, hDC)
        hDC = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function GdipPrepareStringFormat(ByVal lFlags As UcsNineButtonTextFlagsEnum, hStringFormat As Long) As Boolean
    Const FUNC_NAME     As String = "GdipPrepareStringFormat"
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
    GdipPrepareStringFormat = True
QH:
    If hNewFormat <> 0 Then
        Call GdipDeleteStringFormat(hNewFormat)
        hNewFormat = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function GdipPrepareSolidBrush(ByVal clrValue As OLE_COLOR, hBrush As Long, Optional ByVal Alpha As Single = 1) As Boolean
    Const FUNC_NAME     As String = "GdipPrepareSolidBrush"
    Dim hNewBrush       As Long
    
    On Error GoTo EH
    If GdipCreateSolidFill(GdipTranslateColor(clrValue, Alpha), hNewBrush) <> 0 Then
        GoTo QH
    End If
    '--- commit
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
    End If
    hBrush = hNewBrush
    hNewBrush = 0
    '--- success
    GdipPrepareSolidBrush = True
QH:
    If hNewBrush <> 0 Then
        Call GdipDeleteBrush(hNewBrush)
        hNewBrush = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function GdipTranslateColor(ByVal clrValue As OLE_COLOR, Optional ByVal Alpha As Single = 1) As Long
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
    Call CopyMemory(GdipTranslateColor, uQuad, 4)
End Function
#End If ' ImplUseShared
