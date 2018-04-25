VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNinePatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' Nine Patch PNGs for VB6 (c) 2018 by wqweto@gmail.com
'
' cNinePatch.cls -- can parse and draw nine-patch PNGs
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "cNinePatch"

#Const ImplUseShared = NPPNG_USE_SHARED <> 0

'=========================================================================
' API
'=========================================================================

'--- for GlobalAlloc
Private Const GMEM_MOVEABLE                 As Long = 2
'--- for GdipBitmapLockBits
Private Const PixelFormat32bppARGB          As Long = &H26200A
'--- for GdipSetInterpolationMode
Private Const InterpolationModeHighQuality  As Long = 2
'--- for GdipSetPixelOffsetMode
Private Const PixelOffsetModeHighQuality    As Long = 2
'--- for GdipDrawImageXxx
Private Const UnitPixel                     As Long = 2
'--- for GdipSetCompositingMode
Private Const CompositingModeSourceCopy     As Long = 1
'--- for GdipTranslateMatrix, GdipScaleTextureTransform
Private Const MatrixOrderAppend             As Long = 1
'--- for GdipSetClipRectI
Private Const CombineModeIntersect          As Long = 1
'--- for GdipCreateTexture2I
Private Const WrapModeTileFlipX             As Long = 1
Private Const WrapModeTileFlipY             As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, uInputBuf As GdiplusStartupInput, Optional ByVal lOutputBuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal lFilenamePtr As Long, hImage As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal pStream As IUnknown, hImage As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal hImage As Long, sngWidth As Single, sngHeight As Single) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, uRect As APIRECTL, ByVal lFlags As Long, ByVal lPixelFormat As Long, uLockedBitmapData As APIBITMAPDATA) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, uLockedBitmapData As APIBITMAPDATA) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal hImage As Long, hGraphics As Long) As Long
'Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal lPixOffsetMode As Long) As Long
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal srcUnit As Long = UnitPixel, Optional ByVal hImageAttributes As Long, Optional ByVal pfnCallback As Long, Optional ByVal lCallbackData As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipCreateMatrix Lib "gdiplus" (hMatrix As Long) As Long
Private Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal hMatrix As Long, hCloneMatrix As Long) As Long
Private Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal lOffsetX As Single, ByVal lOffsetY As Single, ByVal lOrder As Long) As Long
Private Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal hMatrix As Long) As Long
Private Declare Function GdipSetWorldTransform Lib "gdiplus" (ByVal hGraphics As Long, ByVal hMatrix As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal lColor As Long, hBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal lCompositingMode As Long) As Long
Private Declare Function GdipSetClipRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal lCombineMd As Long) As Long
Private Declare Function GdipCreateTexture2I Lib "gdiplus" (ByVal hImage As Long, ByVal lWrapMd As Long, ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long, hBrush As Long) As Long
Private Declare Function GdipScaleTextureTransform Lib "gdiplus" (ByVal hBrush As Long, ByVal sngX As Single, ByVal sngY As Single, ByVal lOrder As Long) As Long
Private Declare Function GdipBeginContainer2 Lib "gdiplus" (ByVal hGraphics As Long, hState As Long) As Long
Private Declare Function GdipEndContainer Lib "gdiplus" (ByVal hGraphics As Long, ByVal hState As Long) As Long

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type APIRECTL
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type APIBITMAPDATA
    Width               As Long
    Height              As Long
    Stride              As Long
    PixelFormat         As Long
    Scan0               As Long
    Reserved            As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_hBitmap               As Long
Private m_lWidth                As Single
Private m_lHeight               As Single
Private m_uHorParts()           As UcsPartInfoType
Private m_uVertParts()          As UcsPartInfoType
Private m_uPadding              As APIRECTL

Private Enum UcsPartTypeEnum
    ucsPrtStatic
    ucsPrtDynamic
    ucsPrtRepeat
End Enum

Private Type UcsPartInfoType
    Type                As UcsPartTypeEnum
    Position            As Long
    Size                As Long
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Function PrintError(sFunction As String) As VbMsgBoxResult
    Debug.Print STR_MODULE_NAME & "." & sFunction & ": " & Err.Description, Timer
End Function

Private Function RaiseError(sFunction As String) As VbMsgBoxResult
    Err.Raise Err.Number, STR_MODULE_NAME & "." & sFunction & vbCrLf & Err.Source, Err.Description
End Function

'=========================================================================
' Methods
'=========================================================================

Public Function LoadFromFile(sFileName As String) As Boolean
    Const FUNC_NAME     As String = "LoadFromFile"
    
    On Error GoTo EH
    Terminate
    If GdipLoadImageFromFile(StrPtr(sFileName), m_hBitmap) <> 0 Then
        GoTo QH
    End If
    '--- success (or failure)
    LoadFromFile = pvParseDataFrame()
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function LoadFromBitmap(ByVal hBitmap As Long) As Boolean
    Const FUNC_NAME     As String = "LoadFromBitmap"
    
    On Error GoTo EH
    Terminate
    m_hBitmap = hBitmap
    '--- success (or failure)
    LoadFromBitmap = pvParseDataFrame()
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function LoadFromByteArray(baData() As Byte) As Boolean
    Const FUNC_NAME     As String = "LoadFromByteArray"
    
    On Error GoTo EH
    Terminate
    If Not frBitmapFromByteArray(baData, m_hBitmap) Then
        GoTo QH
    End If
    '--- success (or failure)
    LoadFromByteArray = pvParseDataFrame()
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CalcBoundingBox( _
        ByVal ClientWidth As Long, _
        ByVal ClientHeight As Long, _
        Optional Width As Long, _
        Optional Height As Long, _
        Optional ClientX As Long, _
        Optional ClientY As Long) As Boolean
    Width = m_uPadding.Left + ClientWidth + m_uPadding.Right
    If Width < m_uHorParts(0).Size Then
        Width = m_uHorParts(0).Size
    End If
    Height = m_uPadding.Top + ClientHeight + m_uPadding.Bottom
    If Height < m_uVertParts(0).Size Then
        Height = m_uVertParts(0).Size
    End If
    ClientX = m_uPadding.Left
    ClientY = m_uPadding.Top
    '--- success
    CalcBoundingBox = True
End Function

Public Function CalcClientRect( _
        ByVal BoxWidth As Long, _
        ByVal BoxHeight As Long, _
        Optional X As Long, _
        Optional Y As Long, _
        Optional Width As Long, _
        Optional Height As Long) As Boolean
    X = m_uPadding.Left
    Y = m_uPadding.Top
    Width = BoxWidth - m_uPadding.Left - m_uPadding.Right
    If Width < 0 Then
        Width = 0
    End If
    Height = BoxHeight - m_uPadding.Top - m_uPadding.Bottom
    If Height < 0 Then
        Height = 0
    End If
End Function


Public Function DrawToDC(ByVal hDstDC As Long, ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Boolean
    Const FUNC_NAME     As String = "DrawToDC"
    Dim hGraphics       As Long
    
    On Error GoTo EH
    If GdipCreateFromHDC(hDstDC, hGraphics) <> 0 Then
        GoTo QH
    End If
    '-- success
    DrawToDC = DrawToGraphics(hGraphics, lX, lY, lWidth, lHeight)
QH:
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function DrawToGraphics(ByVal hGraphics As Long, ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Boolean
    Const FUNC_NAME     As String = "DrawToGraphics"
    Dim hState          As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim sngDynWidth     As Single
    Dim sngDynHeight    As Single
    Dim hMatrix         As Long
    Dim hHorMatrix      As Long
    Dim sngStepWidth    As Single
    Dim sngStepHeight   As Single
    Dim hBrush          As Long
    
    On Error GoTo EH
    If GdipBeginContainer2(hGraphics, hState) <> 0 Then
        GoTo QH
    End If
'    If GdipSetInterpolationMode(hGraphics, InterpolationModeHighQuality) <> 0 Then
'        GoTo QH
'    End If
    If GdipSetPixelOffsetMode(hGraphics, PixelOffsetModeHighQuality) <> 0 Then
        GoTo QH
    End If
    If GdipSetClipRectI(hGraphics, lX, lY, lWidth, lHeight, CombineModeIntersect) <> 0 Then
        GoTo QH
    End If
    If GdipCreateMatrix(hMatrix) <> 0 Then
        GoTo QH
    End If
    If GdipTranslateMatrix(hMatrix, lX, lY, MatrixOrderAppend) <> 0 Then
        GoTo QH
    End If
    If m_uHorParts(0).Position <> 0 And lWidth > m_uHorParts(0).Size Then
        sngDynWidth = (lWidth - m_uHorParts(0).Size) / m_uHorParts(0).Position
    End If
    If m_uVertParts(0).Position <> 0 And lHeight > m_uVertParts(0).Size Then
        sngDynHeight = (lHeight - m_uVertParts(0).Size) / m_uVertParts(0).Position
    End If
    For lIdx = 1 To UBound(m_uVertParts)
        sngStepHeight = IIf(m_uVertParts(lIdx).Type = ucsPrtStatic, m_uVertParts(lIdx).Size, sngDynHeight)
        If GdipCloneMatrix(hMatrix, hHorMatrix) <> 0 Then
            GoTo QH
        End If
        For lJdx = 1 To UBound(m_uHorParts)
            sngStepWidth = IIf(m_uHorParts(lJdx).Type = ucsPrtStatic, m_uHorParts(lJdx).Size, sngDynWidth)
            If GdipSetWorldTransform(hGraphics, hHorMatrix) <> 0 Then
                GoTo QH
            End If
            If m_uHorParts(lJdx).Type = ucsPrtRepeat Or m_uVertParts(lIdx).Type = ucsPrtRepeat Then
                If GdipCreateTexture2I(m_hBitmap, _
                        WrapModeTileFlipX * -(m_uHorParts(lJdx).Type <> ucsPrtRepeat) + WrapModeTileFlipY * -(m_uVertParts(lIdx).Type <> ucsPrtRepeat), _
                        m_uHorParts(lJdx).Position, m_uVertParts(lIdx).Position, _
                        m_uHorParts(lJdx).Size, m_uVertParts(lIdx).Size, hBrush) <> 0 Then
                    GoTo QH
                End If
                If GdipScaleTextureTransform(hBrush, _
                        IIf(m_uHorParts(lJdx).Type <> ucsPrtRepeat, sngStepWidth / m_uHorParts(lJdx).Size, 1), _
                        IIf(m_uVertParts(lIdx).Type <> ucsPrtRepeat, sngStepHeight / m_uVertParts(lIdx).Size, 1), _
                        MatrixOrderAppend) <> 0 Then
                    GoTo QH
                End If
                If GdipFillRectangleI(hGraphics, hBrush, 0, 0, sngStepWidth, sngStepHeight) <> 0 Then
                    GoTo QH
                End If
                Call GdipDeleteBrush(hBrush)
                hBrush = 0
            Else
                If GdipDrawImageRectRect(hGraphics, m_hBitmap, 0, 0, sngStepWidth, sngStepHeight, _
                        m_uHorParts(lJdx).Position, m_uVertParts(lIdx).Position, _
                        m_uHorParts(lJdx).Size, m_uVertParts(lIdx).Size) <> 0 Then
                    GoTo QH
                End If
            End If
            If GdipTranslateMatrix(hHorMatrix, sngStepWidth, 0, MatrixOrderAppend) <> 0 Then
                GoTo QH
            End If
        Next
        Call GdipDeleteMatrix(hHorMatrix)
        hHorMatrix = 0
        If GdipTranslateMatrix(hMatrix, 0, sngStepHeight, MatrixOrderAppend) <> 0 Then
            GoTo QH
        End If
    Next
    '-- success
    DrawToGraphics = True
QH:
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
        hBrush = 0
    End If
    If hHorMatrix <> 0 Then
        Call GdipDeleteMatrix(hHorMatrix)
        hHorMatrix = 0
    End If
    If hMatrix <> 0 Then
        Call GdipDeleteMatrix(hMatrix)
        hMatrix = 0
    End If
    If hState <> 0 Then
        Call GdipEndContainer(hGraphics, hState)
        hState = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Sub Terminate()
    If m_hBitmap <> 0 Then
        Call GdipDisposeImage(m_hBitmap)
        m_hBitmap = 0
    End If
End Sub

'= friend ================================================================

Friend Function frBitmapFromByteArray(baData() As Byte, hBitmap As Long) As Boolean
    Const FUNC_NAME     As String = "frBitmapFromByteArray"
    Dim hMem            As Long
    Dim lPtr            As Long
    Dim pStream         As IUnknown
    
    On Error GoTo EH
    hMem = GlobalAlloc(GMEM_MOVEABLE, UBound(baData) + 1)
    If hMem = 0 Then
        GoTo QH
    End If
    lPtr = GlobalLock(hMem)
    If lPtr = 0 Then
        GoTo QH
    End If
    Call CopyMemory(ByVal lPtr, baData(0), UBound(baData) + 1)
    Call GlobalUnlock(hMem)
    Call CreateStreamOnHGlobal(hMem, 1, pStream)
    If pStream Is Nothing Then
        GoTo QH
    End If
    hMem = 0 '--- stream will free hMem
    If GdipLoadImageFromStream(pStream, hBitmap) <> 0 Then
        GoTo QH
    End If
    '--- success
    frBitmapFromByteArray = True
QH:
    If hMem <> 0 Then
        Call GlobalFree(hMem)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'= private ===============================================================

Private Function pvParseDataFrame() As Boolean
    Const FUNC_NAME     As String = "pvParseDataFrame"
    Dim sngOrigWidth    As Single
    Dim sngOrigHeight   As Single
    Dim uRect           As APIRECTL
    Dim uData           As APIBITMAPDATA
    Dim clrStatic       As Long
    Dim clrRepeat       As Long
    Dim uPadParts()     As UcsPartInfoType
    
    On Error GoTo EH
    If GdipGetImageDimension(m_hBitmap, sngOrigWidth, sngOrigHeight) <> 0 Then
        GoTo QH
    End If
    m_lWidth = sngOrigWidth
    m_lHeight = sngOrigHeight
    m_uPadding = uRect
    uRect.Right = m_lWidth
    uRect.Bottom = m_lHeight
    If GdipBitmapLockBits(m_hBitmap, uRect, 1, PixelFormat32bppARGB, uData) <> 0 Then
        GoTo QH
    End If
    '--- top leftmost pixel
    clrStatic = Peek(uData.Scan0)
    '--- top rightmost pixel
    clrRepeat = Peek(UnsignedAdd(uData.Scan0, 4 * (m_lWidth - 1)))
    If Not pvGetParts(uData.Scan0, 4, m_lWidth - 2, clrStatic, clrRepeat, m_uHorParts) Then
        GoTo QH
    End If
    If Not pvGetParts(uData.Scan0, uData.Stride, m_lHeight - 2, clrStatic, clrRepeat, m_uVertParts) Then
        GoTo QH
    End If
    '--- figure out padding definition
    If pvGetParts(UnsignedAdd(uData.Scan0, uData.Stride * (m_lHeight - 1)), 4, m_lWidth - 2, clrStatic, clrRepeat, uPadParts) Then
        If UBound(uPadParts) = 3 Then
            If uPadParts(1).Type = ucsPrtStatic And uPadParts(2).Type <> ucsPrtStatic And uPadParts(3).Type = ucsPrtStatic Then
                 m_uPadding.Left = uPadParts(1).Size
                 m_uPadding.Right = uPadParts(3).Size
            End If
        End If
    End If
    If pvGetParts(UnsignedAdd(uData.Scan0, 4 * (m_lWidth - 1)), uData.Stride, m_lHeight - 2, clrStatic, clrRepeat, uPadParts) Then
        If UBound(uPadParts) = 3 Then
            If uPadParts(1).Type = ucsPrtStatic And uPadParts(2).Type <> ucsPrtStatic And uPadParts(3).Type = ucsPrtStatic Then
                 m_uPadding.Top = uPadParts(1).Size
                 m_uPadding.Bottom = uPadParts(3).Size
            End If
        End If
    End If
    Call GdipBitmapUnlockBits(m_hBitmap, uData)
    uData.Scan0 = 0
    '--- success (or failure)
    pvParseDataFrame = pvClearDataFrame(clrStatic)
    Exit Function
QH:
    If uData.Scan0 <> 0 Then
        Call GdipBitmapUnlockBits(m_hBitmap, uData)
        uData.Scan0 = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvClearDataFrame(ByVal clrStatic As Long) As Boolean
    Const FUNC_NAME     As String = "pvClearDataFrame"
    Dim hGraphics       As Long
    Dim hBrush          As Long
    
    On Error GoTo EH
    If GdipGetImageGraphicsContext(m_hBitmap, hGraphics) <> 0 Then
        GoTo QH
    End If
    If GdipCreateSolidFill(clrStatic, hBrush) <> 0 Then
        GoTo QH
    End If
    If GdipSetCompositingMode(hGraphics, CompositingModeSourceCopy) <> 0 Then
        GoTo QH
    End If
    Call GdipFillRectangleI(hGraphics, hBrush, 0, 0, m_lWidth, 1)
    Call GdipFillRectangleI(hGraphics, hBrush, 0, m_lHeight - 1, m_lWidth, 1)
    Call GdipFillRectangleI(hGraphics, hBrush, 0, 0, 1, m_lHeight)
    Call GdipFillRectangleI(hGraphics, hBrush, m_lWidth - 1, 0, 1, m_lHeight)
    '--- success
    pvClearDataFrame = True
QH:
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
        hBrush = 0
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

Private Function pvGetParts( _
            ByVal lPtr As Long, _
            ByVal lStep As Long, _
            ByVal lSize As Long, _
            ByVal clrStatic As Long, _
            ByVal clrRepeat As Long, _
            uRetVal() As UcsPartInfoType) As Boolean
    Dim clrNext         As Long
    Dim eNextType       As UcsPartTypeEnum
    Dim eType           As UcsPartTypeEnum
    Dim lPos            As Long
    Dim lIdx            As Long
    
    ReDim uRetVal(0 To 0) As UcsPartInfoType
    lPtr = UnsignedAdd(lPtr, lStep)
    clrNext = Peek(lPtr)
    eNextType = Switch(clrNext = clrStatic, ucsPrtStatic, clrNext = clrRepeat, ucsPrtRepeat, True, ucsPrtDynamic)
    eType = eNextType
    For lIdx = 1 To lSize - 1
        lPtr = UnsignedAdd(lPtr, lStep)
        clrNext = Peek(lPtr)
        eNextType = Switch(clrNext = clrStatic, ucsPrtStatic, clrNext = clrRepeat, ucsPrtRepeat, True, ucsPrtDynamic)
        If eType <> eNextType Then
            GoSub AppendPart
            lPos = lIdx
            eType = eNextType
        End If
    Next
    GoSub AppendPart
    '--- success
    pvGetParts = True
    Exit Function
AppendPart:
    ReDim Preserve uRetVal(0 To UBound(uRetVal) + 1) As UcsPartInfoType
    With uRetVal(UBound(uRetVal))
        .Type = eType
        .Position = lPos + 1
        .Size = lIdx - lPos
    End With
    If eType = ucsPrtStatic Then
        uRetVal(0).Size = uRetVal(0).Size + lIdx - lPos
    Else
        uRetVal(0).Position = uRetVal(0).Position + 1
    End If
    Return
End Function

#If Not ImplUseShared Then
Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function
#End If

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Dim uStartup        As GdiplusStartupInput
    
    If GetModuleHandle("gdiplus") = 0 Then
        uStartup.GdiplusVersion = 1&
        Call GdiplusStartup(0, uStartup)
    End If
    ReDim m_uHorParts(0 To 0) As UcsPartInfoType
    ReDim m_uVertParts(0 To 0) As UcsPartInfoType
End Sub

Private Sub Class_Terminate()
    Terminate
End Sub
