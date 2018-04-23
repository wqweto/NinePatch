VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDibSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cDibSection"

#Const SelfContained = Not DebugMode

'=========================================================================
' API
'=========================================================================

Private Const DIB_RGB_COLORS                As Long = 0
Private Const GMEM_MOVEABLE                 As Long = 2
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
'--- for GdipCreateBitmapFromScan0
Private Const PixelFormat32bppRGB           As Long = &H22009
Private Const PixelFormat32bppARGB          As Long = &H26200A
Private Const PixelFormatAlpha              As Long = &H40000
'--- for GdipSetInterpolationMode
Private Const InterpolationModeHighQualityBicubic As Long = 7
'--- for SetStretchBltMode
Private Const HALFTONE                      As Long = 4
'--- DrawIconEx constants
Private Const DI_NORMAL                     As Long = &H3

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ApiBitBlt Lib "gdi32" Alias "BitBlt" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ApiStretchBlt Lib "gdi32" Alias "StretchBlt" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function FindResourceLong Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As Long, ByVal lpType As String) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, ByRef mImage As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal mGraphics As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal img As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef nWidth As Single, ByRef nHeight As Single) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal Image As Long, nFormat As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal stride As Long, ByVal Format As Long, ByVal scan0 As Long, ByRef Bitmap As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipCreateEffect Lib "gdiplus" (ByVal dwCid1 As Long, ByVal dwCid2 As Long, ByVal dwCid3 As Long, ByVal dwCid4 As Long, ByRef Effect As Long) As Long
Private Declare Function GdipSetEffectParameters Lib "gdiplus" (ByVal Effect As Long, ByRef Parameters As Any, ByVal Size As Long) As Long
Private Declare Function GdipDeleteEffect Lib "gdiplus" (ByVal Effect As Long) As Long
Private Declare Function GdipBitmapApplyEffect Lib "gdiplus" (ByVal hBitmap As Long, ByVal hEffect As Long, roi As Any, Optional ByVal useAuxData As Long, Optional ByVal auxData As Long, Optional ByVal auxDataSize As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imgAttr As Long, ByVal clrAdjust As Long, ByVal clrAdjustEnabled As Long, ByRef clrMatrix As Any, ByRef grayMatrix As Any, ByVal clrMatrixFlags As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imgAttr As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal Callback As Long, ByVal callbackData As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbByte, Optional ByVal Low As Long, Optional ByVal Count As Long) As Byte()
Private Declare Function CLSIDFromString Lib "ole32" (ByVal szPtr As Long, clsid As GdipGUID) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
#If SelfContained Then
    Private Declare Function GetMem4 Lib "msvbvm60" (src As Any, dst As Any) As Long
    Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
    Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal OutputBuf As Long) As Long
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
#End If

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

Private Type RGBQUAD
    B                   As Byte
    G                   As Byte
    R                   As Byte
    A                   As Byte
End Type

Private Type SAFEARRAYBOUND
    cElements           As Long
    lLbound             As Long
End Type

Private Type SAFEARRAY2D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As Long
    Bounds(0 To 1)      As SAFEARRAYBOUND
End Type

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type PICTDESC
    lSize               As Long
    lType               As Long
    hBmp                As Long
    hPal                As Long
End Type

Private Type ICONINFO
    fIcon               As Long
    xHotspot            As Long
    yHotspot            As Long
    hbmMask             As Long
    hbmColor            As Long
End Type

Private Type GdipGUID
    Data1               As Long
    Data2               As Long
    Data3               As Long
    Data4               As Long
End Type

Private Type GdipEncoderParameter
    GUID                As GdipGUID
    NumberOfValues      As Long
    Type                As Long
    Value               As Long
End Type

Private Type GdipEncoderParameters
    Count               As Long
    Parameters(0 To 5) As GdipEncoderParameter
End Type

Private Type GdipImageCodecInfo
    GUID                As GdipGUID
    FormatID(0 To 3)    As Long
    CodecName           As Long
    DllName             As Long
    FormatDescription   As Long
    FilenameExtension   As Long
    MimeType            As Long
    Flags               As Long
    Version             As Long
    SigCount            As Long
    SigSize             As Long
    SigPattern          As Long
    SigMask             As Long
End Type

Private Type GdipBlurParams
    Radius              As Single
    ExpandEdge          As Long
End Type

#If SelfContained Then
    Private Type GdiplusStartupInput
        GdiplusVersion           As Long
        DebugEventCallback       As Long
        SuppressBackgroundThread As Long
        SuppressExternalCodecs   As Long
    End Type
#End If

'=========================================================================
' Constants and variables
'=========================================================================

Private Const ERR_RESOURCE_NOT_FOUND As Long = 1814

Private m_uBmpHeader            As BITMAPINFOHEADER
Private m_hDC                   As Long
Private m_hDib                  As Long
Private m_hPrevDib              As Long
Private m_lpBits                As Long
Private m_uArray                As SAFEARRAY2D
Private m_aBitsRGB()            As RGBQUAD
Private m_aBits()               As Long
Private m_sngOrigWidth          As Single
Private m_sngOrigHeight         As Single
Private m_lPixelFormat          As Long
#If SelfContained Then
    Private g_hGdiPlus          As Long
#End If

'=========================================================================
' Properties
'=========================================================================

Public Property Get hDC() As Long
    hDC = m_hDC
End Property

Public Property Get hDib() As Long
    hDib = m_hDib
End Property

Public Property Let hDib(ByVal lValue As Long)
    m_hDib = lValue
End Property

Public Property Get DIBitsPtr() As Long
    DIBitsPtr = m_lpBits
End Property

Public Property Get Size() As Long
    Size = m_uBmpHeader.biSizeImage
End Property

Public Property Get BytesPerScanLine() As Long
    BytesPerScanLine = 4 * m_uBmpHeader.biWidth
End Property

Public Property Get Width() As Long
    Width = m_uBmpHeader.biWidth
End Property

Public Property Get Height() As Long
    Height = -m_uBmpHeader.biHeight
End Property

Public Property Get OrigWidth() As Long
    OrigWidth = m_sngOrigWidth
End Property

Public Property Get OrigHeight() As Long
    OrigHeight = m_sngOrigHeight
End Property

Public Property Get HasAlpha() As Boolean
    HasAlpha = (m_lPixelFormat And PixelFormatAlpha) <> 0
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init( _
            ByVal NewWidth As Long, _
            ByVal NewHeight As Long, _
            Optional ByVal BackColor As OLE_COLOR = vbWhite, _
            Optional ByVal StretchMode As Long = HALFTONE) As Boolean
    '-- Destroy previous
    Terminate
    
    If (4# * NewWidth) * NewHeight > 2# ^ 31 Then
        Exit Function
    End If
    
    '-- Prepare header
    With m_uBmpHeader
        .biSize = Len(m_uBmpHeader)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = NewWidth
        .biHeight = -NewHeight
        .biSizeImage = (4 * NewWidth) * NewHeight
    End With
    m_lPixelFormat = PixelFormat32bppRGB
    
    '-- Create DIB section
    m_hDC = CreateCompatibleDC(0)
    If m_hDC <> 0 Then
        If StretchMode > 0 Then
            Call SetStretchBltMode(m_hDC, StretchMode)
        End If
        '-- Create DIB
        m_hDib = CreateDIBSection(m_hDC, m_uBmpHeader, DIB_RGB_COLORS, m_lpBits, 0, 0)
        If m_hDib <> 0 Then
            '-- Select into a DC
            m_hPrevDib = SelectObject(m_hDC, m_hDib)
            '-- Set backColor
            Cls BackColor
            '--- create safe array for RGB
            With m_uArray
                .cbElements = 4
                .cDims = 2
                .Bounds(0).lLbound = 0
                .Bounds(0).cElements = Height
                .Bounds(1).lLbound = 0
                .Bounds(1).cElements = Width
                .pvData = DIBitsPtr
            End With
            Call CopyMemory(ByVal ArrPtr(m_aBitsRGB()), VarPtr(m_uArray), 4)
            Call CopyMemory(ByVal ArrPtr(m_aBits()), VarPtr(m_uArray), 4)
        Else
            Terminate
        End If
    End If
    
    '-- Success
    Init = (m_hDib <> 0)
End Function

Public Sub Terminate()
    If m_hDC <> 0 Then
        If m_hPrevDib <> 0 Then
            Call SelectObject(m_hDC, m_hPrevDib)
            m_hPrevDib = 0
        End If
        If m_hDib <> 0 Then
            Call DeleteObject(m_hDib)
            m_hDib = 0
            m_lpBits = 0
        End If
        Call DeleteDC(m_hDC)
        m_hDC = 0
    End If
    Call CopyMemory(ByVal ArrPtr(m_aBitsRGB()), 0&, 4)
    Call CopyMemory(ByVal ArrPtr(m_aBits()), 0&, 4)
End Sub

Public Sub Cls( _
            Optional ByVal BackColor As OLE_COLOR = vbBlack, _
            Optional ByVal xDst As Long, _
            Optional ByVal yDst As Long, _
            Optional ByVal cxDst As Long, _
            Optional ByVal cyDst As Long)
    Dim uRect         As RECT
    Dim hBrush        As Long
    
    On Error GoTo EH
    If m_hDib <> 0 Then
        uRect.Left = xDst
        uRect.Top = yDst
        uRect.Right = xDst + IIf(cxDst <> 0, cxDst, m_uBmpHeader.biWidth)
        uRect.Bottom = yDst + IIf(cyDst <> 0, cyDst, -m_uBmpHeader.biHeight)
        hBrush = CreateSolidBrush(TranslateColor(BackColor))
        Call FillRect(m_hDC, uRect, hBrush)
        Call DeleteObject(hBrush)
        hBrush = 0
    End If
    Exit Sub
EH:
    If hBrush <> 0 Then
        Call DeleteObject(hBrush)
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub FillCheckered( _
            Optional ByVal xDst As Long, _
            Optional ByVal yDst As Long, _
            Optional ByVal cxDst As Long, _
            Optional ByVal cyDst As Long, _
            Optional ByVal clrLight As OLE_COLOR = vbWhite, _
            Optional ByVal clrDark As OLE_COLOR = &HBFBFBF, _
            Optional ByVal StepInPx As Long = 8)
    Dim lX              As Long
    Dim lY              As Long
    Dim bRow            As Boolean
    Dim bEven           As Boolean
    Dim uRect         As RECT
    Dim hBrush        As Long
    
    On Error GoTo EH
    If StepInPx <= 0 Then
        StepInPx = 1
    End If
    If cxDst = 0 Then
        cxDst = m_uBmpHeader.biWidth
    End If
    If cyDst = 0 Then
        cyDst = -m_uBmpHeader.biHeight
    End If
    Cls clrLight, xDst, yDst, cxDst, cyDst
    hBrush = CreateSolidBrush(TranslateColor(clrDark))
    For lY = 0 To cyDst Step StepInPx
        bEven = bRow
        For lX = 0 To cxDst Step StepInPx
            If Not bEven Then
                uRect.Left = xDst + lX
                uRect.Top = xDst + lY
                uRect.Right = uRect.Left + LimitLong(StepInPx, Max:=cxDst - lX)
                uRect.Bottom = uRect.Top + LimitLong(StepInPx, Max:=cyDst - lY)
                Call FillRect(m_hDC, uRect, hBrush)
            End If
            bEven = Not bEven
        Next
        bRow = Not bRow
    Next
    Call DeleteObject(hBrush)
    hBrush = 0
    Exit Sub
EH:
    If hBrush <> 0 Then
        Call DeleteObject(hBrush)
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub LoadBlt( _
            ByVal hSrcDC As Long, _
            Optional ByVal xDst As Long, _
            Optional ByVal yDst As Long, _
            Optional ByVal cxDst As Long, _
            Optional ByVal cyDst As Long, _
            Optional ByVal xSrc As Long, _
            Optional ByVal ySrc As Long, _
            Optional ByVal cxSrc As Long, _
            Optional ByVal cySrc As Long)
    If m_hDib <> 0 Then
        If cxDst = 0 Then
            cxDst = m_uBmpHeader.biWidth
        End If
        If cyDst = 0 Then
            cyDst = -m_uBmpHeader.biHeight
        End If
        If cxSrc = 0 Then
            cxSrc = m_uBmpHeader.biWidth
        End If
        If cySrc = 0 Then
            cySrc = -m_uBmpHeader.biHeight
        End If
        Call ApiStretchBlt(m_hDC, xDst, yDst, cxDst, cyDst, hSrcDC, xSrc, ySrc, cxSrc, cySrc, vbSrcCopy)
    End If
End Sub

Public Function LoadFromPicture( _
            oPic As StdPicture, _
            Optional BackColor As OLE_COLOR = vbWhite, _
            Optional ByVal Width As Single, _
            Optional ByVal Height As Single, _
            Optional ByVal KeepAspect As Boolean) As Boolean
    Dim uInfo           As ICONINFO
    Dim hMaskDC         As Long
    Dim hPrevMaskBmp    As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    If oPic Is Nothing Then
        Exit Function
    End If
    If oPic.Handle = 0 Then
        Exit Function
    End If
    m_sngOrigWidth = HM2Pix(oPic.Width)
    m_sngOrigHeight = HM2Pix(oPic.Height)
    pvCalcDimensions Width, Height, KeepAspect
    If Not Init(Width, Height, BackColor:=BackColor) Then
        GoTo QH
    End If
    RenderPicture oPic, m_hDC, 0, 0, Me.Width, Me.Height, 0, oPic.Height, oPic.Width, -oPic.Height
    Select Case oPic.Type
    Case vbPicTypeIcon
        m_lPixelFormat = PixelFormat32bppARGB
        If Not pvHasAlpha() Then
            Call GetIconInfo(oPic.Handle, uInfo)
            If uInfo.hbmMask <> 0 Then
                hMaskDC = CreateCompatibleDC(m_hDC)
                hPrevMaskBmp = SelectObject(hMaskDC, uInfo.hbmMask)
                With New cDibSection
                    .Init Width, Height
                    '--- note: uInfo.hbmColor=0 -> monochrome icon -> hbmMask has double height
                    .LoadBlt hMaskDC, ySrc:=IIf(uInfo.hbmColor = 0, HM2Pix(oPic.Height), 0), cxSrc:=HM2Pix(oPic.Width), cySrc:=HM2Pix(oPic.Height)
                    For lJdx = 0 To Height - 1
                        For lIdx = 0 To Width - 1
                            If .Pixel(lIdx, lJdx) = 0 Then
                                m_aBits(lIdx, lJdx) = m_aBits(lIdx, lJdx) Or &HFF000000
                            Else
                                m_aBits(lIdx, lJdx) = vbMagenta Or &H1000000    ' vbMagenta ' Or &HFF000000 ' m_aBits(lIdx, lJdx) And &HFFFFFF
                            End If
                        Next
                    Next
                End With
                Call SelectObject(hMaskDC, hPrevMaskBmp)
                Call DeleteDC(hMaskDC)
            End If
            If uInfo.hbmColor <> 0 Then
                Call DeleteObject(uInfo.hbmColor)
            End If
            If uInfo.hbmMask <> 0 Then
                Call DeleteObject(uInfo.hbmMask)
            End If
        End If
    Case vbPicTypeMetafile, vbPicTypeEMetafile
        m_lPixelFormat = IIf(pvHasAlpha(), PixelFormat32bppARGB, PixelFormat32bppRGB)
    Case Else
        m_lPixelFormat = PixelFormat32bppRGB
    End Select
    '--- success
    LoadFromPicture = True
QH:
End Function

Private Function pvHasAlpha() As Boolean
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lAlpha          As Long
    
    For lJdx = 0 To Height - 1
        For lIdx = 0 To Width - 1
            lAlpha = (m_aBits(lIdx, lJdx) And &HFF000000)
            If lAlpha <> 0 And lAlpha <> &HFF000000 Then
                pvHasAlpha = True
                Exit Function
            End If
        Next
    Next
End Function

Public Function LoadFromFile( _
            sFileName As String, _
            Optional ByVal BackColor As OLE_COLOR = vbWhite, _
            Optional ByVal Width As Single, _
            Optional ByVal Height As Single, _
            Optional ByVal KeepAspect As Boolean) As Boolean
    Const FUNC_NAME     As String = "LoadFromFile"
    Dim hGraphics       As Long
    Dim hBitmap         As Long
    Dim lGdipStatus     As Long
    Dim lLastDllError   As Long
    Dim sSource         As String
    
    On Error GoTo EH
    If Not StartGdip() Then
        Exit Function
    End If
    lGdipStatus = GdipLoadImageFromFile(StrPtr(sFileName), hBitmap)
    If lGdipStatus <> 0 Then
        Exit Function
    End If
    lGdipStatus = GdipGetImageDimension(hBitmap, m_sngOrigWidth, m_sngOrigHeight)
    If lGdipStatus <> 0 Then
        sSource = "GdipGetImageDimension"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    pvCalcDimensions Width, Height, KeepAspect
    If Not Init(Width, Height, BackColor:=BackColor) Then
        GoTo QH
    End If
    lGdipStatus = GdipGetImagePixelFormat(hBitmap, m_lPixelFormat)
    If lGdipStatus <> 0 Then
        sSource = "GdipGetImagePixelFormat"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    lGdipStatus = GdipCreateFromHDC(hDC, hGraphics)
    If lGdipStatus <> 0 Then
        sSource = "GdipCreateFromHDC"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    lGdipStatus = GdipSetInterpolationMode(hGraphics, InterpolationModeHighQualityBicubic)
    If lGdipStatus <> 0 Then
        sSource = "GdipSetInterpolationMode"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    lGdipStatus = GdipDrawImageRectI(hGraphics, hBitmap, 0, 0, Me.Width, Me.Height)
    If lGdipStatus <> 0 Then
        sSource = "GdipDrawImageRectI"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    '--- success
    LoadFromFile = True
QH:
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    If lGdipStatus <> 0 Then
        Err.Raise vbObjectError, MODULE_NAME & "." & FUNC_NAME & "(sSource=" & sSource & ", sFileName=" & sFileName & ")", GdipGetSystemMessage(lGdipStatus, lLastDllError)
    ElseIf lLastDllError <> 0 Then
        Err.Raise vbObjectError, MODULE_NAME & "." & FUNC_NAME & "(sSource=" & sSource & ", sFileName=" & sFileName & ")", GetSystemMessage(lLastDllError)
    End If
    Exit Function
EH:
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function LoadFromResource( _
            ByVal hRll As Long, _
            sResName As String, _
            Optional ResType As String = "PNG", _
            Optional ByVal BackColor As OLE_COLOR = vbWhite, _
            Optional ByVal Width As Single, _
            Optional ByVal Height As Single, _
            Optional ByVal KeepAspect As Boolean) As Boolean
    Const FUNC_NAME     As String = "LoadFromResource"
    Dim hResInfo        As Long
    Dim hGlobal         As Long
    Dim hMem            As Long
    Dim lPtr            As Long
    Dim pStream         As IUnknown
    Dim lGdipStatus     As Long
    Dim lLastDllError   As Long
    Dim sSource         As String
    
    On Error GoTo EH
    If C_Lng(sResName) <> 0 Then
        hResInfo = FindResourceLong(hRll, C_Lng(sResName), ResType)
    End If
    If hResInfo = 0 Then
        hResInfo = FindResource(hRll, sResName, ResType)
    End If
    If hResInfo = 0 Then
        If Err.LastDllError = ERR_RESOURCE_NOT_FOUND Then
            Exit Function
        End If
        sSource = "FindResource"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    hGlobal = LoadResource(hRll, hResInfo)
    If hGlobal = 0 Then
        sSource = "LoadResource"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    hMem = GlobalAlloc(GMEM_MOVEABLE, SizeofResource(hRll, hResInfo))
    If hMem = 0 Then
        sSource = "GlobalAlloc"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    lPtr = GlobalLock(hMem)
    If lPtr = 0 Then
        sSource = "GlobalLock"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    Call CopyMemory(ByVal lPtr, ByVal LockResource(hGlobal), SizeofResource(hRll, hResInfo))
    Call GlobalUnlock(hMem)
    Set pStream = CreateStreamOnHGlobal(hMem, 1)
    If pStream Is Nothing Then
        sSource = "CreateStreamOnHGlobal"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    hMem = 0 '--- stream will free hMem
    If pvLoadFromStream(pStream, BackColor, Width, Height, KeepAspect, lGdipStatus, lLastDllError) Then
        '--- success
        LoadFromResource = True
    Else
        sSource = "pvLoadFromStream"
    End If
QH:
    If hMem <> 0 Then
        Call GlobalFree(hMem)
    End If
    If hGlobal <> 0 Then
        Call FreeResource(hGlobal)
    End If
    If lGdipStatus <> 0 Then
        Err.Raise vbObjectError, MODULE_NAME & "." & FUNC_NAME & "(sSource=" & sSource & ", sResName=" & sResName & ")", GdipGetSystemMessage(lGdipStatus, lLastDllError)
    ElseIf lLastDllError <> 0 Then
        Err.Raise vbObjectError, MODULE_NAME & "." & FUNC_NAME & "(sSource=" & sSource & ", sResName=" & sResName & ")", GetSystemMessage(lLastDllError)
    End If
    Exit Function
EH:
    If hMem <> 0 Then
        Call GlobalFree(hMem)
    End If
    If hGlobal <> 0 Then
        Call FreeResource(hGlobal)
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function LoadFromByteArray( _
            baData() As Byte, _
            Optional ByVal BackColor As OLE_COLOR = vbWhite, _
            Optional ByVal Width As Single, _
            Optional ByVal Height As Single, _
            Optional ByVal KeepAspect As Boolean) As Boolean
    Const FUNC_NAME     As String = "LoadFromByteArray"
    Dim hMem            As Long
    Dim lPtr            As Long
    Dim pStream         As IUnknown
    Dim lGdipStatus     As Long
    Dim lLastDllError   As Long
    Dim sSource         As String
    
    On Error GoTo EH
    '--- sanity check
    If Peek(ArrPtr(baData)) = 0 Then
        Exit Function
    End If
    If UBound(baData) < 0 Then
        Exit Function
    End If
    hMem = GlobalAlloc(GMEM_MOVEABLE, UBound(baData) + 1)
    If hMem = 0 Then
        sSource = "GlobalAlloc"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    lPtr = GlobalLock(hMem)
    If lPtr = 0 Then
        sSource = "GlobalLock"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    Call CopyMemory(ByVal lPtr, baData(0), UBound(baData) + 1)
    Call GlobalUnlock(hMem)
    Set pStream = CreateStreamOnHGlobal(hMem, 1)
    If pStream Is Nothing Then
        sSource = "CreateStreamOnHGlobal"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    hMem = 0 '--- stream will free hMem
    If pvLoadFromStream(pStream, BackColor, Width, Height, KeepAspect, lGdipStatus, lLastDllError) Then
        '--- success
        LoadFromByteArray = True
    Else
        sSource = "pvLoadFromStream"
    End If
QH:
    If hMem <> 0 Then
        Call GlobalFree(hMem)
    End If
    If lGdipStatus <> 0 Then
        Err.Raise vbObjectError, MODULE_NAME & "." & FUNC_NAME & "(sSource=" & sSource & ")", GdipGetSystemMessage(lGdipStatus, lLastDllError)
    ElseIf lLastDllError <> 0 Then
        Err.Raise vbObjectError, MODULE_NAME & "." & FUNC_NAME & "(sSource=" & sSource & ")", GetSystemMessage(lLastDllError)
    End If
    Exit Function
EH:
    If hMem <> 0 Then
        Call GlobalFree(hMem)
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function pvLoadFromStream( _
            pStream As IUnknown, _
            BackColor As OLE_COLOR, _
            ByVal sngWidth As Single, _
            ByVal sngHeight As Single, _
            ByVal bKeepAspect As Boolean, _
            lGdipStatus As Long, _
            lLastDllError As Long) As Boolean
    Const FUNC_NAME     As String = "pvLoadFromStream"
    Dim hBitmap         As Long
    Dim hGraphics       As Long
    Dim sSource         As String
    Dim lFormat         As Long
    
    On Error GoTo EH
    If Not StartGdip() Then
        GoTo QH
    End If
    If GdipLoadImageFromStream(pStream, hBitmap) <> 0 Then
        '--- don't raise error
        GoTo QH
    End If
    lGdipStatus = GdipGetImageDimension(hBitmap, m_sngOrigWidth, m_sngOrigHeight)
    If lGdipStatus <> 0 Then
        sSource = "GdipGetImageDimension"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    lGdipStatus = GdipGetImagePixelFormat(hBitmap, m_lPixelFormat)
    If lGdipStatus <> 0 Then
        sSource = "GdipGetImagePixelFormat"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    If sngWidth < 0 Or sngHeight < 0 Then
        Exit Function
    End If
    pvCalcDimensions sngWidth, sngHeight, bKeepAspect
    lFormat = m_lPixelFormat
    If Not Init(sngWidth, sngHeight, BackColor:=BackColor) Then
        GoTo QH
    End If
    m_lPixelFormat = lFormat
    lGdipStatus = GdipCreateFromHDC(hDC, hGraphics)
    If lGdipStatus <> 0 Then
        sSource = "GdipCreateFromHDC"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    lGdipStatus = GdipSetInterpolationMode(hGraphics, InterpolationModeHighQualityBicubic)
    If lGdipStatus <> 0 Then
        sSource = "GdipSetInterpolationMode"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    lGdipStatus = GdipDrawImageRectI(hGraphics, hBitmap, 0, 0, Width, Height)
    If lGdipStatus <> 0 Then
        sSource = "GdipDrawImageRectI"
        lLastDllError = Err.LastDllError
        GoTo QH
    End If
    '--- success
    pvLoadFromStream = True
QH:
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    If lGdipStatus <> 0 Then
        Err.Raise vbObjectError, MODULE_NAME & "." & FUNC_NAME & "(sSource=" & sSource & ")", GdipGetSystemMessage(lGdipStatus, lLastDllError)
    ElseIf lLastDllError <> 0 Then
        Err.Raise vbObjectError, MODULE_NAME & "." & FUNC_NAME & "(sSource=" & sSource & ")", GetSystemMessage(lLastDllError)
    End If
    Exit Function
EH:
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SaveToByteArray(Optional MimeType As String = "image/jpeg", Optional ByVal Quality As Long = 50) As Byte()
    Const EncoderColorDepth As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
    Const EncoderQuality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
    Const EncoderParameterValueTypeLong As Long = 4
    Dim hBitmap         As Long
    Dim pStream         As stdole.IUnknown
    Dim uEncoder        As GdipGUID
    Dim uParams         As GdipEncoderParameters
    Dim hMem            As Long
    Dim lPtr            As Long
    Dim baRetVal()      As Byte
    Dim lColorDepth     As Long
    
    On Error GoTo EH
    SaveToByteArray = ApiEmptyByteArray()
    If Not StartGdip() Then
        GoTo QH
    End If
    If LenB(MimeType) = 0 Then
        MimeType = "image/jpeg"
    End If
    uEncoder = pvGetEncoderClsid(MimeType)
    If uEncoder.Data1 = 0 Then
        GoTo QH
    End If
    Set pStream = CreateStreamOnHGlobal(0, 1)
    If pStream Is Nothing Then
        GoTo QH
    End If
    If GdipCreateBitmapFromScan0(Width, Height, BytesPerScanLine, IIf(Not HasAlpha, PixelFormat32bppRGB, PixelFormat32bppARGB), DIBitsPtr, hBitmap) <> 0 Then
        GoTo QH
    End If
    With uParams.Parameters(uParams.Count)
        Call CLSIDFromString(StrPtr(EncoderColorDepth), .GUID)
        .NumberOfValues = 1
        .Type = EncoderParameterValueTypeLong
        lColorDepth = IIf(HasAlpha, 32, 24)
        .Value = VarPtr(lColorDepth)
    End With
    uParams.Count = uParams.Count + 1
    If MimeType = "image/jpeg" Then
        With uParams.Parameters(uParams.Count)
            Call CLSIDFromString(StrPtr(EncoderQuality), .GUID)
            .NumberOfValues = 1
            .Type = EncoderParameterValueTypeLong
            .Value = VarPtr(Quality)
        End With
        uParams.Count = uParams.Count + 1
    End If
    If GdipSaveImageToStream(hBitmap, pStream, uEncoder, uParams) <> 0 Then
        GoTo QH
    End If
    '--- hMem destroyed by stream
    If GetHGlobalFromStream(ObjPtr(pStream), hMem) <> 0 Then
        GoTo QH
    End If
    lPtr = GlobalLock(hMem)
    If lPtr = 0 Or GlobalSize(hMem) = 0 Then
        GoTo QH
    End If
    ReDim baRetVal(0 To GlobalSize(hMem) - 1) As Byte
    Call CopyMemory(baRetVal(0), ByVal lPtr, UBound(baRetVal) + 1)
    If GlobalUnlock(hMem) <> 0 Then
        GoTo QH
    End If
    '--- suuccess
    SaveToByteArray = baRetVal
QH:
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    Exit Function
EH:
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub BitBlt( _
            ByVal hDstDC As Long, _
            Optional ByVal xDst As Long, _
            Optional ByVal yDst As Long, _
            Optional ByVal cxDst As Long, _
            Optional ByVal cyDst As Long, _
            Optional ByVal xSrc As Long, _
            Optional ByVal ySrc As Long, _
            Optional ByVal OpacityPrc As Long)
    Const UnitPixel     As Long = 2&
    Dim hBitmap         As Long
    Dim hGraphics       As Long
    Dim hAttributes     As Long
    Dim clrMatrix(0 To 4, 0 To 4) As Single

    If m_hDib <> 0 Then
        If cxDst = 0 Then
            cxDst = m_uBmpHeader.biWidth
        End If
        If cyDst = 0 Then
            cyDst = -m_uBmpHeader.biHeight
        End If
        If OpacityPrc <> 0 Then
            If GdipCreateBitmapFromScan0(Width, Height, BytesPerScanLine, IIf(Not HasAlpha, PixelFormat32bppRGB, PixelFormat32bppARGB), DIBitsPtr, hBitmap) <> 0 Then
                GoTo QH
            End If
            If GdipCreateFromHDC(hDstDC, hGraphics) <> 0 Then
                GoTo QH
            End If
            If GdipCreateImageAttributes(hAttributes) <> 0 Then
                GoTo QH
            End If
            clrMatrix(0, 0) = IIf(OpacityPrc < 0, -OpacityPrc / 100#, 1)
            clrMatrix(1, 1) = IIf(OpacityPrc < 0, -OpacityPrc / 100#, 1)
            clrMatrix(2, 2) = IIf(OpacityPrc < 0, -OpacityPrc / 100#, 1)
            clrMatrix(3, 3) = IIf(OpacityPrc > 0, OpacityPrc / 100#, 1)
            clrMatrix(0, 4) = IIf(OpacityPrc < 0, OpacityPrc / 100#, 0)
            clrMatrix(1, 4) = IIf(OpacityPrc < 0, OpacityPrc / 100#, 0)
            clrMatrix(2, 4) = IIf(OpacityPrc < 0, OpacityPrc / 100#, 0)
            clrMatrix(4, 4) = 1
            If GdipSetImageAttributesColorMatrix(hAttributes, 0, 1, clrMatrix(0, 0), clrMatrix(0, 0), 0) <> 0 Then '
                GoTo QH
            End If
            If GdipDrawImageRectRectI(hGraphics, hBitmap, xDst, yDst, cxDst, cyDst, xSrc, ySrc, cxDst, cyDst, UnitPixel, hAttributes, 0, 0) <> 0 Then
                GoTo QH
            End If
QH:
            If hBitmap <> 0 Then
                Call GdipDisposeImage(hBitmap)
            End If
            If hGraphics <> 0 Then
                Call GdipDeleteGraphics(hGraphics)
            End If
            If hAttributes <> 0 Then
                Call GdipDisposeImageAttributes(hAttributes)
            End If
        Else
            Call ApiBitBlt(hDstDC, xDst, yDst, cxDst, cyDst, m_hDC, xSrc, ySrc, vbSrcCopy)
        End If
    End If
End Sub

Property Get Pixel(ByVal X As Long, ByVal Y As Long) As Long
    Pixel = (m_aBits(X, Y) And &HFFFFFF)
End Property

Property Let Pixel(ByVal X As Long, ByVal Y As Long, ByVal clrValue As Long)
    m_aBits(X, Y) = (clrValue And &HFFFFFF) Or (m_aBits(X, Y) And Not &HFFFFFF)
End Property

Property Get Alpha(ByVal X As Long, ByVal Y As Long) As Long
    Alpha = m_aBitsRGB(X, Y).A
End Property

Property Let Alpha(ByVal X As Long, ByVal Y As Long, ByVal lValue As Long)
    m_aBitsRGB(X, Y).A = lValue
End Property

Public Sub GetPixel(ByVal X As Long, ByVal Y As Long, R As Long, G As Long, B As Long, Optional A As Long)
    With m_aBitsRGB(X, Y)
        R = .R
        B = .B
        G = .G
        A = .A
    End With
End Sub

Public Sub SetPixel(ByVal X As Long, ByVal Y As Long, ByVal R As Long, ByVal G As Long, ByVal B As Long, Optional ByVal A As Long = -1)
    With m_aBitsRGB(X, Y)
        .R = R
        .B = B
        .G = G
        If A >= 0 Then
            .A = A
            m_lPixelFormat = m_lPixelFormat Or PixelFormatAlpha
        End If
    End With
End Sub

Public Function ToPictureAndTerminate() As StdPicture
    Dim uDesc           As PICTDESC
    Dim aGUID(0 To 3)   As Long
    Dim baData()        As Byte
    Dim hIcon           As Long
    
    If HasAlpha Then
        baData = SaveToByteArray("image/png")
        hIcon = CreateIconFromResourceEx(baData(0), UBound(baData) + 1, 1, &H30000, 0, 0, 0)
    End If
    '--- fill struct
    With uDesc
        .lSize = Len(uDesc)
        .lType = IIf(hIcon <> 0, vbPicTypeIcon, vbPicTypeBitmap)
        .hBmp = IIf(hIcon <> 0, hIcon, m_hDib)
    End With
    '--- Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    '--- Create picture from bitmap handle
    If OleCreatePictureIndirect(uDesc, aGUID(0), 1, ToPictureAndTerminate) <> 0 Then ' S_OK
        Call DestroyIcon(hIcon)
    End If
    If hIcon = 0 Then
        m_hDib = 0
    End If
    Terminate
End Function

Public Function ApplyEffect( _
            Optional BlurRadius As Double, _
            Optional ByVal BlurExpandEdges As Boolean, _
            Optional Brighthness As Double) As Boolean
    Const UnitPixel     As Long = 2&
    Dim hBitmap         As Long
    Dim hGraphics       As Long
    Dim hEffect         As Long
    Dim uParams         As GdipBlurParams
    Dim hAttributes     As Long
    Dim clrMatrix(0 To 4, 0 To 4) As Single
    
    If Not StartGdip() Then
        GoTo QH
    End If
    If GdipCreateBitmapFromScan0(Width, Height, BytesPerScanLine, IIf(Not HasAlpha, PixelFormat32bppRGB, PixelFormat32bppARGB), DIBitsPtr, hBitmap) <> 0 Then
        GoTo QH
    End If
    If GdipCreateFromHDC(hDC, hGraphics) <> 0 Then
        GoTo QH
    End If
    If BlurRadius > 0 Then
        If GdipCreateEffect(&H633C80A4, &H482B1843, &H28BEF29E, &HD4FDC534, hEffect) <> 0 Then
            GoTo QH
        End If
        uParams.Radius = BlurRadius
        uParams.ExpandEdge = -BlurExpandEdges
        If GdipSetEffectParameters(hEffect, uParams, Len(uParams)) <> 0 Then
            GoTo QH
        End If
        If GdipBitmapApplyEffect(hBitmap, hEffect, ByVal 0) <> 0 Then
            GoTo QH
        End If
    End If
    If Brighthness <> 0 Then
        If GdipCreateImageAttributes(hAttributes) <> 0 Then
            GoTo QH
        End If
        clrMatrix(0, 0) = Brighthness
        clrMatrix(1, 1) = Brighthness
        clrMatrix(2, 2) = Brighthness
        clrMatrix(3, 3) = 1
        clrMatrix(0, 4) = 1
        clrMatrix(1, 4) = 1
        clrMatrix(2, 4) = 1
        clrMatrix(4, 4) = 1
        If GdipSetImageAttributesColorMatrix(hAttributes, 0, 1, clrMatrix(0, 0), clrMatrix(0, 0), 0) <> 0 Then
            GoTo QH
        End If
    End If
    If GdipDrawImageRectRectI(hGraphics, hBitmap, 0, 0, Me.Width, Me.Height, 0, 0, Me.Width, Me.Height, UnitPixel, hAttributes, 0, 0) <> 0 Then
        GoTo QH
    End If
    '--- success
    ApplyEffect = True
QH:
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    If hEffect <> 0 Then
        Call GdipDeleteEffect(hEffect)
    End If
    If hAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hAttributes)
    End If
    Exit Function
EH:
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    If hEffect <> 0 Then
        Call GdipDeleteEffect(hEffect)
    End If
    If hAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hAttributes)
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ApplyBlur(BlurRadius As Double) As Boolean
    pvBlurChannel m_aBits, Width, Height, BlurRadius, &H1
    pvBlurChannel m_aBits, Width, Height, BlurRadius, &H100
    pvBlurChannel m_aBits, Width, Height, BlurRadius, &H10000
    If HasAlpha Then
        pvBlurChannel m_aBits, Width, Height, BlurRadius, &H1000000
    End If
    '--- success
    ApplyBlur = True
End Function

'= private ===============================================================

Private Function pvBlurChannel( _
            laBits() As Long, _
            ByVal lWidth As Long, _
            ByVal lHeight As Long, _
            dblRadius As Double, _
            ByVal lBitShift As Long) As Boolean
'--- based on https://github.com/tannerhelland/PhotoDemon/blob/master/Modules/Filters_ByteArray.bas#L40
    Const ONE_DIV_255   As Double = 1# / 255#
    Const NUM_ITERS     As Long = 3
    Dim lX              As Long
    Dim lY              As Long
    Dim lInitX          As Long
    Dim lInitY          As Long
    Dim lFinalX         As Long
    Dim lFinalY         As Long
    Dim lG              As Long
    Dim dblNu           As Double
    Dim dblBndryScale   As Double
    Dim dblPostScale    As Double
    Dim lIdx            As Long
    Dim dblTemp         As Double
    Dim dblArray()      As Double
    Dim lBitMask        As Long
    
    If lBitShift < &H1000000 Then
        lBitMask = &HFF& * lBitShift
    Else
        lBitMask = &HFF000000
    End If
    lInitX = 0
    lInitY = 0
    lFinalX = lWidth - 1
    lFinalY = lHeight - 1
    'Prep some IIR-specific values next
    dblTemp = Sqr(-(dblRadius * dblRadius) / (2 * Log(1# / 255#)))
    If dblTemp <= 0 Then
        dblTemp = 0.01
    End If
    dblTemp = dblTemp * (1# + (0.3165 * NUM_ITERS + 0.5695) / ((NUM_ITERS + 0.7818) * (NUM_ITERS + 0.7818)))
    dblTemp = (dblTemp * dblTemp) / (2# * NUM_ITERS)
    dblNu = (1# + 2# * dblTemp - Sqr(1# + 4# * dblTemp)) / (2# * dblTemp)
    dblBndryScale = (1# / (1# - dblNu))
    dblPostScale = ((dblNu / dblTemp) ^ (2# * NUM_ITERS)) * 255#
    'Copy the contents of the incoming byte array into the float array
    ReDim dblArray(lInitX To lFinalX, lInitY To lFinalY) As Double
    For lX = lInitX To lFinalX
        For lY = lInitY To lFinalY
            lG = ((laBits(lX, lY) And lBitMask) \ lBitShift) And &HFF
            dblArray(lX, lY) = lG * ONE_DIV_255
        Next
    Next
    'Filter horizontally along each row
    For lY = lInitY To lFinalY
        For lIdx = 0 To NUM_ITERS - 1
            dblArray(lInitX, lY) = dblArray(lInitX, lY) * dblBndryScale
            For lX = lInitX + 1 To lFinalX
                dblArray(lX, lY) = dblArray(lX, lY) + dblNu * dblArray(lX - 1, lY)
            Next
            dblArray(lFinalX, lY) = dblArray(lFinalX, lY) * dblBndryScale
            For lX = lFinalX - 1 To lInitX Step -1
                dblArray(lX, lY) = dblArray(lX, lY) + dblNu * dblArray(lX + 1, lY)
            Next
        Next
    Next
    'Now repeat all the above steps, but filtering vertically along each column, instead
    For lX = lInitX To lFinalX
        For lIdx = 0 To NUM_ITERS - 1
            dblArray(lX, lInitY) = dblArray(lX, lInitY) * dblBndryScale
            For lY = lInitY + 1 To lFinalY
                dblArray(lX, lY) = dblArray(lX, lY) + dblNu * dblArray(lX, lY - 1)
            Next
            dblArray(lX, lFinalY) = dblArray(lX, lFinalY) * dblBndryScale
            For lY = lFinalY - 1 To lInitY Step -1
                dblArray(lX, lY) = dblArray(lX, lY) + dblNu * dblArray(lX, lY + 1)
            Next
        Next
    Next
    'Apply final post-scaling
    For lX = lInitX To lFinalX
        For lY = lInitY To lFinalY
            lG = dblArray(lX, lY) * dblPostScale
            If lG > 255 Then
                lG = 255
            End If
            laBits(lX, lY) = (laBits(lX, lY) And Not lBitMask) Or (lG * lBitShift)
        Next
    Next
    pvBlurChannel = True
End Function

Private Function pvGetEncoderClsid(sMimeType As String) As GdipGUID
    Dim lNumCoders      As Long
    Dim lSize           As Long
    Dim uInfo()         As GdipImageCodecInfo
    Dim lIdx            As Long

    Call GdipGetImageEncodersSize(lNumCoders, lSize)
    If lSize = 0 Then
        GoTo QH
    End If
    ReDim uInfo(0 To lSize \ LenB(uInfo(0))) As GdipImageCodecInfo
    If GdipGetImageEncoders(lNumCoders, lSize, uInfo(0)) <> 0 Then
        GoTo QH
    End If
    For lIdx = 0 To lNumCoders - 1
        If StrComp(SysAllocString(uInfo(lIdx).MimeType), sMimeType, vbTextCompare) = 0 Then
            pvGetEncoderClsid = uInfo(lIdx).GUID
        End If
    Next
QH:
End Function

Private Sub pvCalcDimensions(sngWidth As Single, sngHeight As Single, ByVal bKeepAspect As Boolean)
    If Not bKeepAspect And (sngWidth = 0 Or sngHeight = 0) Then
        sngWidth = m_sngOrigWidth
        sngHeight = m_sngOrigHeight
    ElseIf bKeepAspect Then
        If sngWidth <> 0 And sngHeight = 0 Then
            If m_sngOrigWidth <> 0 Then
                sngHeight = m_sngOrigHeight * sngWidth / m_sngOrigWidth
            End If
        ElseIf sngWidth = 0 And sngHeight <> 0 Then
            If m_sngOrigHeight <> 0 Then
                sngWidth = m_sngOrigWidth * sngHeight / m_sngOrigHeight
            End If
        ElseIf sngWidth = 0 Or sngHeight = 0 Or m_sngOrigHeight = 0 Or m_sngOrigWidth = 0 Then
            sngWidth = m_sngOrigWidth
            sngHeight = m_sngOrigHeight
        ElseIf sngWidth / sngHeight > m_sngOrigWidth / m_sngOrigHeight Then
            sngWidth = m_sngOrigWidth * sngHeight / m_sngOrigHeight
        Else
            sngHeight = m_sngOrigHeight * sngWidth / m_sngOrigWidth
        End If
    End If
End Sub

#If SelfContained Then
Private Function Peek(ByVal lPtr As Long) As Long
    Call GetMem4(ByVal lPtr, Peek)
End Function

Private Function TranslateColor(ByVal clrValue As OLE_COLOR) As OLE_COLOR
    Call OleTranslateColor(clrValue, 0, VarPtr(TranslateColor))
End Function

Private Function C_Lng(Value As Variant) As Long
    On Error GoTo RH
    C_Lng = CLng(Value)
RH:
End Function

Private Function StartGdip() As Boolean
    Const FUNC_NAME     As String = "StartGdip"
    Dim lGdipStatus     As Long
    Dim uStartup        As GdiplusStartupInput
    Dim sSource         As String

    If g_hGdiPlus = 0 Then
        uStartup.GdiplusVersion = 1&
        On Error GoTo RH
        lGdipStatus = GdiplusStartup(g_hGdiPlus, uStartup)
RH:
        If g_hGdiPlus = 0 Then
            If lGdipStatus = 0 Then
                Err.Raise vbObjectError, MODULE_NAME & "." & FUNC_NAME & "(sSource=" & sSource & ")", "GDI+ not installed"
            Else
                sSource = "GdiplusStartup"
                Err.Raise vbObjectError, MODULE_NAME & "." & FUNC_NAME & "(sSource=" & sSource & ")", GdipGetSystemMessage(lGdipStatus, Err.LastDllError)
            End If
        End If
    End If
    '--- success
    StartGdip = True
End Function

Private Function GdipGetSystemMessage(ByVal lGdipStatus As Long, ByVal lLastDllError As Long) As String
    Const STR_STATUS    As String = "|Generic Error|Invalid Parameter|Out Of Memory|Object Busy|Insufficient Buffer|Not Implemented|Win32 Error|Wrong State|AbortedFile Not Found|Value Overflow|Access Denied|Unknown Image Format|Font Family Not Found|Font Style Not Found|Not True Type Font|Unsupported Gdiplus Version|Gdiplus Not Initialized|Property Not Found|Property Not Supported|Profile Not Found"
    
    If lGdipStatus <> 0 Then
        GdipGetSystemMessage = At(Split(STR_STATUS, "|"), lGdipStatus, "Unknown error: " & lGdipStatus)
        If lGdipStatus = 7 And lLastDllError <> 0 Then ' Win32Error
            GdipGetSystemMessage = GdipGetSystemMessage & ". " & GetSystemMessage(lLastDllError)
        End If
    End If
End Function

Private Function GetSystemMessage(ByVal lLastDllError As Long) As String
    Dim ret             As Long
   
    GetSystemMessage = Space$(2000)
    ret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lLastDllError, 0&, GetSystemMessage, Len(GetSystemMessage), 0&)
    If ret > 2 Then
        If Mid$(GetSystemMessage, ret - 1, 2) = vbCrLf Then
            ret = ret - 2
        End If
    End If
    GetSystemMessage = Left$(GetSystemMessage, ret)
End Function

Private Function At(Data As Variant, ByVal Index As Long, Optional Default As String) As String
    On Error GoTo RH
    At = Default
    If IsArray(Data) Then
        If LBound(Data) <= Index And Index <= UBound(Data) Then
            At = Data(Index)
        End If
    End If
RH:
End Function

Private Function HM2Pix(ByVal Value As Double) As Long
   HM2Pix = Int(Value * 1440 / 2540 / Screen.TwipsPerPixelX + 0.5)
End Function

Private Sub RenderPicture(pPicture As IPicture, hDC As Long, X As Long, Y As Long, cx As Long, cy As Long, xSrc As OLE_XPOS_HIMETRIC, ySrc As OLE_YPOS_HIMETRIC, cxSrc As OLE_XSIZE_HIMETRIC, cySrc As OLE_YSIZE_HIMETRIC)
    '--- pod Wine OLEPictureImpl_Render ne e realizirana dokraj
    '--- [quote]
    '--- case PICTYPE_ICON:
    '---   FIXME("Not quite correct implementation of rendering icons...\n");
    '---   DrawIcon(hdc,x,y,This->desc.u.icon.hicon);
    '---   break;
    '--- [/quote]
    If pPicture Is Nothing Then
        Exit Sub
    End If
    If pPicture.Handle = 0 Then
        Exit Sub
    End If
    If pPicture.Type = vbPicTypeIcon Then
        Call DrawIconEx(hDC, X, Y, pPicture.Handle, cx, cy, 0, 0, DI_NORMAL)
    Else
        pPicture.Render hDC, X, Y, cx, cy, xSrc, ySrc, cxSrc, cySrc, ByVal 0
    End If
End Sub

Private Function LimitLong( _
            ByVal lValue As Long, _
            Optional ByVal Min As Long = -2147483647, _
            Optional ByVal Max As Long = 2147483647) As Long
    If lValue < Min Then
        LimitLong = Min
    ElseIf lValue > Max Then
        LimitLong = Max
    Else
        LimitLong = lValue
    End If
End Function
#End If

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    Terminate
End Sub
