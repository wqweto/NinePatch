VERSION 5.00
Begin VB.UserControl ctxTouchKeyboard 
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   Windowless      =   -1  'True
   Begin Project1.ctxTouchButton btn 
      Height          =   684
      Index           =   0
      Left            =   168
      Top             =   0
      Visible         =   0   'False
      Width           =   768
      _ExtentX        =   1355
      _ExtentY        =   1207
      AnimationDuration=   0.1
      Caption         =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "PT Sans Narrow"
         Size            =   13.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "ctxTouchKeyboard.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "ctxTouchKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "ctxTouchKeyboard"

#Const ImplUseShared = NPPNG_USE_SHARED <> 0

'=========================================================================
' Public Events
'=========================================================================

Event ButtonClick(ByVal Index As Long)
Event ButtonMouseDown(ByVal Index As Long)
Event RegisterCancelMode(oCtl As Object, Handled As Boolean)

'=========================================================================
' API
'=========================================================================

Private Const UnitPixel                     As Long = 2
'--- colors
Private Const Transparent                   As Long = &HFFFFFF
'--- for GdipCreateBitmapFromScan0
Private Const PixelFormat32bppARGB          As Long = &H26200A
'--- for GdipSetCompositingMode
Private Const CompositingModeSourceCopy     As Long = 1
'--- for GdipCreatePath
Private Const FillModeAlternate             As Long = 0
'--- for GdipSetPenDashStyle
Private Const DashStyleSolid                As Long = 0
'--- for GdipSetSmoothingMode
Private Const SmoothingModeAntiAlias        As Long = 4

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'--- gdi+
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, pInputBuf As Any, Optional ByVal pOutputBuf As Long = 0) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal hImage As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal lWidth As Long, ByVal lHeight As Long, ByVal lStride As Long, ByVal lPixelFormat As Long, ByVal Scan0 As Long, hBitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As Long
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal srcUnit As Long = UnitPixel, Optional ByVal hImageAttributes As Long, Optional ByVal pfnCallback As Long, Optional ByVal lCallbackData As Long) As Long
Private Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal hBitmap As Long, ByVal lX As Long, ByVal lY As Long, ByVal lColor As Long) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal hGraphics As Long, ByVal lColor As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal lArgb As Long, hBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As Long
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal lCompositingMode As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBmp As Long, ByVal hPal As Long, hBtmap As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef nWidth As Single, ByRef nHeight As Single) As Long '
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (hImgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hImgAttr As Long, ByVal lAdjustType As Long, ByVal fAdjustEnabled As Long, clrMatrix As Any, grayMatrix As Any, ByVal lFlags As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImgAttr As Long) As Long
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

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const DEF_LAYOUT1           As String = "q w e r t y u i o p <=|1.25|D " & _
                                                "|0.5|S|N a s d f g h j k l Done|1.85|B " & _
                                                "^|||N z x c v b n m ! ? ^|1.25| " & _
                                                "?!123|3|D|N _|6 ?!123|1.25|D keyb||D"
Private Const DEF_LAYOUT2           As String = "Q W E R T Y U I O P <=|1.25|D " & _
                                                "|0.5|S|N A S D F G H J K L Done|1.85|B " & _
                                                "^||L|N Z X C V B N M ! ? ^|1.25|L " & _
                                                "?!123|3|D|N _|6 ?!123|1.25|D keyb||D"

Private m_sLayout               As String
Private m_lButtonCurrent        As Long
Private m_cButtonRows()         As Collection
Private m_oCtlCancelMode        As Object
Private m_hForeBitmap           As Long
Private m_cButtonImageCache     As Collection

Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    B                   As Byte
    A                   As Byte
End Type

Private Type UcsHsbColor
    Hue                 As Double
    Sat                 As Double
    Bri                 As Double
    A                   As Byte
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Function PrintError(sFunction As String) As VbMsgBoxResult
    Debug.Print Err.Description & " [" & STR_MODULE_NAME & "." & sFunction & "]", Timer
End Function

'Private Function RaiseError(sFunction As String) As VbMsgBoxResult
'    Err.Raise Err.Number, STR_MODULE_NAME & "." & sFunction & vbCrLf & Err.Source, Err.Description
'End Function

'=========================================================================
' Properties
'=========================================================================

Property Get Layout() As String
    Layout = m_sLayout
End Property

Property Let Layout(sValue As String)
    m_sLayout = sValue
    pvLoadLayout sValue
    pvSizeLayout
End Property

Property Get ButtonCaption(ByVal Index As Long) As String
    ButtonCaption = btn(Index).Caption
End Property

Property Get ButtonTag(ByVal Index As Long) As String
    ButtonTag = btn(Index).Tag
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub RegisterCancelMode(oCtl As Object)
    pvRegisterCancelMode Me
    If Not m_oCtlCancelMode Is Nothing And Not m_oCtlCancelMode Is oCtl Then
        m_oCtlCancelMode.CancelMode
    End If
    Set m_oCtlCancelMode = oCtl
End Sub

Public Sub CancelMode()
    If Not m_oCtlCancelMode Is Nothing Then
        m_oCtlCancelMode.CancelMode
        Set m_oCtlCancelMode = Nothing
    End If
End Sub

'= private ===============================================================

Private Sub pvLoadLayout(sLayout As String)
    Const FUNC_NAME     As String = "pvLoadLayout"
    Const CLR_GREY      As Long = &HFF484848
    Const CLR_BLUE      As Long = &HFF0565FF ' &HFF1971FE
    Const CLR_LIGHT     As Long = &HFFA9A9A9 ' &HFF737373
    Const CLR_DARK      As Long = &HFF232323
    Dim lIdx            As Long
    Dim lRow            As Long
    Dim vElem           As Variant
    Dim vSplit          As Variant
    
    On Error GoTo EH
    For m_lButtonCurrent = m_lButtonCurrent To 1 Step -1
        btn(m_lButtonCurrent).Visible = False
    Next
    ReDim m_cButtonRows(0 To 0) As Collection
    Set m_cButtonRows(0) = New Collection
    For Each vElem In Split(sLayout)
        vSplit = Split(Replace(vElem, "_", " "), "|")
        Select Case At(vSplit, 2)
        Case "D"
            lIdx = pvLoadButton(CLR_DARK)
        Case "L"
            lIdx = pvLoadButton(CLR_LIGHT)
        Case "B"
            lIdx = pvLoadButton(CLR_BLUE)
        Case "S"
            lIdx = pvLoadButton(Transparent)
        Case Else
            lIdx = pvLoadButton(CLR_GREY)
        End Select
        If lIdx > 0 Then
            With btn(lIdx - 1)
                Select Case IIf(lIdx = 1, "F", At(vSplit, 3))
                Case "F"
                    btn(lIdx).Move 0, 0, btn(0).Width, btn(0).Width
                    lRow = 0
                Case "N"
                    btn(lIdx).Move 0, .Top + .Height, btn(0).Width, btn(0).Width
                    lRow = lRow + 1
                    ReDim Preserve m_cButtonRows(0 To lRow) As Collection
                    Set m_cButtonRows(lRow) = New Collection
                Case Else
                    btn(lIdx).Move .Left + .Width, .Top, btn(0).Width, btn(0).Width
                End Select
            End With
            With btn(lIdx)
                .Caption = vSplit(0)
                .Tag = vElem
                .Visible = True
                If LenB(At(vSplit, 1)) <> 0 Then
                    .Width = C_Dbl(At(vSplit, 1)) * .Width
                End If
                m_cButtonRows(lRow).Add Array(lIdx, .Width)
            End With
        End If
    Next
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvSizeLayout()
    Const FUNC_NAME     As String = "pvSizeLayout"
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim dblLeft         As Double
    Dim dblTotal        As Double
    
    On Error GoTo EH
    For lIdx = 0 To UBound(m_cButtonRows)
        dblTotal = 0
        For Each vElem In m_cButtonRows(lIdx)
            dblTotal = dblTotal + vElem(1)
        Next
        dblLeft = 0
        For Each vElem In m_cButtonRows(lIdx)
            With btn(vElem(0))
                .Left = AlignTwipsToPix(dblLeft * ScaleWidth / dblTotal)
                dblLeft = dblLeft + vElem(1)
                .Width = AlignTwipsToPix(dblLeft * ScaleWidth / dblTotal) - .Left
            End With
        Next
    Next
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvLoadButton(ByVal clrBack As Long) As Long
    Dim clrBorder       As Long
    Dim clrShadow       As Long
    Dim hNormalBitmap   As Long
    Dim hHoverBitmap    As Long
    Dim hPressedBitmap  As Long
    
    On Error GoTo EH
    If clrBack <> Transparent Then
        If SearchCollection(m_cButtonImageCache, "N" & Hex(clrBack)) Then
            hNormalBitmap = m_cButtonImageCache.Item("N" & Hex(clrBack))
            hHoverBitmap = m_cButtonImageCache.Item("H" & Hex(clrBack))
            hPressedBitmap = m_cButtonImageCache.Item("P" & Hex(clrBack))
        Else
            clrBorder = pvAdjustColor(clrBack, AdjustBri:=0.3, AdjustAlpha:=-0.5)
            clrShadow = pvAdjustColor(clrBack, AdjustBri:=-0.8, AdjustAlpha:=-0.75)
            If Not pvPrepareButtonBitmap(6, 5, clrBorder, clrBack, clrShadow, hNormalBitmap) Then
                GoTo QH
            End If
            If Not pvPrepareButtonBitmap(6, 5, clrBorder, pvAdjustColor(clrBack, AdjustBri:=-0.2), clrShadow, hHoverBitmap) Then
                GoTo QH
            End If
            clrShadow = pvAdjustColor(clrShadow, AdjustAlpha:=-0.75)
            If Not pvPrepareButtonBitmap(6, 5, clrBorder, pvAdjustColor(clrBack, AdjustBri:=0.25, AdjustSat:=-0.25), clrShadow, hPressedBitmap) Then
                GoTo QH
            End If
            m_cButtonImageCache.Add hNormalBitmap, "N" & Hex(clrBack)
            m_cButtonImageCache.Add hHoverBitmap, "H" & Hex(clrBack)
            m_cButtonImageCache.Add hPressedBitmap, "P" & Hex(clrBack)
        End If
    End If
    m_lButtonCurrent = m_lButtonCurrent + 1
    If m_lButtonCurrent > btn.UBound Then
        Load btn(m_lButtonCurrent)
    End If
    pvLoadButton = m_lButtonCurrent
    With btn(pvLoadButton)
        .ZOrder vbBringToFront
        .ButtonImageBitmap(ucsBstNormal) = hNormalBitmap
        .ButtonImageBitmap(ucsBstHover) = hHoverBitmap
        .ButtonImageBitmap(ucsBstPressed) = hPressedBitmap
    End With
QH:
    Exit Function
EH:
End Function

Private Function pvPrepareButtonBitmap( _
            ByVal sngRadius As Single, _
            ByVal sngBlur As Single, _
            ByVal clrPen As Long, _
            ByVal clrBack As Long, _
            ByVal clrShadow As Long, _
            hBitmap As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareButtonBitmap"
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
    If Not pvDrawRoundedRectangle(hGraphics, 1 + Ceil(sngBlur) + SHADOW_OFFSET / 2, 1 + Ceil(sngBlur) + SHADOW_OFFSET, _
            lRoundWidth, lRoundWidth, sngRadius, clrShadow, clrBack:=clrShadow) Then
        GoTo QH
    End If
    If Not BlurBitmap(hDropShadow, sngBlur) Then
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
    If Not pvDrawRoundedRectangle(hGraphics, 1 + Ceil(sngBlur), 1 + Ceil(sngBlur), lRoundWidth, lRoundWidth, sngRadius, clrPen, clrBack:=clrBack) Then
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
    pvPrepareButtonBitmap = True
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

Private Function pvPrepareForeground(hFore As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareForeground"
    Dim hBrush          As Long
    Dim hGraphics       As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim lIdx            As Long
    Dim hBitmap         As Long
    Dim sngWidth        As Single
    Dim sngHeight       As Single
    Dim hAttributes     As Long
    
    On Error GoTo EH
    lWidth = ScaleWidth \ Screen.TwipsPerPixelX
    lHeight = ScaleWidth \ Screen.TwipsPerPixelX
    If hFore <> 0 Then
        Call GdipDisposeImage(hFore)
        hFore = 0
    End If
    If GdipCreateBitmapFromScan0(lWidth, lHeight, lWidth * 4, PixelFormat32bppARGB, 0, hFore) <> 0 Then
        GoTo QH
    End If
    If GdipGetImageGraphicsContext(hFore, hGraphics) <> 0 Then
        GoTo QH
    End If
    If GdipCreateSolidFill(&H1000000, hBrush) <> 0 Then
        GoTo QH
    End If
    If GdipSetCompositingMode(hGraphics, CompositingModeSourceCopy) <> 0 Then
        GoTo QH
    End If
    If GdipFillRectangleI(hGraphics, hBrush, 0, 0, lWidth, lHeight) <> 0 Then
        GoTo QH
    End If
    lIdx = Sqr(lWidth * lHeight)
    If GdipCreateBitmapFromHBITMAP(Image1.Picture.Handle, 0, hBitmap) <> 0 Then
        GoTo QH
    End If
    If GdipGetImageDimension(hBitmap, sngWidth, sngHeight) <> 0 Then
        GoTo QH
    End If
    If Not pvPrepareAttribs(0.1, hAttributes) Then
        GoTo QH
    End If
    If GdipDrawImageRectRect(hGraphics, hBitmap, 0, 0, lWidth, lHeight, 0, 0, sngWidth, sngHeight, , hAttributes) <> 0 Then
        GoTo QH
    End If
QH:
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
    End If
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    If hAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hAttributes)
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

Private Function pvDrawRoundedRectangle( _
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
    Const FUNC_NAME     As String = "pvDrawRoundedRectangle"
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
    pvDrawRoundedRectangle = True
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

Private Function pvRegisterCancelMode(oCtl As Object) As Boolean
    Dim bHandled        As Boolean
    
    RaiseEvent RegisterCancelMode(oCtl, bHandled)
    If Not bHandled Then
        On Error GoTo QH
        Parent.RegisterCancelMode oCtl
        On Error GoTo 0
    End If
    '--- success
    pvRegisterCancelMode = True
QH:
End Function

Private Function pvAdjustColor( _
            ByVal clrValue As Long, _
            Optional ByVal AdjustBri As Double, _
            Optional ByVal AdjustSat As Double, _
            Optional ByVal AdjustAlpha As Double) As Long
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
    pvAdjustColor = pvHSBToRGB(hsbColor)
End Function

Private Function pvHSBToRGB(hsbColor As UcsHsbColor) As Long
'--- based on *cool* code by Branco Medeiros (http://www.myrealbox.com/branco_medeiros)
'--- Converts an HSB value to the RGB color model. Adapted from Java.awt.Color.java
    Dim dblH            As Double
    Dim dblS            As Double
    Dim dblL            As Double
    Dim dblF            As Double
    Dim dblP            As Double
    Dim dblQ            As Double
    Dim dblT            As Double
    Dim lH              As Long
    Dim rgbColor        As UcsRgbQuad

    With rgbColor
        If hsbColor.Sat > 0 Then
            dblH = hsbColor.Hue * 6 '/ 60
            dblL = hsbColor.Bri '/ 100
            dblS = hsbColor.Sat '/ 100
            lH = Int(dblH)
            dblF = dblH - lH
            dblP = dblL * (1 - dblS)
            dblQ = dblL * (1 - dblS * dblF)
            dblT = dblL * (1 - dblS * (1 - dblF))
            Select Case lH
            Case 0
                .R = dblL * 255
                .G = dblT * 255
                .B = dblP * 255
            Case 1
                .R = dblQ * 255
                .G = dblL * 255
                .B = dblP * 255
            Case 2
                .R = dblP * 255
                .G = dblL * 255
                .B = dblT * 255
            Case 3
                .R = dblP * 255
                .G = dblQ * 255
                .B = dblL * 255
            Case 4
                .R = dblT * 255
                .G = dblP * 255
                .B = dblL * 255
            Case 5
                .R = dblL * 255
                .G = dblP * 255
                .B = dblQ * 255
            End Select
        Else
            .R = hsbColor.Bri * 255
            .G = .R
            .B = .R
        End If
        .A = hsbColor.A
    End With
    Call CopyMemory(pvHSBToRGB, rgbColor, 4)
End Function

Private Function pvRGBToHSB(ByVal clrValue As OLE_COLOR) As UcsHsbColor
'--- based on *cool* code by Branco Medeiros (http://www.myrealbox.com/branco_medeiros)
'--- Converts an RGB value to the HSB color model. Adapted from Java.awt.Color.java
    Dim dblTemp         As Double
    Dim lMin            As Long
    Dim lMax            As Long
    Dim lDelta          As Long
    Dim rgbColor        As UcsRgbQuad
  
    Call CopyMemory(rgbColor, clrValue, 4)
    If rgbColor.R > rgbColor.G Then
        If rgbColor.R > rgbColor.B Then
            lMax = rgbColor.R
        Else
            lMax = rgbColor.B
        End If
    ElseIf rgbColor.G > rgbColor.B Then
        lMax = rgbColor.G
    Else
        lMax = rgbColor.B
    End If
    If rgbColor.R < rgbColor.G Then
        If rgbColor.R < rgbColor.B Then
            lMin = rgbColor.R
        Else
            lMin = rgbColor.B
        End If
    ElseIf rgbColor.G < rgbColor.B Then
        lMin = rgbColor.G
    Else
        lMin = rgbColor.B
    End If
    lDelta = lMax - lMin
    pvRGBToHSB.Bri = lMax / 255
    If lMax > 0 Then
        pvRGBToHSB.Sat = lDelta / lMax
        If lDelta > 0 Then
            If lMax = rgbColor.R Then
                dblTemp = (CLng(rgbColor.G) - rgbColor.B) / lDelta
            ElseIf lMax = rgbColor.G Then
                dblTemp = 2 + (CLng(rgbColor.B) - rgbColor.R) / lDelta
            Else
                dblTemp = 4 + (CLng(rgbColor.R) - rgbColor.G) / lDelta
            End If
            pvRGBToHSB.Hue = dblTemp / 6
            If pvRGBToHSB.Hue < 0 Then
                pvRGBToHSB.Hue = pvRGBToHSB.Hue + 1
            End If
        End If
    End If
    pvRGBToHSB.A = rgbColor.A
'    Debug.Assert pvHSBToRGB(pvRGBToHSB) = clrValue
End Function

Private Function Ceil(ByVal Value As Double) As Double
    Ceil = -Int(CStr(-Value))
End Function

Private Function IsCompileTime(Extender As Object) As Boolean
    Dim oTopParent      As Object
    Dim oUserControl    As UserControl
    
    On Error GoTo QH
    Set oTopParent = Extender.Parent
    Set oUserControl = AsUserControl(oTopParent)
    Do While Not oUserControl Is Nothing
        If oUserControl.Parent Is Nothing Then
            Exit Do
        End If
        Set oTopParent = oUserControl.Parent
        Set oUserControl = AsUserControl(oTopParent)
    Loop
    Select Case TypeName(oTopParent)
    Case "Form", "UserControl"
        IsCompileTime = True
    End Select
QH:
End Function

Private Function AsUserControl(oObj As Object) As UserControl
    Dim pControl        As UserControl
  
    If TypeOf oObj Is Form Then
        '--- do nothing
    Else
        Call CopyMemory(pControl, ObjPtr(oObj), 4)
        Set AsUserControl = pControl
        Call CopyMemory(pControl, 0&, 4)
    End If
End Function

#If Not ImplUseShared Then

Private Function At(Data As Variant, ByVal Index As Long, Optional Default As String) As String
    On Error GoTo RH
    At = Default
    If LBound(Data) <= Index And Index <= UBound(Data) Then
        At = CStr(Data(Index))
    End If
RH:
End Function

Private Function C_Dbl(Value As Variant) As Double
    Dim vDest           As Variant
    
    If VarType(Value) = vbDouble Then
        C_Dbl = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbDouble) = 0 Then
        C_Dbl = vDest
    End If
End Function

Public Function AlignTwipsToPix(ByVal dblTwips As Double) As Double
    AlignTwipsToPix = Int(dblTwips / Screen.TwipsPerPixelX + 0.5) * Screen.TwipsPerPixelX
End Function

#End If ' Not ImplUseShared

'=========================================================================
' Control events
'=========================================================================

Private Sub btn_Click(Index As Integer)
    RaiseEvent ButtonClick(Index)
End Sub

'Private Sub btn_DblClick(Index As Integer)
'    RaiseEvent Click(Index)
'End Sub

Private Sub btn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent ButtonMouseDown(Index)
End Sub

Private Sub btn_OwnerDraw(Index As Integer, ByVal hGraphics As Long, ByVal hFont As Long, ByVal ButtonState As UcsTouchButtonStateEnum, ClientLeft As Long, ClientTop As Long, ClientWidth As Long, ClientHeight As Long, Caption As String, ByVal hPicture As Long)
    Const FUNC_NAME     As String = "btn_OwnerDraw"
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    
    On Error GoTo EH
    If m_hForeBitmap = 0 Then
        GoTo QH
    End If
    With btn(Index)
        lLeft = .Left \ Screen.TwipsPerPixelX
        lTop = .Top \ Screen.TwipsPerPixelY
        lWidth = .Width \ Screen.TwipsPerPixelX
        lHeight = .Height \ Screen.TwipsPerPixelX
    End With
    If GdipDrawImageRectRect(hGraphics, m_hForeBitmap, 0, 0, lWidth, lHeight, lLeft, lTop, lWidth, lHeight) <> 0 Then
        GoTo QH
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CancelMode
End Sub

Private Sub UserControl_InitProperties()
    pvPrepareForeground m_hForeBitmap
    Layout = DEF_LAYOUT1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If IsCompileTime(Extender) Then
        Exit Sub
    End If
    pvPrepareForeground m_hForeBitmap
    Layout = DEF_LAYOUT1
End Sub

Private Sub UserControl_Initialize()
    Dim aInput(0 To 3)  As Long
    
    If GetModuleHandle("gdiplus") = 0 Then
        aInput(0) = 1
        Call GdiplusStartup(0, aInput(0))
    End If
    Set m_cButtonImageCache = New Collection
    ReDim m_cButtonRows(0 To 0) As Collection
    Set m_cButtonRows(0) = New Collection
End Sub

Private Sub UserControl_Resize()
    pvPrepareForeground m_hForeBitmap
    pvSizeLayout
End Sub

Private Sub UserControl_Terminate()
    Dim vElem           As Variant
    
    For Each vElem In m_cButtonImageCache
        Call GdipDisposeImage(vElem)
    Next
End Sub

