VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   7308
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13524
   LinkTopic       =   "Form2"
   ScaleHeight     =   7308
   ScaleWidth      =   13524
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   684
      Left            =   4956
      TabIndex        =   2
      Top             =   168
      Width           =   1356
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   684
      Left            =   3444
      TabIndex        =   0
      Top             =   168
      Width           =   1356
   End
   Begin Project1.ctxNineButton ctxNineButton1 
      Height          =   1020
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   2268
      Visible         =   0   'False
      Width           =   1104
      _extentx        =   1947
      _extenty        =   1799
      style           =   0
      animationduration=   0.1
      caption         =   "D"
      font            =   "Form2.frx":0000
      forecolor       =   15001582
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7644
      Picture         =   "Form2.frx":0030
      Top             =   420
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const UnitPixel                     As Long = 2

'--- for GdipCreateBitmapFromScan0
Private Const PixelFormat32bppARGB          As Long = &H26200A
'Private Const PixelFormat32bppPARGB         As Long = &HE200B
'--- for GdipSetCompositingMode
'Private Const CompositingModeSourceOver     As Long = 0
Private Const CompositingModeSourceCopy     As Long = 1

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
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
Private Declare Function GdipCloneImage Lib "gdiplus" (ByVal hImage As Long, hCloneImage As Long) As Long
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal lCompositingMode As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBmp As Long, ByVal hPal As Long, hBtmap As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef nWidth As Single, ByRef nHeight As Single) As Long '
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (hImgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hImgAttr As Long, ByVal lAdjustType As Long, ByVal fAdjustEnabled As Long, clrMatrix As Any, grayMatrix As Any, ByVal lFlags As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImgAttr As Long) As Long

Private m_oCtlCancelMode        As Object
Private m_hForeBitmap           As Long

Private Function pvCreateButton(ByVal sngRadius As Single, ByVal sngBlur As Single, ByVal clrPen As Long, ByVal clrBack As Long, ByVal clrShadow As Long, hBitmap As Long) As Boolean
    Const SHADOW_OFFSET As Single = 1
    Const CLR_BLACK     As Long = &HFF000000
    Dim lIdx            As Long
    Dim lRoundWidth     As Long
    Dim lWidth          As Long
    Dim hNewBitmap      As Long
    Dim hDropShadow     As Long
    Dim hGraphics       As Long
    Dim hBrush          As Long
    
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
    If Not DrawRoundedRectangle(hGraphics, 1 + Ceil(sngBlur) + SHADOW_OFFSET / 2, 1 + Ceil(sngBlur) + SHADOW_OFFSET, _
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
    If Not DrawRoundedRectangle(hGraphics, 1 + Ceil(sngBlur), 1 + Ceil(sngBlur), lRoundWidth, lRoundWidth, sngRadius, clrPen, clrBack:=clrBack) Then
        GoTo QH
    End If
    '--- draw nine-patch markers
    If GdipCreateSolidFill(&HFFFFFFFF, hBrush) <> 0 Then
        GoTo QH
    End If
    Call GdipFillRectangleI(hGraphics, hBrush, 0, 0, lWidth, 1)
    Call GdipFillRectangleI(hGraphics, hBrush, 0, lWidth - 1, lWidth, 1)
    Call GdipFillRectangleI(hGraphics, hBrush, 0, 0, 1, lWidth)
    Call GdipFillRectangleI(hGraphics, hBrush, lWidth - 1, 0, 1, lWidth)
    lIdx = 1 + Ceil(sngBlur) + Ceil(sngRadius)
    Call GdipBitmapSetPixel(hNewBitmap, lIdx, 0, CLR_BLACK)
    Call GdipBitmapSetPixel(hNewBitmap, 0, lIdx, CLR_BLACK)
    Call GdipBitmapSetPixel(hNewBitmap, lIdx, lWidth - 1, CLR_BLACK)
    Call GdipBitmapSetPixel(hNewBitmap, lWidth - 1, lIdx, CLR_BLACK)
    '--- commit
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    hBitmap = hNewBitmap
    hNewBitmap = 0
    '--- success
    pvCreateButton = True
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
    Debug.Print "Critical error: " & Err.Description & " [pvCreateButton]"
    Resume QH
End Function

Private Function pvCloneBitmap(ByVal hBitmap As Long) As Long
    Call GdipCloneImage(hBitmap, pvCloneBitmap)
End Function

Public Function Ceil(ByVal Value As Double) As Double
    Ceil = -Int(CStr(-Value))
End Function

Private Sub Command1_Click()
    Const CLR_GREY      As Long = &HFF535353
    Const CLR_GHILIGHT  As Long = &HFF434343
    Const CLR_GBORDER   As Long = &H80838383
    Const CLR_GPRESSED  As Long = &HFF737373
    Const CLR_GSHADOW   As Long = &H40101008
    Const CLR_GSHADOW2  As Long = &H10101008
    Const CLR_BLUE      As Long = &HFF1971FE
    Const CLR_BHILIGHT  As Long = &HFF0961EE
    Const CLR_BBORDER   As Long = &H8059B1FE
    Const CLR_BPRESSED  As Long = &HFF49A1FE
    Const CLR_BSHADOW   As Long = &H40101008
    Const CLR_BSHADOW2  As Long = &H10101008
    Const CLR_LIGHT     As Long = &HFF737373
    Const CLR_LHILIGHT  As Long = &HFF636363
    Const CLR_LBORDER   As Long = &H80A3A3A3
    Const CLR_LPRESSED  As Long = &HFFB3B3B3
    Const CLR_LSHADOW   As Long = &H40101008
    Const CLR_LSHADOW2  As Long = &H10101008
    Const CLR_DARK      As Long = &HFF333333
    Const CLR_DHILIGHT  As Long = &HFF232323
    Const CLR_DBORDER   As Long = &H80535353
    Const CLR_DPRESSED  As Long = &HFF636363
    Const CLR_DSHADOW   As Long = &H40101008
    Const CLR_DSHADOW2  As Long = &H10101008
    Dim hGraphics       As Long
    Dim hNormalGrey     As Long
    Dim hHoverGrey      As Long
    Dim hPressedGrey    As Long
    Dim hNormalBlue     As Long
    Dim hHoverBlue      As Long
    Dim hPressedBlue    As Long
    Dim hNormalLight    As Long
    Dim hHoverLight     As Long
    Dim hPressedLight   As Long
    Dim hNormalDark     As Long
    Dim hHoverDark      As Long
    Dim hPressedDark    As Long
    
    If Not pvCreateButton(6, 5, CLR_GBORDER, CLR_GREY, CLR_GSHADOW, hNormalGrey) Then
        GoTo QH
    End If
    If Not pvCreateButton(6, 5, CLR_GBORDER, CLR_GHILIGHT, CLR_GSHADOW, hHoverGrey) Then
        GoTo QH
    End If
    If Not pvCreateButton(6, 5, CLR_GBORDER, CLR_GPRESSED, CLR_GSHADOW2, hPressedGrey) Then
        GoTo QH
    End If
    If Not pvCreateButton(6, 5, CLR_BBORDER, CLR_BLUE, CLR_BSHADOW, hNormalBlue) Then
        GoTo QH
    End If
    If Not pvCreateButton(6, 5, CLR_BBORDER, CLR_BHILIGHT, CLR_BSHADOW, hHoverBlue) Then
        GoTo QH
    End If
    If Not pvCreateButton(6, 5, CLR_BBORDER, CLR_BPRESSED, CLR_BSHADOW2, hPressedBlue) Then
        GoTo QH
    End If
    If Not pvCreateButton(6, 5, CLR_LBORDER, CLR_LIGHT, CLR_LSHADOW, hNormalLight) Then
        GoTo QH
    End If
    If Not pvCreateButton(6, 5, CLR_LBORDER, CLR_LHILIGHT, CLR_LSHADOW, hHoverLight) Then
        GoTo QH
    End If
    If Not pvCreateButton(6, 5, CLR_LBORDER, CLR_LPRESSED, CLR_LSHADOW2, hPressedLight) Then
        GoTo QH
    End If
    If Not pvCreateButton(6, 5, CLR_DBORDER, CLR_DARK, CLR_DSHADOW, hNormalDark) Then
        GoTo QH
    End If
    If Not pvCreateButton(6, 5, CLR_DBORDER, CLR_DHILIGHT, CLR_DSHADOW, hHoverDark) Then
        GoTo QH
    End If
    If Not pvCreateButton(6, 5, CLR_DBORDER, CLR_DPRESSED, CLR_DSHADOW2, hPressedDark) Then
        GoTo QH
    End If
    Dim lIdx            As Long
    Dim vElem           As Variant
    
    BackColor = &H182021
    lIdx = 1
    For Each vElem In Split("Q W E R T Y U I O P <= A S D F G H J K L Done ^ Z X C V B N M ! ? ^_ ?!123 _ ?!123_ keyb")
        vElem = Replace(vElem, "_", " ")
        Load ctxNineButton1(lIdx)
        ctxNineButton1(lIdx).ZOrder vbBringToFront
        If vElem = "Q" Then
            With ctxNineButton1(0)
                ctxNineButton1(lIdx).Move 80, 2268
            End With
        ElseIf vElem = "A" Then
            With ctxNineButton1(0)
                ctxNineButton1(lIdx).Move 80 + 0.4 * .Width, 2268 + .Height
            End With
        ElseIf vElem = "^" Then
            With ctxNineButton1(0)
                ctxNineButton1(lIdx).Move 80, 2268 + 2 * .Height
            End With
        ElseIf vElem = "?!123" Then
            With ctxNineButton1(0)
                ctxNineButton1(lIdx).Move 80, 2268 + 3 * .Height
            End With
        Else
            With ctxNineButton1(lIdx - 1)
                ctxNineButton1(lIdx).Move .Left + .Width, .Top ' , .Width, .Height
            End With
        End If
        With ctxNineButton1(lIdx)
            .Caption = vElem
            .Tag = vElem
            Select Case vElem
            Case "?!123", "?!123_", "keyb", "<="
                .ButtonImageBitmap(ucsBstNormal) = pvCloneBitmap(hNormalDark)
                .ButtonImageBitmap(ucsBstHover) = pvCloneBitmap(hHoverDark)
                .ButtonImageBitmap(ucsBstPressed) = pvCloneBitmap(hPressedDark)
            Case "^", "^ "
                .ButtonImageBitmap(ucsBstNormal) = pvCloneBitmap(hNormalLight)
                .ButtonImageBitmap(ucsBstHover) = pvCloneBitmap(hHoverLight)
                .ButtonImageBitmap(ucsBstPressed) = pvCloneBitmap(hPressedLight)
            Case "Done"
                .ButtonImageBitmap(ucsBstNormal) = pvCloneBitmap(hNormalBlue)
                .ButtonImageBitmap(ucsBstHover) = pvCloneBitmap(hHoverBlue)
                .ButtonImageBitmap(ucsBstPressed) = pvCloneBitmap(hPressedBlue)
            Case Else
                .ButtonImageBitmap(ucsBstNormal) = pvCloneBitmap(hNormalGrey)
                .ButtonImageBitmap(ucsBstHover) = pvCloneBitmap(hHoverGrey)
                .ButtonImageBitmap(ucsBstPressed) = pvCloneBitmap(hPressedGrey)
            End Select
            .Visible = True
            Select Case vElem
            Case "Done"
                .Width = 1.85 * .Width
            Case " "
                .Width = 6 * .Width
            Case "?!123"
                .Width = 3 * .Width
            Case "<=", "^ ", "?!123 "
                .Width = 1.25 * .Width
            End Select
        End With
        lIdx = lIdx + 1
    Next
    If GdipCreateFromHDC(hDC, hGraphics) <> 0 Then
        GoTo QH
    End If
    If GdipDrawImageRectRect(hGraphics, hNormalGrey, 0, 0, 500, 500, 0, 0, 500, 500) <> 0 Then
        GoTo QH
    End If
    If GdipDrawImageRectRect(hGraphics, hHoverGrey, 0, 50, 500, 500, 0, 0, 500, 500) <> 0 Then
        GoTo QH
    End If
    If GdipDrawImageRectRect(hGraphics, hPressedGrey, 0, 100, 500, 500, 0, 0, 500, 500) <> 0 Then
        GoTo QH
    End If
QH:
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    Call GdipDisposeImage(hNormalGrey)
    Call GdipDisposeImage(hHoverGrey)
    Call GdipDisposeImage(hPressedGrey)
    Call GdipDisposeImage(hNormalBlue)
    Call GdipDisposeImage(hHoverBlue)
    Call GdipDisposeImage(hPressedBlue)
    Call GdipDisposeImage(hNormalLight)
    Call GdipDisposeImage(hHoverLight)
    Call GdipDisposeImage(hPressedLight)
End Sub

Private Sub Command2_Click()
    Dim hBrush  As Long
    Dim hGraphics   As Long
    Dim lWidth      As Long
    Dim lHeight     As Long
    Dim lIdx        As Long
    Dim hBitmap     As Long
    Dim sngWidth    As Single
    Dim sngHeight   As Single
    Dim hAttributes As Long
    
    BackColor = &H182021
    lWidth = ScaleWidth \ Screen.TwipsPerPixelX
    lHeight = ScaleHeight \ Screen.TwipsPerPixelX
    If m_hForeBitmap <> 0 Then
        Call GdipDisposeImage(m_hForeBitmap)
        m_hForeBitmap = 0
    End If
    If GdipCreateBitmapFromScan0(lWidth, lHeight, lWidth * 4, PixelFormat32bppARGB, 0, m_hForeBitmap) <> 0 Then
        GoTo QH
    End If
    If GdipGetImageGraphicsContext(m_hForeBitmap, hGraphics) <> 0 Then
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
'    If Not BlurBitmap(m_hForeBitmap, Sqr(lIdx)) Then
'        GoTo QH
'    End If
    Call GdipDeleteGraphics(hGraphics)
    If GdipCreateFromHDC(hDC, hGraphics) <> 0 Then
        GoTo QH
    End If
    If GdipDrawImageRectRect(hGraphics, m_hForeBitmap, 0, 0, lWidth, lHeight, 0, 0, lWidth, lHeight) <> 0 Then
        GoTo QH
    End If
    For lIdx = ctxNineButton1.LBound To ctxNineButton1.UBound
        ctxNineButton1(lIdx).Repaint
    Next
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
End Sub

Private Sub ctxNineButton1_OwnerDraw(Index As Integer, ByVal hGraphics As Long, ByVal hFont As Long, ByVal ButtonState As UcsNineButtonStateEnum, ClientLeft As Long, ClientTop As Long, ClientWidth As Long, ClientHeight As Long, Caption As String, ByVal hPicture As Long)
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    
    If m_hForeBitmap = 0 Then
        GoTo QH
    End If
    With ctxNineButton1(Index)
        lLeft = .Left \ Screen.TwipsPerPixelX
        lTop = .Top \ Screen.TwipsPerPixelY
        lWidth = .Width \ Screen.TwipsPerPixelX
        lHeight = .Height \ Screen.TwipsPerPixelX
    End With
    If GdipDrawImageRectRect(hGraphics, m_hForeBitmap, 0, 0, lWidth, lHeight, lLeft, lTop, lWidth, lHeight) <> 0 Then
        GoTo QH
    End If
QH:
End Sub

Private Sub Form_Load()
    StartGdip
End Sub

Private Sub Form_Click()
    Const Width         As Long = 500 ' 1920
    Const Height        As Long = 500 ' 1200
    Const BurlyWood = &HFFDEB887
    Dim hBitmap         As Long
    Dim hGraphics       As Long
    Dim dblTimer        As Double
    
    If GdipCreateBitmapFromScan0(Width, Height, Width * 4, PixelFormat32bppARGB, 0, hBitmap) <> 0 Then
        GoTo QH
    End If
    '--- setup graphics
    If GdipGetImageGraphicsContext(hBitmap, hGraphics) <> 0 Then
        GoTo QH
    End If
    DrawRoundedRectangle hGraphics, 20 + 0.5, 20 + 1, 400, 400, 5, &H40101008, clrBack:=&H40101008
    dblTimer = Timer
    If Not BlurBitmap(hBitmap, 6) Then
        GoTo QH
    End If
    Caption = Format$(Timer - dblTimer, "0.000")
    DrawRoundedRectangle hGraphics, 20, 20, 400, 400, 5, BurlyWood, clrBack:=BurlyWood
    Call GdipDeleteGraphics(hGraphics)
    If GdipCreateFromHDC(hDC, hGraphics) <> 0 Then
        GoTo QH
    End If
    If GdipDrawImageRectRect(hGraphics, hBitmap, 0, 0, 500, 500, 0, 0, 500, 500) <> 0 Then
        GoTo QH
    End If
QH:
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
End Sub

Public Sub RegisterCancelMode(oCtl As Object)
    If Not m_oCtlCancelMode Is Nothing And Not m_oCtlCancelMode Is oCtl Then
        m_oCtlCancelMode.CancelMode
    End If
    Set m_oCtlCancelMode = oCtl
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_oCtlCancelMode Is Nothing Then
        m_oCtlCancelMode.CancelMode
        Set m_oCtlCancelMode = Nothing
    End If
End Sub

Private Function pvPrepareAttribs(ByVal sngAlpha As Single, hAttributes As Long) As Boolean
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
    Resume QH
End Function
