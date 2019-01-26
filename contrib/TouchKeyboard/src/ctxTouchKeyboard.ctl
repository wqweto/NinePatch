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
'--- for GdipCreateBitmapFromScan0
Private Const PixelFormat32bppARGB          As Long = &H26200A
'--- for GdipSetCompositingMode
Private Const CompositingModeSourceCopy     As Long = 1
'--- GDI+ colors
Private Const Transparent                   As Long = &HFFFFFF

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function ApiUpdateWindow Lib "user32" Alias "UpdateWindow" (ByVal hWnd As Long) As Long
'--- gdi+
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, pInputBuf As Any, Optional ByVal pOutputBuf As Long = 0) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal hImage As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal lWidth As Long, ByVal lHeight As Long, ByVal lStride As Long, ByVal lPixelFormat As Long, ByVal Scan0 As Long, hBitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As Long
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal srcUnit As Long = UnitPixel, Optional ByVal hImageAttributes As Long, Optional ByVal pfnCallback As Long, Optional ByVal lCallbackData As Long) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal lArgb As Long, hBrush As Long) As Long
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal lCompositingMode As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBmp As Long, ByVal hPal As Long, hBtmap As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef nWidth As Single, ByRef nHeight As Single) As Long '
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (hImgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hImgAttr As Long, ByVal lAdjustType As Long, ByVal fAdjustEnabled As Long, clrMatrix As Any, grayMatrix As Any, ByVal lFlags As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImgAttr As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal hGraphics As Long, ByVal lStrPtr As Long, ByVal lLength As Long, ByVal hFont As Long, uRect As RECTF, ByVal hStringFormat As Long, ByVal hBrush As Long) As Long

Private Type RECTF
   Left                 As Single
   Top                  As Single
   Right                As Single
   Bottom               As Single
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const DEF_LAYOUT1           As String = "q w e r t y u i o p <=|1.25|D " & _
                                                "|0.5|S|N a s d f g h j k l Done|1.85|B " & _
                                                "^^|||N z x c v b n m ! ? ^^|1.25| " & _
                                                "?!123|3|D|N _|6 ?!123|1.25|D keyb||D"
Private Const DEF_FORECOLOR         As Long = vbWindowBackground
Private Const DEF_ENABLED           As Boolean = True

Private m_clrFore               As OLE_COLOR
Private WithEvents m_oFont      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_sLayout               As String
'--- run-time
Private m_lButtonCurrent        As Long
Private m_cButtonRows()         As Collection
Private m_oCtlCancelMode        As Object
Private m_hForeBitmap           As Long
Private m_cButtonImageCache     As Collection
Private m_bShown                As Boolean
Private m_hAwesomeRegular       As Long
Private m_hAwesomeColRegular    As Long
Private m_hAwesomeSolid         As Long
Private m_hAwesomeColSolid      As Long
'--- debug
Private m_sInstanceName         As String
#If DebugMode Then
    Private m_sDebugID          As String
#End If

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
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]", Timer
#End If
End Function

'Private Function RaiseError(sFunction As String) As VbMsgBoxResult
'    Err.Raise Err.Number, STR_MODULE_NAME & "." & sFunction & vbCrLf & Err.Source, Err.Description
'End Function

'=========================================================================
' Properties
'=========================================================================

Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_clrFore
End Property

Property Let ForeColor(ByVal clrValue As OLE_COLOR)
    If m_clrFore <> clrValue Then
        m_clrFore = clrValue
        pvLoadLayout m_sLayout
        pvSizeLayout
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
        pvLoadLayout m_sLayout
        pvSizeLayout
        pvPrepareFontAwesome
        PropertyChanged
    End If
End Property

Property Get Layout() As String
    Layout = m_sLayout
End Property

Property Let Layout(sValue As String)
    If m_sLayout <> sValue Then
        m_sLayout = sValue
        pvLoadLayout sValue
        pvSizeLayout
        Repaint
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
    End If
    PropertyChanged
End Property

'= run-time ==============================================================

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
    pvParentRegisterCancelMode Me
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

Public Sub Repaint()
    If m_bShown Then
        UserControl.Refresh
        Call ApiUpdateWindow(ContainerHwnd) '--- pump WM_PAINT
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
        If vElem = "|" Then
            vSplit = Array(vElem)
        Else
            vSplit = Split(vElem, "|")
        End If
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
                    .Width = Val(At(vSplit, 1)) * .Width
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
                .Move AlignTwipsToPix(dblLeft * ScaleWidth / dblTotal), AlignTwipsToPix(lIdx * ScaleHeight / (UBound(m_cButtonRows) + 1))
                dblLeft = dblLeft + vElem(1)
                .Width = AlignTwipsToPix(dblLeft * ScaleWidth / dblTotal) - .Left
                .Height = AlignTwipsToPix((lIdx + 1) * ScaleHeight / (UBound(m_cButtonRows) + 1)) - .Top
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
            clrBorder = GdipAdjustColor(clrBack, AdjustBri:=0.3, AdjustAlpha:=-0.5)
            clrShadow = GdipAdjustColor(clrBack, AdjustBri:=-0.8, AdjustAlpha:=-0.75)
            If Not GdipPrepareButtonBitmap(6, 5, clrBorder, clrBack, clrShadow, hNormalBitmap) Then
                GoTo QH
            End If
            If Not GdipPrepareButtonBitmap(6, 5, clrBorder, GdipAdjustColor(clrBack, AdjustBri:=-0.2), clrShadow, hHoverBitmap) Then
                GoTo QH
            End If
            clrShadow = GdipAdjustColor(clrShadow, AdjustAlpha:=-0.75)
            If Not GdipPrepareButtonBitmap(6, 5, clrBorder, GdipAdjustColor(clrBack, AdjustBri:=0.25, AdjustSat:=-0.25), clrShadow, hPressedBitmap) Then
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
        .ForeColor = m_clrFore
        Set .Font = m_oFont
        .ButtonImageBitmap(ucsBstNormal) = hNormalBitmap
        .ButtonImageBitmap(ucsBstHover) = hHoverBitmap
        .ButtonImageBitmap(ucsBstPressed) = hPressedBitmap
    End With
QH:
    Exit Function
EH:
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
    lHeight = ScaleHeight \ Screen.TwipsPerPixelX
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
    If Not pvPrepareAttribs(0.15, hAttributes) Then
        GoTo QH
    End If
    If GdipDrawImageRectRect(hGraphics, hBitmap, lWidth * -0.25, lHeight * -0.25, lWidth * 1.5, lHeight * 1.5, 0, 0, sngWidth, sngHeight, , hAttributes) <> 0 Then
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

Private Sub pvPrepareFontAwesome()
    If Not m_oFont Is Nothing Then
        GdipPreparePrivateFont App.Path & "\fa-regular-400.ttf", m_oFont.Size, m_hAwesomeRegular, m_hAwesomeColRegular
        GdipPreparePrivateFont App.Path & "\fa-solid-900.ttf", m_oFont.Size, m_hAwesomeSolid, m_hAwesomeColSolid
    End If
End Sub

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

Private Sub btn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent ButtonMouseDown(Index)
End Sub

Private Sub btn_OwnerDraw(Index As Integer, ByVal hGraphics As Long, ByVal hFont As Long, ByVal ButtonState As UcsNineButtonStateEnum, ClientLeft As Long, ClientTop As Long, ClientWidth As Long, ClientHeight As Long, Caption As String, ByVal hPicture As Long)
    Const FUNC_NAME     As String = "btn_OwnerDraw"
    Const FA_ARROW_ALT_CIRCLE_UP As Long = &HF35B&
    Const FA_BACKSPACE  As Long = &HF55A&
    Const FA_KEYBOARD   As Long = &HF11C&
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim hStringFormat   As Long
    Dim hBrush          As Long
    Dim uRect           As RECTF
    Dim bShift          As Boolean
    
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
    If m_hAwesomeRegular <> 0 And m_hAwesomeSolid <> 0 Then
        If Not GdipPrepareStringFormat(ucsBflCenter, hStringFormat) Then
            GoTo QH
        End If
        If GdipCreateSolidFill(GdipTranslateColor(m_clrFore), hBrush) <> 0 Then
            GoTo QH
        End If
        uRect.Right = lWidth
        uRect.Bottom = lHeight
        Select Case Caption
        Case "^^"
            bShift = InStr(btn(Index).Tag, "|L") > 0
            If (ButtonState And ucsBstPressed) <> 0 Then
                bShift = Not bShift
            End If
            If GdipDrawString(hGraphics, StrPtr(ChrW(FA_ARROW_ALT_CIRCLE_UP)), -1, IIf(bShift, m_hAwesomeSolid, m_hAwesomeRegular), uRect, hStringFormat, hBrush) <> 0 Then
                GoTo QH
            End If
            Caption = vbNullString
        Case "<="
            If GdipDrawString(hGraphics, StrPtr(ChrW(FA_BACKSPACE)), -1, m_hAwesomeSolid, uRect, hStringFormat, hBrush) <> 0 Then
                GoTo QH
            End If
            Caption = vbNullString
        Case "keyb"
            If GdipDrawString(hGraphics, StrPtr(ChrW(FA_KEYBOARD)), -1, m_hAwesomeRegular, uRect, hStringFormat, hBrush) <> 0 Then
                GoTo QH
            End If
            Caption = vbNullString
        End Select
    End If
    If GdipDrawImageRectRect(hGraphics, m_hForeBitmap, 0, 0, lWidth, lHeight, lLeft, lTop, lWidth, lHeight) <> 0 Then
        GoTo QH
    End If
QH:
    If hStringFormat <> 0 Then
        Call GdipDeleteStringFormat(hStringFormat)
    End If
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub btn_RegisterCancelMode(Index As Integer, oCtl As Object, Handled As Boolean)
    Const FUNC_NAME     As String = "btn_RegisterCancelMode"
    
    On Error GoTo EH
    pvParentRegisterCancelMode Me
    If Not m_oCtlCancelMode Is Nothing And Not m_oCtlCancelMode Is oCtl Then
        m_oCtlCancelMode.CancelMode
    End If
    Set m_oCtlCancelMode = oCtl
    Handled = True
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    Const FUNC_NAME     As String = "m_oFont_FontChanged"
    
    On Error GoTo EH
    pvLoadLayout m_sLayout
    pvSizeLayout
    pvPrepareFontAwesome
    PropertyChanged
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "UserControl_MouseMove"
    
    On Error GoTo EH
    CancelMode
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_InitProperties()
    Const FUNC_NAME     As String = "UserControl_InitProperties"
    
    On Error GoTo EH
    ForeColor = DEF_FORECOLOR
    Set Font = Ambient.Font
    Layout = DEF_LAYOUT1
    Enabled = DEF_ENABLED
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
    If Ambient.UserMode Then
        If IsCompileTime(Extender) Then
            Exit Sub
        End If
    End If
    With PropBag
        m_clrFore = .ReadProperty("ForeColor", DEF_FORECOLOR)
        Set m_oFont = .ReadProperty("Font", Ambient.Font)
        pvPrepareFontAwesome
        Layout = .ReadProperty("Layout", DEF_LAYOUT1)
        Enabled = .ReadProperty("Enabled", DEF_ENABLED)
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
        Call .WriteProperty("ForeColor", ForeColor, DEF_FORECOLOR)
        Call .WriteProperty("Font", Font, Ambient.Font)
        Call .WriteProperty("Layout", Layout, DEF_LAYOUT1)
        Call .WriteProperty("Enabled", Enabled, DEF_ENABLED)
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_Resize()
    Const FUNC_NAME     As String = "UserControl_Resize"
    
    On Error GoTo EH
    pvPrepareForeground m_hForeBitmap
    pvSizeLayout
    Repaint
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_Show()
    Const FUNC_NAME     As String = "UserControl_Show"
    
    On Error GoTo EH
    If Not m_bShown Then
        m_bShown = True
        Repaint
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_Hide()
    m_bShown = False
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
    Set m_cButtonImageCache = New Collection
    ReDim m_cButtonRows(0 To 0) As Collection
    Set m_cButtonRows(0) = New Collection
End Sub

Private Sub UserControl_Terminate()
    Dim vElem           As Variant
    
    For Each vElem In m_cButtonImageCache
        Call GdipDisposeImage(vElem)
    Next
    If m_hAwesomeRegular <> 0 Then
        Call GdipDeleteFont(m_hAwesomeRegular)
        m_hAwesomeRegular = 0
    End If
    If m_hAwesomeColRegular <> 0 Then
        Call GdipDeletePrivateFontCollection(m_hAwesomeColRegular)
        m_hAwesomeColRegular = 0
    End If
    If m_hAwesomeSolid <> 0 Then
        Call GdipDeleteFont(m_hAwesomeSolid)
        m_hAwesomeSolid = 0
    End If
    If m_hAwesomeColSolid <> 0 Then
        Call GdipDeletePrivateFontCollection(m_hAwesomeColSolid)
        m_hAwesomeColSolid = 0
    End If
    #If DebugMode Then
        DebugInstanceTerm MODULE_NAME, m_sDebugID
    #End If
End Sub

