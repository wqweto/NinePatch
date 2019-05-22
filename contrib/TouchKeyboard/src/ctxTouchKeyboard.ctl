VERSION 5.00
Begin VB.UserControl ctxTouchKeyboard 
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
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
#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)
#Const ImplSelfContained = True

'=========================================================================
' Public Events
'=========================================================================

Event ButtonClick(ByVal Index As Long)
Event ButtonMouseDown(ByVal Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event ButtonMouseMove(ByVal Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event ButtonMouseUp(ByVal Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
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
'--- for SystemParametersInfo
Private Const SPI_GETKEYBOARDSPEED          As Long = 10
Private Const SPI_GETKEYBOARDDELAY          As Long = 22
'--- for thunks
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
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
#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If
#If Not ImplUseShared Then
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

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const DEF_LAYOUT1           As String = "q w e r t y u i o p <=|1.25|D " & _
                                                "|0.5|S|N a s d f g h j k l Done|1.85|B " & _
                                                "^^|||N z x c v b n m ! ? ^^|1.25| " & _
                                                "?!123|3|D|N _|6 ?!123|1.25|D keyb||D"
Private Const DEF_FORECOLOR         As Long = vbWindowBackground
Private Const DEF_ENABLED           As Boolean = True
Private Const DEF_USEFOREBITMAP     As Boolean = True
Private Const DEF_KEYSALLOWREPEAT   As String = "<="
Private Const BUTTON_RADIUS         As Single = 6
Private Const BUTTON_BLUR           As Single = 5

Private m_pTimer                As IUnknown
Private m_lRepeatIndex          As Long
'--- design-time
Private m_clrFore               As OLE_COLOR
Private WithEvents m_oFont      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_sLayout               As String
Private m_bUseForeBitmap        As Boolean
Private m_sKeysAllowRepeat      As String
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
    Debug.Print "Critical error: " & Err.Description & " [" & STR_MODULE_NAME & "." & sFunction & "]", Timer
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
        Repaint
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
        Repaint
        PropertyChanged
    End If
End Property

Property Get UseForeBitmap() As Boolean
    UseForeBitmap = m_bUseForeBitmap
End Property

Property Let UseForeBitmap(ByVal bValue As Boolean)
    If m_bUseForeBitmap <> bValue Then
        m_bUseForeBitmap = bValue
        If bValue Then
            pvPrepareForeground m_hForeBitmap
        ElseIf m_hForeBitmap <> 0 Then
            Call GdipDisposeImage(m_hForeBitmap)
            m_hForeBitmap = 0
        End If
        Repaint
        PropertyChanged
    End If
End Property

Property Get KeysAllowRepeat() As String
    KeysAllowRepeat = m_sKeysAllowRepeat
End Property

Property Let KeysAllowRepeat(sValue As String)
    m_sKeysAllowRepeat = sValue
End Property

'= run-time ==============================================================

Property Get ButtonCaption(ByVal Index As Long) As String
    ButtonCaption = btn(Index).Caption
End Property

Property Get ButtonTag(ByVal Index As Long) As String
    ButtonTag = btn(Index).Tag
End Property

Private Property Get pvAddressOfTimerProc() As ctxTouchKeyboard
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
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

Public Sub Refresh()
    UserControl.Refresh
End Sub

Public Sub Repaint()
    Dim lIdx            As Long
    
    If m_bShown Then
        For lIdx = 1 To m_lButtonCurrent
            btn(lIdx).Repaint
        Next
        UserControl.Refresh
'        Call ApiUpdateWindow(ContainerHwnd) '--- pump WM_PAINT
    End If
End Sub

Public Function TimerProc() As Long
Attribute TimerProc.VB_MemberFlags = "40"
    Const FUNC_NAME     As String = "TimerProc"
    
    On Error GoTo EH
    If m_lRepeatIndex <> 0 Then
        RaiseEvent ButtonClick(m_lRepeatIndex)
        If m_lRepeatIndex <> 0 Then
            Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc, pvGetKeyboardSpeed)
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

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
                    MoveCtl btn(lIdx), 0, 0, btn(0).Width, btn(0).Width
                    lRow = 0
                Case "N"
                    MoveCtl btn(lIdx), 0, .Top + .Height, btn(0).Width, btn(0).Width
                    lRow = lRow + 1
                    ReDim Preserve m_cButtonRows(0 To lRow) As Collection
                    Set m_cButtonRows(lRow) = New Collection
                Case Else
                    MoveCtl btn(lIdx), .Left + .Width, .Top, btn(0).Width, btn(0).Width
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
    Dim dblCurrent      As Double
    Dim dblTotal        As Double
    Dim dblLeft         As Double
    Dim dblTop          As Double
    
    On Error GoTo EH
    For lIdx = 0 To UBound(m_cButtonRows)
        dblTotal = 0
        For Each vElem In m_cButtonRows(lIdx)
            dblTotal = dblTotal + vElem(1)
        Next
        dblCurrent = 0
        For Each vElem In m_cButtonRows(lIdx)
            dblLeft = AlignOrigTwipsToPix(dblCurrent * ScaleWidth / dblTotal)
            dblTop = AlignOrigTwipsToPix(lIdx * ScaleHeight / (UBound(m_cButtonRows) + 1))
            MoveCtl btn(vElem(0)), dblLeft, dblTop, _
                AlignOrigTwipsToPix((dblCurrent + vElem(1)) * ScaleWidth / dblTotal - dblLeft), _
                AlignOrigTwipsToPix((lIdx + 1) * ScaleHeight / (UBound(m_cButtonRows) + 1) - dblTop)
            dblCurrent = dblCurrent + vElem(1)
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
    Dim sngRadius       As Single
    Dim sngBlur         As Single
    
    On Error GoTo EH
    If clrBack <> Transparent Then
        sngRadius = IconScale(BUTTON_RADIUS)
        sngBlur = IconScale(BUTTON_BLUR)
        If SearchCollection(m_cButtonImageCache, "N" & Hex(clrBack)) Then
            hNormalBitmap = m_cButtonImageCache.Item("N" & Hex(clrBack))
            hHoverBitmap = m_cButtonImageCache.Item("H" & Hex(clrBack))
            hPressedBitmap = m_cButtonImageCache.Item("P" & Hex(clrBack))
        Else
            clrBorder = GdipAdjustColor(clrBack, AdjustBri:=0.3, AdjustAlpha:=-0.5)
            clrShadow = GdipAdjustColor(clrBack, AdjustBri:=-0.8, AdjustAlpha:=-0.75)
            If Not GdipPrepareButtonBitmap(sngRadius, sngBlur, clrBorder, clrBack, clrShadow, hNormalBitmap) Then
                GoTo QH
            End If
            If Not GdipPrepareButtonBitmap(sngRadius, sngBlur, clrBorder, GdipAdjustColor(clrBack, AdjustBri:=-0.2), clrShadow, hHoverBitmap) Then
                GoTo QH
            End If
            clrShadow = GdipAdjustColor(clrShadow, AdjustAlpha:=-0.75)
            If Not GdipPrepareButtonBitmap(sngRadius, sngBlur, clrBorder, GdipAdjustColor(clrBack, AdjustBri:=0.25, AdjustSat:=-0.25), clrShadow, hPressedBitmap) Then
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
    lWidth = Int(ScaleWidth / OrigTwipsPerPixelX + 0.5)
    lHeight = Int(ScaleHeight / OrigTwipsPerPixelX + 0.5)
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
        GdipPreparePrivateFont LocateFile(PathCombine(App.Path, "External\fa-regular-400.ttf")), m_oFont.Size, m_hAwesomeRegular, m_hAwesomeColRegular
        GdipPreparePrivateFont LocateFile(PathCombine(App.Path, "External\fa-solid-900.ttf")), m_oFont.Size, m_hAwesomeSolid, m_hAwesomeColSolid
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

Private Function pvIsRepeatKey(ByVal Index As Long) As Boolean
    If LenB(m_sKeysAllowRepeat) <> 0 Then
        If InStr("|" & btn.Item(Index).Caption & "|", "|" & m_sKeysAllowRepeat & "|") > 0 Then
            pvIsRepeatKey = True
        End If
    End If
End Function

Private Function pvGetKeyboardDelay() As Long
    Dim lValue          As Long
    
    Call SystemParametersInfo(SPI_GETKEYBOARDDELAY, 0, lValue, 0)
    If lValue < 0 Or lValue > 3 Then
        lValue = 0
    End If
    pvGetKeyboardDelay = (lValue + 1) * 250
End Function

Private Function pvGetKeyboardSpeed() As Long
    Dim lValue          As Long
    
    Call SystemParametersInfo(SPI_GETKEYBOARDSPEED, 0, lValue, 0)
    If lValue < 0 Or lValue > 29 Then
        lValue = 29
    End If
    pvGetKeyboardSpeed = CSng(31 - lValue) * (400 - 1000! / 30) / 31 + 1000! / 300
End Function

#If Not ImplUseShared Then
Private Function At(Data As Variant, ByVal Index As Long, Optional Default As String) As String
    On Error GoTo QH
    At = Default
    If LBound(Data) <= Index And Index <= UBound(Data) Then
        At = CStr(Data(Index))
    End If
QH:
End Function

Private Function AlignOrigTwipsToPix(ByVal dblTwips As Double) As Double
    AlignOrigTwipsToPix = Int(dblTwips / OrigTwipsPerPixelX + 0.5) * OrigTwipsPerPixelX
End Function

Private Function IconScale(ByVal sngSize As Single) As Long
    Select Case ScreenTwipsPerPixelX
    Case Is < 6.5
        IconScale = Int(sngSize * 3)
    Case Is < 9.5
        IconScale = Int(sngSize * 2)
    Case Is < 11.5
        IconScale = Int(sngSize * 3 \ 2)
    Case Else
        IconScale = Int(sngSize * 1)
    End Select
End Function

Private Sub MoveCtl(oCtl As Object, ByVal Left As Single, ByVal Top As Variant, ByVal Width As Variant, ByVal Height As Variant)
    If 1440 \ ScreenTwipsPerPixelX = 1440 / ScreenTwipsPerPixelX Then
        oCtl.Move Left, Top, Width, Height
    Else
        oCtl.Move Left + ScreenTwipsPerPixelX, Top, Width, Height
        oCtl.Move Left
    End If
End Sub

Private Function LocateFile(sFile As String) As String
    LocateFile = sFile
End Function

Private Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Private Property Get ScreenTwipsPerPixelX() As Single
    ScreenTwipsPerPixelX = Screen.TwipsPerPixelX
End Property

Private Property Get ScreenTwipsPerPixelY() As Single
    ScreenTwipsPerPixelY = Screen.TwipsPerPixelY
End Property

Private Property Get OrigTwipsPerPixelX() As Single
    OrigTwipsPerPixelX = Screen.TwipsPerPixelX
End Property

Private Property Get OrigTwipsPerPixelY() As Single
    OrigTwipsPerPixelY = Screen.TwipsPerPixelY
End Property

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
    Call GetEnvironmentVariable("_MST_GLOBAL" & App.hInstance & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & App.hInstance & "_" & sKey, lValue)
End Property
#End If ' Not ImplUseShared

'=========================================================================
' Control events
'=========================================================================

Private Sub btn_Click(Index As Integer)
    Const FUNC_NAME     As String = "btn_Click"
    
    On Error GoTo EH
    Set m_pTimer = Nothing
    m_lRepeatIndex = 0
    If Not pvIsRepeatKey(Index) Then
        RaiseEvent ButtonClick(Index)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub btn_DblClick(Index As Integer)
    Const FUNC_NAME     As String = "btn_DblClick"
    
    On Error GoTo EH
    Set m_pTimer = Nothing
    m_lRepeatIndex = 0
    If pvIsRepeatKey(Index) Then
        RaiseEvent ButtonClick(Index)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub btn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "btn_MouseDown"
    
    On Error GoTo EH
    Set m_pTimer = Nothing
    m_lRepeatIndex = 0
    RaiseEvent ButtonMouseDown(Index, Button, Shift, X, Y)
    If pvIsRepeatKey(Index) Then
        RaiseEvent ButtonClick(Index)
        m_lRepeatIndex = Index
        If m_lRepeatIndex <> 0 Then
            Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc, pvGetKeyboardDelay)
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub btn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent ButtonMouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub btn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "btn_MouseUp"
    
    On Error GoTo EH
    Set m_pTimer = Nothing
    m_lRepeatIndex = 0
    RaiseEvent ButtonMouseUp(Index, Button, Shift, X, Y)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub btn_OwnerDraw(Index As Integer, ByVal hGraphics As Long, ByVal hFont As Long, ByVal ButtonState As UcsNineButtonStateEnum, ClientLeft As Long, ClientTop As Long, ClientWidth As Long, ClientHeight As Long, Caption As String, ByVal hPicture As Long)
    Const FUNC_NAME     As String = "btn_OwnerDraw"
    Const FA_ARROW_ALT_CIRCLE_UP As Long = &HF35B&
    Const FA_BACKSPACE  As Long = &HF55A&
    Const FA_KEYBOARD   As Long = &HF11C&
    Const FA_CHECK_CIRCLE As Long = &HF058&
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim hStringFormat   As Long
    Dim hBrush          As Long
    Dim uRect           As RECTF
    Dim bShift          As Boolean
    Dim sText           As String
    Dim hTextFont       As Long
    
    On Error GoTo EH
    With btn(Index)
        lLeft = .Left \ ScreenTwipsPerPixelX
        lTop = .Top \ ScreenTwipsPerPixelY
        lWidth = Int(.Width / OrigTwipsPerPixelX + 0.5)
        lHeight = Int(.Height / OrigTwipsPerPixelY + 0.5)
    End With
    If m_hAwesomeRegular <> 0 And m_hAwesomeSolid <> 0 Then
        With uRect
            .Left = ClientLeft + IconScale(BUTTON_RADIUS)
            .Top = ClientTop + IconScale(BUTTON_RADIUS)
            .Right = ClientWidth - 2 * IconScale(BUTTON_RADIUS)
            .Bottom = ClientHeight - 2 * IconScale(BUTTON_RADIUS)
        End With
        Select Case Caption
        Case "^^"
            bShift = InStr(btn(Index).Tag, "|L") > 0
            If (ButtonState And ucsBstPressed) <> 0 Then
                bShift = Not bShift
            End If
            sText = ChrW$(FA_ARROW_ALT_CIRCLE_UP)
            hTextFont = IIf(bShift, m_hAwesomeSolid, m_hAwesomeRegular)
        Case "<="
            sText = ChrW$(FA_BACKSPACE)
            hTextFont = m_hAwesomeSolid
        Case "keyb"
            sText = ChrW$(FA_KEYBOARD)
            hTextFont = m_hAwesomeRegular
        Case "Done", "Готово"
            If uRect.Right < uRect.Bottom * 1.2 Then
                sText = ChrW$(FA_CHECK_CIRCLE)
                hTextFont = m_hAwesomeRegular
            End If
        End Select
        If LenB(sText) <> 0 Then
            If Not GdipPrepareStringFormat(ucsBflCenter, hStringFormat) Then
                GoTo QH
            End If
            If GdipCreateSolidFill(GdipTranslateColor(m_clrFore), hBrush) <> 0 Then
                GoTo QH
            End If
            If GdipDrawString(hGraphics, StrPtr(sText), -1, hTextFont, uRect, hStringFormat, hBrush) <> 0 Then
                GoTo QH
            End If
            Caption = vbNullString
        End If
    End If
    If m_hForeBitmap <> 0 Then
        If GdipDrawImageRectRect(hGraphics, m_hForeBitmap, 0, 0, lWidth, lHeight, lLeft, lTop, lWidth, lHeight) <> 0 Then
            GoTo QH
        End If
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
    UseForeBitmap = DEF_USEFOREBITMAP
    KeysAllowRepeat = DEF_KEYSALLOWREPEAT
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
        UseForeBitmap = .ReadProperty("UseForeBitmap", DEF_USEFOREBITMAP)
        KeysAllowRepeat = .ReadProperty("KeysAllowRepeat", DEF_KEYSALLOWREPEAT)
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
        Call .WriteProperty("UseForeBitmap", UseForeBitmap, DEF_USEFOREBITMAP)
        Call .WriteProperty("KeysAllowRepeat", KeysAllowRepeat, DEF_KEYSALLOWREPEAT)
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_Resize()
    Const FUNC_NAME     As String = "UserControl_Resize"
    
    On Error GoTo EH
    If m_bUseForeBitmap Then
        pvPrepareForeground m_hForeBitmap
    End If
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
    
    Set m_pTimer = Nothing
    m_lRepeatIndex = 0
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
    If m_hForeBitmap <> 0 Then
        Call GdipDisposeImage(m_hForeBitmap)
        m_hForeBitmap = 0
    End If
    #If DebugMode Then
        DebugInstanceTerm MODULE_NAME, m_sDebugID
    #End If
End Sub
