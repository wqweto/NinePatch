Attribute VB_Name = "mdTouchKeyboard"
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "mdTouchKeyboard"

#Const ImplUseShared = NPPNG_USE_SHARED <> 0

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsTouchButtonTextFlagsEnum
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
Private Const UnitPoint                     As Long = 3

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'--- gdi+
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, lpRect As Any, ByVal lFlags As Long, ByVal lPixelFormat As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal hFormatAttributes As Long, ByVal nLanguage As Integer, hStringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal hStringFormat As Long, ByVal lFlags As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal hStringFormat As Long, ByVal eAlign As StringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal hStringFormat As Long, ByVal eAlign As StringAlignment) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal lColor As Long, hBrush As Long) As Long
Private Declare Function GdipCreateFontFromDC Lib "gdiplus" (ByVal hDC As Long, hCreatedFont As Long) As Long
Private Declare Function GdipNewPrivateFontCollection Lib "gdiplus" (hFontCollection As Long) As Long
Private Declare Function GdipPrivateAddFontFile Lib "gdiplus" (ByVal hFontCollection As Long, ByVal lpFileName As Long) As Long
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal hFontFamily As Long, ByVal emSize As Single, ByVal lStyle As Long, ByVal lUnit As Long, hFont As Long) As Long
Private Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal hFontCollection As Long, ByVal lNumSought As Long, aFamilies As Any, lNumFound As Long) As Long
'--- public
Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal hStringFormat As Long) As Long
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal hFont As Long) As Long
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As Long
Public Declare Function GdipDeletePrivateFontCollection Lib "gdiplus" (hFontCollection As Long) As Long
#If Not ImplUseShared Then
    Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
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
    Debug.Print Err.Description & " [" & STR_MODULE_NAME & "." & sFunction & "]", Timer
End Sub

'Private Function RaiseError(sFunction As String) As VbMsgBoxResult
'    Err.Raise Err.Number, STR_MODULE_NAME & "." & sFunction & vbCrLf & Err.Source, Err.Description
'End Function

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
    dblTemp = Sqr(-(dblRadius * dblRadius) / (2 * Log(1# / 255#)))
    If dblTemp <= 0 Then
        dblTemp = 0.01
    End If
    dblTemp = dblTemp * (1# + (0.3165 * NUM_ITERS + 0.5695) / ((NUM_ITERS + 0.7818) * (NUM_ITERS + 0.7818)))
    dblTemp = (dblTemp * dblTemp) / (2# * NUM_ITERS)
    dblNu = (1# + 2# * dblTemp - Sqr(1# + 4# * dblTemp)) / (2# * dblTemp)
    dblBndryScale = (1# / (1# - dblNu))
    dblPostScale = ((dblNu / dblTemp) ^ (2# * NUM_ITERS)) * 255#
    ' Copy the contents of the incoming byte array into the double array buffer
    LoadSave dblBuffer(0, 0), 1# / 255#, lpBits + (lTop * lStride + lLeft) * 4 + lChannel, lStride, lWidth, lHeight, 0
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
    Call CryptStringToBinary(StrPtr(sThunkStr), Len(sThunkStr), CRYPT_STRING_BASE64, 0, lThunkSize, 0, 0)
    lThunkPtr = VirtualAlloc(0, lThunkSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CryptStringToBinary(StrPtr(sThunkStr), Len(sThunkStr), CRYPT_STRING_BASE64, lThunkPtr, lThunkSize, 0, 0)
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
    
    If GdipNewPrivateFontCollection(hNewFontCol) <> 0 Then
        GoTo QH
    End If
    If GdipPrivateAddFontFile(hNewFontCol, StrPtr(sFileName)) <> 0 Then
        GoTo QH
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
    If hFontCollection <> 0 Then
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
    If hNewFontCol <> 0 Then
        Call GdipDeletePrivateFontCollection(hNewFontCol)
        hNewFontCol = 0
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

Public Function GdipPrepareStringFormat(ByVal lFlags As UcsTouchButtonTextFlagsEnum, hStringFormat As Long) As Boolean
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

#If Not ImplUseShared Then

'= push-param thunk ======================================================

Public Sub InitPushParamThunk(Thunk As PushParamThunk, ByVal ParamValue As Long, ByVal pfnDest As Long)
'push [esp]
'mov eax, 16h // Dummy value for parameter value
'mov [esp + 4], eax
'nop // Adjustment so the next long is nicely aligned
'nop
'nop
'mov eax, 1234h // Dummy value for function
'jmp eax
'nop
'nop
    Dim dwDummy         As Long

    With Thunk.Code
        .Thunk(0) = &HB82434FF
        .Thunk(1) = ParamValue
        .Thunk(2) = &H4244489
        .Thunk(3) = &HB8909090
        .Thunk(4) = pfnDest
        .Thunk(5) = &H9090E0FF
        Call VirtualProtect(.Thunk(0), Len(Thunk), PAGE_EXECUTE_READWRITE, dwDummy)
    End With
    Thunk.pfn = VarPtr(Thunk.Code)
End Sub

'= fire-once timers ======================================================

Public Sub InitFireOnceTimer(Data As FireOnceTimerData, ByVal ThisPtr As Long, ByVal pfnRedirect As Long, Optional ByVal Delay As Long)
    With Data
        InitPushParamThunk .TimerProcThunkData, VarPtr(Data), pfnRedirect
        InitPushParamThunk .TimerProcThunkThis, ThisPtr, .TimerProcThunkData.pfn
        .TimerID = SetTimer(0, 0, Delay, .TimerProcThunkThis.pfn)
    End With
End Sub

Public Sub TerminateFireOnceTimer(Data As FireOnceTimerData)
    With Data
        If .TimerID <> 0 Then
            Call KillTimer(0, .TimerID)
            .TimerID = 0
        End If
    End With
End Sub

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

#End If ' ImplUseShared

'==============================================================================
' Redirectors
'==============================================================================

Public Sub RedirectTouchButtonTimerProc( _
            Data As FireOnceTimerData, _
            ByVal This As ctxTouchButton, _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal idEvent As Long, _
            ByVal dwTime As Long)
    #If hWnd And wMsg And dwTime Then '--- touch
    #End If
    Data.TimerID = idEvent
    TerminateFireOnceTimer Data
    This.frTimer
End Sub

