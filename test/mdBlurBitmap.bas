Attribute VB_Name = "mdBlurBitmap"
Option Explicit
DefObj A-Z

'--- for VirtualProtect
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const MEM_COMMIT                    As Long = &H1000
'--- for CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1
'--- for gdi+
Private Const ImageLockModeRead             As Long = &H1
Private Const ImageLockModeWrite            As Long = &H2
Private Const PixelFormat32bppARGB          As Long = &H26200A

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
'--- gdi+
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, lpRect As Any, ByVal lFlags As Long, ByVal lPixelFormat As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, uLockedBitmapData As BitmapData) As Long

Private Type BitmapData
    Width               As Long
    Height              As Long
    Stride              As Long
    PixelFormat         As Long
    Scan0               As Long
    Reserved            As Long
End Type

Public Function BlurBitmap( _
            ByVal hBitmap As Long, _
            ByVal sngRadius As Single, _
            Optional ByVal AffectChannels As Long = 15) As Boolean
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
    BlurBitmap = True
QH:
    On Error Resume Next
    If uData.Scan0 <> 0 Then
        Call GdipBitmapUnlockBits(hBitmap, uData)
    End If
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description & " [BlurBitmap]"
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
    pvPatchThunk AddressOf mdBlurBitmap.LoadSave, STR_THUNK
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
    pvPatchThunk AddressOf mdBlurBitmap.ProcessRow, STR_THUNK
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
