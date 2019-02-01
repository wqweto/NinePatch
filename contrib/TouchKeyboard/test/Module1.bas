Attribute VB_Name = "Module1"
Option Explicit

'--- for GetLocaleInfo
Private Const LOCALE_SISO639LANGNAME        As Long = &H59
'--- for ActivateKeyboardLayout
Private Const HKL_NEXT                      As Long = 1
Private Const KLF_ACTIVATE                  As Long = &H1
Private Const KLF_SETFORPROCESS             As Long = &H100
'--- for LoadKeyboardLayout
Private Const KLID_BULGARIAN                As String = "00030402"
Private Const KLID_US                       As String = "00000409"

Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hKL As Long, ByVal Flags As Long) As Long
Private Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Property Get KeybLanguage() As String
    KeybLanguage = pvGetUserLocaleInfo(GetKeyboardLayout(0) And &HFFFF&, LOCALE_SISO639LANGNAME)
End Property

Property Let KeybLanguage(sValue As String)
    Dim hKL             As Long
    Dim hActive         As Long
    Dim sKLID           As String
    
    If LenB(sValue) = 0 Then
        GoTo QH
    End If
    hActive = GetKeyboardLayout(0)
    hKL = hActive
    Do
        If LCase$(pvGetUserLocaleInfo(hKL And &HFFFF&, LOCALE_SISO639LANGNAME)) = LCase$(sValue) Then
            GoTo QH
        End If
        Call ActivateKeyboardLayout(HKL_NEXT, 0)
        hKL = GetKeyboardLayout(0)
    Loop While hKL <> hActive
    Select Case LCase$(sValue)
    Case "bg"
        sKLID = KLID_BULGARIAN
    Case "en"
        sKLID = KLID_US
    End Select
    If LoadKeyboardLayout(sKLID, KLF_ACTIVATE Or KLF_SETFORPROCESS) = 0 Then
        Debug.Print "LoadKeyboardLayout, sKLID=" & sKLID & ", Err.LastDllError=" & Err.LastDllError, vbCritical
    End If
    If ActivateKeyboardLayout(sKLID, KLF_SETFORPROCESS) = 0 Then
        If Err.LastDllError <> 0 Then
            Debug.Print "ActivateKeyboardLayout, sKLID=" & sKLID & ", Err.LastDllError=" & Err.LastDllError, vbCritical
        End If
    End If
QH:
End Property

Private Function pvGetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String
   Dim sReturn          As String
   Dim nSize            As Long

   nSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
   If nSize > 0 Then
      sReturn = Space$(nSize)
      nSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
      If nSize > 0 Then
         pvGetUserLocaleInfo = Left$(sReturn, nSize - 1)
      End If
   End If
End Function
