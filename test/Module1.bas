Attribute VB_Name = "Module1"
Option Explicit

'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
'--- for CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, pcchString As Long) As Long

Public Function GdipGetSystemMessage(ByVal lStatus As Long, ByVal lLastDllError As Long) As String
    Const STR_STATUS    As String = "|Generic Error|Invalid Parameter|Out Of Memory|Object Busy|Insufficient Buffer|Not Implemented|Win32 Error|Wrong State|AbortedFile Not Found|Value Overflow|Access Denied|Unknown Image Format|Font Family Not Found|Font Style Not Found|Not True Type Font|Unsupported Gdiplus Version|Gdiplus Not Initialized|Property Not Found|Property Not Supported|Profile Not Found"
    
    If lStatus <> 0 Then
        GdipGetSystemMessage = At(Split(STR_STATUS, "|"), lStatus, "Unknown error: " & lStatus)
        If lStatus = 7 And lLastDllError <> 0 Then ' Win32Error
            GdipGetSystemMessage = GdipGetSystemMessage & ". " & GetSystemMessage(lLastDllError)
        End If
    End If
End Function

Public Function GetSystemMessage(ByVal lLastDllError As Long) As String
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


Public Function At(Data As Variant, ByVal Index As Long, Optional Default As String) As String
    On Error GoTo RH
    At = Default
    If LBound(Data) <= Index And Index <= UBound(Data) Then
        At = CStr(Data(Index))
    End If
RH:
End Function

Public Function ReadBinaryFile(sFile As String) As Byte()
    Dim baBuffer()      As Byte
    Dim nFile           As Integer

    On Error GoTo EH
    baBuffer = vbNullString
    If GetAttr(sFile) Or True Then
        nFile = FreeFile
        Open sFile For Binary Access Read Shared As nFile
        If LOF(nFile) > 0 Then
            ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
            Get nFile, , baBuffer
        End If
        Close nFile
    End If
    ReadBinaryFile = baBuffer
EH:
End Function

Public Function ToBase64Array(baData() As Byte) As String
    Dim lSize           As Long
    
'    If Peek(ArrPtr(baData)) <> 0 Then
        If UBound(baData) >= 0 Then
            Call CryptBinaryToString(VarPtr(baData(0)), UBound(baData) + 1, CRYPT_STRING_BASE64, 0, lSize)
            If lSize > 0 Then
                ToBase64Array = String$(lSize - 1, 0)
                Call CryptBinaryToString(VarPtr(baData(0)), UBound(baData) + 1, CRYPT_STRING_BASE64, StrPtr(ToBase64Array), lSize)
            End If
        End If
'    End If
End Function
