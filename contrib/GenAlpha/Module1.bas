Attribute VB_Name = "Module1"
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const STD_INPUT_HANDLE              As Long = -10&
Private Const STD_OUTPUT_HANDLE             As Long = -11&
Private Const STD_ERROR_HANDLE              As Long = -12&
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, lpszDst As Any, ByVal cchDstLength As Long) As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
Private Declare Function ApiCreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function ApiDeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oOpt          As Object

Private Type RGBQUAD
    B                   As Long
    G                   As Long
    R                   As Long
    A                   As Long
End Type

'=========================================================================
' Error management
'=========================================================================

Private Function PrintError(sFunction As String) As VbMsgBoxResult
    Debug.Print sFunction & ": " & Err.Description, Timer
End Function

'=========================================================================
' Functions
'=========================================================================

Sub Main()
    Dim oWhite          As cDibSection
    Dim oBlack          As cDibSection
    Dim oOutput         As cDibSection
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim uWhite          As RGBQUAD
    Dim uBlack          As RGBQUAD
    Dim lAlpha          As Long
    Dim lRetVal         As Long
    
    On Error GoTo EH
    Set m_oOpt = GetOpt(SplitArgs(Command$), "w:b:o")
    If Not m_oOpt.Exists("-w") Or Not m_oOpt.Exists("-b") Or Not m_oOpt.Exists("-o") Or m_oOpt.Exists("-h") Then
        ConsoleError "Usage: GenAlpha.exe -w <white_background> -b <black_background> -o <output_png>" & vbCrLf
        Exit Sub
    End If
    Set oWhite = New cDibSection
    If Not oWhite.LoadFromFile(m_oOpt.Item("-w")) Then
        ConsoleError "Error: Loading white image from '%1'" & vbCrLf, m_oOpt.Item("-w")
        lRetVal = 1
        GoTo QH
    End If
    Set oBlack = New cDibSection
    If Not oBlack.LoadFromFile(m_oOpt.Item("-b")) Then
        ConsoleError "Error: Loading black image from '%1'" & vbCrLf, m_oOpt.Item("-b")
        lRetVal = 1
        GoTo QH
    End If
    lWidth = LimitLong(oWhite.Width, 0, oBlack.Width)
    lHeight = LimitLong(oWhite.Height, 0, oBlack.Height)
    Set oOutput = New cDibSection
    If Not oOutput.Init(lWidth, lHeight) Then
        ConsoleError "Error: Creating output image w/ size %1px x %2px" & vbCrLf, lWidth, lHeight
        lRetVal = 1
        GoTo QH
    End If
    For lJdx = 0 To lHeight - 1
        For lIdx = 0 To lWidth - 1
            oWhite.GetPixel lIdx, lJdx, uWhite.R, uWhite.G, uWhite.B, uWhite.A
            oBlack.GetPixel lIdx, lJdx, uBlack.R, uBlack.G, uBlack.B, uBlack.A
            lAlpha = 255 - LimitLong(LimitLong(LimitLong(uWhite.R - uBlack.R, 0, 255), LimitLong(uWhite.G - uBlack.G, 0, 255)), LimitLong(uWhite.B - uBlack.B, 0, 255))
            If lAlpha = 0 Then
                oOutput.SetPixel lIdx, lJdx, 0, 0, 0, 0
            Else
                oOutput.SetPixel lIdx, lJdx, LimitLong(uBlack.R * 255 / lAlpha, 0, 255), LimitLong(uBlack.G * 255 / lAlpha, 0, 255), LimitLong(uBlack.B * 255 / lAlpha, 0, 255), lAlpha
            End If
        Next
    Next
    WriteBinaryFile m_oOpt.Item("-o"), oOutput.SaveToByteArray("image/png")
    If Not FileExists(m_oOpt.Item("-o")) Then
        ConsoleError "Error: Writing output image '%1'" & vbCrLf, m_oOpt.Item("-o")
        lRetVal = 1
        GoTo QH
    End If
QH:
    If Not InIde Then
        Call ExitProcess(lRetVal)
    End If
    Exit Sub
EH:
    ConsoleError "Critical: %1" & vbCrLf, Err.Description
    lRetVal = -1
    Resume QH
End Sub

Public Function GetOpt(vArgs As Variant, Optional OptionsWithArg As String) As Object
    Dim oRetVal         As Object
    Dim lIdx            As Long
    Dim bNoMoreOpt      As Boolean
    Dim vOptArg         As Variant
    Dim vElem           As Variant

    vOptArg = Split(OptionsWithArg, ":")
    Set oRetVal = CreateObject("Scripting.Dictionary")
    With oRetVal
        .CompareMode = vbTextCompare
        For lIdx = 0 To UBound(vArgs)
            Select Case Left$(At(vArgs, lIdx), 1 + bNoMoreOpt)
            Case "-", "/"
                For Each vElem In vOptArg
                    If Mid$(At(vArgs, lIdx), 2, Len(vElem)) = vElem Then
                        If Mid(At(vArgs, lIdx), Len(vElem) + 2, 1) = ":" Then
                            .Item("-" & vElem) = Mid$(At(vArgs, lIdx), Len(vElem) + 3)
                        ElseIf Len(At(vArgs, lIdx)) > Len(vElem) + 1 Then
                            .Item("-" & vElem) = Mid$(At(vArgs, lIdx), Len(vElem) + 2)
                        ElseIf LenB(At(vArgs, lIdx + 1)) <> 0 Then
                            .Item("-" & vElem) = At(vArgs, lIdx + 1)
                            lIdx = lIdx + 1
                        Else
                            .Item("error") = "Option -" & vElem & " requires an argument"
                        End If
                        GoTo Conitnue
                    End If
                Next
                .Item("-" & Mid$(At(vArgs, lIdx), 2)) = True
            Case Else
                .Item("numarg") = .Item("numarg") + 1
                .Item("arg" & .Item("numarg")) = At(vArgs, lIdx)
            End Select
Conitnue:
        Next
    End With
    Set GetOpt = oRetVal
End Function

Public Function ConsolePrint(ByVal sText As String, ParamArray A() As Variant) As String
    ConsolePrint = pvConsoleOutput(GetStdHandle(STD_OUTPUT_HANDLE), sText, CVar(A))
End Function

Public Function ConsoleError(ByVal sText As String, ParamArray A() As Variant) As String
    ConsoleError = pvConsoleOutput(GetStdHandle(STD_ERROR_HANDLE), sText, CVar(A))
End Function

Private Function pvConsoleOutput(ByVal hOut As Long, ByVal sText As String, A As Variant) As String
    Dim lIdx            As Long
    Dim sArg            As String
    Dim baBuffer()      As Byte
    Dim dwDummy         As Long

    '--- format
    For lIdx = UBound(A) To LBound(A) Step -1
        sArg = Replace(A(lIdx), "%", ChrW$(&H101))
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), sArg)
    Next
    pvConsoleOutput = Replace(sText, ChrW$(&H101), "%")
    '--- output
    If hOut = 0 Then
        Debug.Print pvConsoleOutput;
    Else
        ReDim baBuffer(0 To Len(pvConsoleOutput) - 1) As Byte
        If CharToOemBuff(pvConsoleOutput, baBuffer(0), UBound(baBuffer) + 1) Then
            Call WriteFile(hOut, baBuffer(0), UBound(baBuffer) + 1, dwDummy, ByVal 0&)
        End If
    End If
End Function

Public Function SplitArgs(sText As String) As Variant
    Dim vRetVal         As Variant
    Dim lPtr            As Long
    Dim lArgc           As Long
    Dim lIdx            As Long
    Dim lArgPtr         As Long

    If LenB(sText) <> 0 Then
        lPtr = CommandLineToArgvW(StrPtr(sText), lArgc)
    End If
    If lArgc > 0 Then
        ReDim vRetVal(0 To lArgc - 1) As String
        For lIdx = 0 To UBound(vRetVal)
            Call CopyMemory(lArgPtr, ByVal lPtr + 4 * lIdx, 4)
            vRetVal(lIdx) = SysAllocString(lArgPtr)
        Next
    Else
        vRetVal = Split(vbNullString)
    End If
    Call LocalFree(lPtr)
    SplitArgs = vRetVal
End Function

Private Function SysAllocString(ByVal lPtr As Long) As String
    Dim lTemp           As Long

    lTemp = ApiSysAllocString(lPtr)
    Call CopyMemory(ByVal VarPtr(SysAllocString), lTemp, 4)
End Function

Public Function At(vArray As Variant, ByVal lIdx As Long) As Variant
    On Error GoTo QH
    At = vArray(lIdx)
QH:
End Function

Public Function FileAttr(sFile As String) As VbFileAttribute
    FileAttr = GetFileAttributes(sFile)
    If FileAttr = -1 Then
        FileAttr = &H8000
    End If
End Function

Public Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Public Property Get InIde() As Boolean
    Debug.Assert pvSetTrue(InIde)
End Property

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

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

Public Sub WriteBinaryFile(sFile As String, baBuffer() As Byte)
    Dim nFile           As Integer
    
    On Error GoTo EH
    If InStrRev(sFile, "\") > 1 Then
        MkPath Left$(sFile, InStrRev(sFile, "\") - 1)
    End If
    DeleteFile sFile
    nFile = FreeFile
    Open sFile For Binary Access Write Shared As nFile
'    If Peek(ArrPtr(baBuffer)) <> 0 Then
        If UBound(baBuffer) >= 0 Then
            Put nFile, , baBuffer
        End If
'    End If
    Close nFile
    Exit Sub
EH:
    Close nFile
End Sub

Public Function MkPath(sPath As String, Optional sError As String) As Boolean
    Const FUNC_NAME     As String = "MkPath"
    
    On Error GoTo EH
    MkPath = (FileAttr(sPath) And vbDirectory) <> 0
    If Not MkPath Then
        If ApiCreateDirectory(sPath, 0) = 0 Then
            sError = GetSystemMessage(Err.LastDllError)
        End If
        MkPath = (FileAttr(sPath) And vbDirectory) <> 0
        If Not MkPath And InStrRev(sPath, "\") <> 0 Then
            MkPath Left$(sPath, InStrRev(sPath, "\") - 1)
            Call ApiCreateDirectory(sPath, 0)
            MkPath = (FileAttr(sPath) And vbDirectory) <> 0
        End If
    End If
    Exit Function
EH:
    If PrintError(FUNC_NAME & "(sPath=" & sPath & ")") = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Public Function FileExists(sFile As String) As Boolean
    If GetFileAttributes(sFile) = -1 Then ' INVALID_FILE_ATTRIBUTES
    Else
        FileExists = True
    End If
End Function

Public Function DeleteFile(sFileName As String) As Boolean
    Const FUNC_NAME     As String = "DeleteFile"
    
    On Error GoTo EH
    If LenB(sFileName) <> 0 Then
        If ApiDeleteFile(sFileName) = 0 Then
            If FileExists(sFileName) Then
                Call SetFileAttributes(sFileName, vbArchive)
                Call ApiDeleteFile(sFileName)
                DeleteFile = Not FileExists(sFileName)
                Exit Function
            End If
        End If
    End If
    DeleteFile = True
    Exit Function
EH:
    If PrintError(FUNC_NAME & "(sFileName=" & sFileName & ")") = vbRetry Then
        Resume
    End If
    Resume Next
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
