Attribute VB_Name = "mdRoundedRectangle"
Option Explicit
DefObj A-Z

Private Const Transparent                   As Long = &HFFFFFF
Private Const FillModeAlternate             As Long = 0
Private Const DashStyleSolid                As Long = 0
Private Const UnitPixel                     As Long = 2
Private Const SmoothingModeAntiAlias        As Long = 4

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, pInputBuf As Any, Optional ByVal pOutputBuf As Long = 0) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal lSmoothingMd As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal lColor As Long, ByVal sngWidth As Single, ByVal lUnit As Long, hPen As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal hPen As Long, ByVal dStyle As Long) As Long
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal lBrushmode As Long, hPath As Long) As Long
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal hPath As Long, ByVal sngX As Single, ByVal sngY As Single, ByVal sngWidth As Single, ByVal sngHeight As Single, ByVal sngStartAngle As Single, ByVal sngSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal hPath As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal lArgb As Long, hBrush As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal hPath As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal pen As Long, ByVal hPath As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal hPen As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal hPath As Long) As Long
Private Declare Function GdipBeginContainer2 Lib "gdiplus" (ByVal hGraphics As Long, hState As Long) As Long
Private Declare Function GdipEndContainer Lib "gdiplus" (ByVal hGraphics As Long, ByVal hState As Long) As Long

Public Function DrawRoundedRectangle( _
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
    DrawRoundedRectangle = True
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
    Debug.Print "Critical error: " & Err.Description & " [DrawRoundedRectangle]"
    Resume QH
End Function

Public Sub StartGdip()
    Dim aInput(0 To 3)  As Long
    
    If GetModuleHandle("gdiplus") = 0 Then
        aInput(0) = 1
        Call GdiplusStartup(0, aInput(0))
    End If
End Sub

