VERSION 5.00
Begin VB.UserControl ctxNineButton 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4044
   ClipBehavior    =   0  'None
   ScaleHeight     =   105
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   Windowless      =   -1  'True
End
Attribute VB_Name = "ctxNineButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' Nine Patch PNGs for VB6 (c) 2018 by wqweto@gmail.com
'
' ctxNineButton.ctl -- windowless 9-patch button control w/ state animation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "ctxNineButton"

#Const ImplUseShared = NPPNG_USE_SHARED <> 0
#Const ImplHasTimers = True

'=========================================================================
' Public events
'=========================================================================

Event Click()
Event DblClick()
Event ContextMenu()
Event OwnerDraw(ByVal hGraphics As Long, ByVal hFont As Long, ByVal ButtonState As UcsNineButtonStateEnum, ClientLeft As Long, ClientTop As Long, ClientWidth As Long, ClientHeight As Long)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsNineButtonStateEnum
    ucsBstNormal = 0
    ucsBstHover = 1
    ucsBstPressed = 2
    ucsBstHoverPressed = 3
    ucsBstDisabled = 4
    ucsBstFocused = 8
End Enum
Private Const ucsBstLast = ucsBstFocused

Public Enum UcsNineButtonTextFlagsEnum
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

Public Enum UcsNineButtonStyleEnum
    ucsBtyNone
    ucsBtyButtonDefault
    ucsBtyButtonGreen
    ucsBtyButtonRed
    ucsBtyButtonTurnGreen
    ucsBtyButtonTurnRed
    ucsBtyFlatPrimary
    ucsBtyFlatSecondary
    ucsBtyFlatSuccess
    ucsBtyFlatDanger
    ucsBtyFlatWarning
    ucsBtyFlatInfo
    ucsBtyFlatLight
    ucsBtyFlatDark
    ucsBtyOutlinePrimary
    ucsBtyOutlineSecondary
    ucsBtyOutlineSuccess
    ucsBtyOutlineDanger
    ucsBtyOutlineWarning
    ucsBtyOutlineInfo
    ucsBtyOutlineLight
    ucsBtyOutlineDark
    ucsBtyCardDefault
    ucsBtyCardPrimary
    ucsBtyCardSuccess
    ucsBtyCardOrange
    ucsBtyCardDanger
    ucsBtyCardWarning
    ucsBtyCardPurple
    ucsBtyCardFocus
End Enum

'=========================================================================
' API
'=========================================================================

'--- for GdipCreateBitmapFromScan0
Private Const PixelFormat32bppARGB          As Long = &H26200A
'--- for GdipDrawImageXxx
Private Const UnitPixel                     As Long = 2
Private Const UnitPoint                     As Long = 3
'--- for RedrawWindow
Private Const RDW_INVALIDATE                As Long = &H1
Private Const RDW_ERASE                     As Long = &H4
Private Const RDW_ALLCHILDREN               As Long = &H80
Private Const RDW_UPDATENOW                 As Long = &H100
Private Const RDW_FRAME                     As Long = &H400
'--- for CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1
'--- for GdipSetTextRenderingHint
Private Const TextRenderingHintAntiAlias    As Long = 4
Private Const TextRenderingHintClearTypeGridFit As Long = 5

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, uInputBuf As GdiplusStartupInput, Optional ByVal lOutputBuf As Long = 0) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal lWidth As Long, ByVal lHeight As Long, ByVal lStride As Long, ByVal lPixelFormat As Long, Scan0 As Any, hBitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal hImage As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal srcUnit As Long = UnitPixel, Optional ByVal hImageAttributes As Long, Optional ByVal pfnCallback As Long, Optional ByVal lCallbackData As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (hImgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hImgAttr As Long, ByVal clrAdjust As Long, ByVal clrAdjustEnabled As Long, clrMatrix As Any, grayMatrix As Any, ByVal clrMatrixFlags As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImgAttr As Long) As Long
Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal hBitmap As Long, ByVal lX As Long, ByVal lY As Long, clrCurrent As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal lNamePtr As Long, ByVal hFontCollection As Long, hFontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "gdiplus" (hFontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal hFontFamily As Long) As Long
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal hFontFamily As Long, ByVal emSize As Single, ByVal lStyle As Long, ByVal lUnit As Long, hFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal hFont As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, hBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal hGraphics As Long, ByVal lStrPtr As Long, ByVal lLength As Long, ByVal hFont As Long, uRect As RECTF, ByVal hStringFormat As Long, ByVal hBrush As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal hFormatAttributes As Long, ByVal nLanguage As Integer, hStringFormat As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal hStringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal hStringFormat As Long, ByVal lFlags As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal hStringFormat As Long, ByVal eAlign As StringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal hStringFormat As Long, ByVal eAlign As StringAlignment) As Long
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal hGraphics As Long, ByVal lMode As Long) As Long
Private Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal lPixelFormat As Long, ByVal srcBitmap As Long, dstBitmap As Long) As Long
#If Not ImplUseShared Then
    Private Declare Function GetSystemTimeAsFileTime Lib "kernel32" (lpSystemTimeAsFileTime As Currency) As Long
    Private Declare Function ApiRedrawWindow Lib "user32" Alias "RedrawWindow" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
    Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, pdwSkip As Long, pdwFlags As Long) As Long
#End If

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type RECTF
   Left             As Single
   Top              As Single
   Right            As Single
   Bottom           As Single
End Type

Private Enum FontStyle
   FontStyleRegular = 0
   FontStyleBold = 1
   FontStyleItalic = 2
   FontStyleBoldItalic = 3
   FontStyleUnderline = 4
   FontStyleStrikeout = 8
End Enum

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

'=========================================================================
' Constants and variables
'=========================================================================

Private Const DBL_EPLISON           As Double = 0.000001
Private Const DEF_STYLE             As Long = ucsBtyButtonDefault
Private Const DEF_ENABLED           As Boolean = True
Private Const DEF_OPACITY           As Double = 1
Private Const DEF_ANIMATIONDURATION As Double = 0
Private Const DEF_FORECOLOR         As Long = vbButtonText
Private Const DEF_TEXTOPACITY       As Single = 1
Private Const DEF_TEXTCOLOR         As Long = -1  '--- none
Private Const DEF_TEXTFLAGS         As Long = ucsBflCenter
Private Const DEF_IMAGEOPACITY      As Single = 1
Private Const DEF_SHADOWOPACITY     As Single = 0.5
Private Const DEF_SHADOWCOLOR       As Long = vbButtonShadow
Private Const STR_RES_PNG1          As String = "iVBORw0KGgoAAAANSUhEUgAAAOcAAACfCAYAAAAChc6MAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAA3iElEQVR4Xu2dCXwUVbr2HUa+uYv+7vUOKpqNJQmQAAJJIBtJQBRxQWRxx2V0vAqK2yzK9bsGRMAFGMEQUBCFYMISEFwAHQkkCEICEUdIgqaTKCiy6ziRhPk43/tUdyUnp9/KqUo6QEzV7/dPU+d93qfb7jye6q7TlfOEELagjW74WlNxPflaU3E9+VpTOdue3h8B3NQ7cBHnPTC/qP3tU3aFjZm8K2H0c4VpKhhHHTqu36V1Mn/+/PYvv7og7KVZ8xKmzcpMU8E46tBx/d4fAdzUO2jLpAvR7ubJxUk3T945cfSkwnQd0N06pTgRfZyfS6uh3YuzMpOmz5w7" & _
                                                "cdqMjHQd0E1/JSMRfbKP90cAN9m8LTM2ffclo54rHM+FUAf6bp3ytxDO1+XcZkZGRsj0WXPHcyHUgT70m17eH8p2Y/quqJGTizJGTSoqHjW5SPhB46jflF7Y09dSt5nGKqMmfxbt7SkqGDmpUKiQV4HpyfVzJP1+ZXTyg+9kpIxbUzzwoXeESvJD7xSjnvjgKtueA+7Nioq/b+nchPveLiYEQzHqsfctbtTz5ud2DTHDdvOUXaPGTPksyOqwFeOoj3mucHRdz6SiQZwW0Parym8O3l75zeHCqv2Hq6sOHBF+0DjqFV8fvg16zkdl6svzoqbOeHXu8zMy8ulW+PHyqwV0O/f5l16x9Xzifr8oKb+9tMyzo3Sfp5oQDNWol5SUO36cU2dmFE+bmSFUMO70cW7btuOO7dsLC7dvL6resWOnUME46lu3br+9scf50sy5Q8ywvfiXzFEzZ2YGWR22Yhz1F2ZljDZ70G/WvT+kzQjmpKKKUZML" & _
                                                "a9lg1lFYO3JyoUcNqGksA89R6TsKbppU+PDIqYUdhqV/2l4F46gjpHYCmjR+bfTAcWs8KQ9/UJs64UOR9uhHfmAc9YHj3vHYCWjc79+Ojv9ddkXSAytqkx9c7Rd2I/A0jnr8fdmexPtyWM/0PHH+6MlFzxhBe27n9ZzGCujRN+a5nf9D++wvQcXX3z+0/+AxUf1zjTh9+jQN+W8YRx066GnIz0cGv/DPvzxnCzH+z9OndyDaM3RAferLcwrs/OJ/sffLcV+VV4mjR4+Jf/zjH6K6utoPjKMOHfScj4wRzJkZFRS+U1wwTVAnPHYe57ZtRQ/t2vmZ8JR7xIH9+8W3Bw74gXHUoYOe8yF+RYeozxjBnJHp6HWHHn0vzMqse92NgryNmrxzyajndp7mA6lg6HYu8bUam3lnMpgRR6YXjudCqQId9JyPzMAH12SlPryODaUKdNBzPjLx97+dlfTfK9lQqkAHPedzxwu7g0f7ZsAx6Ts6chor" & _
                                                "oK/rffELtrfim0O7T9bU0j/1G3TQ0z/9fGSMmcYbTC6UDTACSnrOR6aktHz38eMn2FCqQAc95yND95vFhdEK6DkfmU8/LdpdWVHBhlIFOug5nxczFnU0Z8AXMzIcve7Q1/cuMnqNgrzRrFnCBtGKSTv3+FqNTb5DE8yGFLzfqkG04LcjSM/5yCQ/tKY0dQIfRhXoSL+H85GhQ9ZSOhRmw6gCXfzv3mY9b5myM360L2D3pOf9C6exAnqzFz6cpnL/kRNWM6a6QQc9/dPPRwaHrBQ8qxlTpcPzM17N53xkSveVn7CaMVWgg57zkaGwlXAhtGLqjLna133HjqITVjOmCnTQcz4vzJ4bbwYsfdEiR6879GYvfDBmFOSNDaAGX6uxyXdogveUTAgtgV71UEFAuCBaAb3qoYL3lHIAdUCvegDjFIkvYFxdh9kLH66O95RONp/ez0eGfukFhY4LIgv0qocK3lNyQbQCetVDhQugDtVDBe8puSBa" & _
                                                "Ab3qAYxTJL6AcXUdZi98sG8MyhsXPh2+VmOT78zEDaczzF43nP5w4dOheqics+FEnoyfvo0Lnw5fq7lhn7zr77S54eQ8EQ4uhFZAL/dzns0Np+kZyHByj7O54eQ8mxtOzrO54eQ8ufDpkPs5z+aG0/QMZDgNTwzKGxc+Hb5WY1PvELgzpzPMXnfm9IcLnw7VQ+WcnTmNH9LGhU+Hr9XY5DszccPpDLPXDac/XPh0qB4qrSacIycVlXEBtAJ6X6uxyXdm4vTTWug5H5nkcWvKHH1aS3rORybx/uwyJ5/WQs/5nIlwOvm01lY4HX5aCz3nI4OwOfu01kY4Z2SUcQG0hPScjwzC5uzT2rMXziwuhFZA72s1NvnOTNraec4x03aGj/YFrFnnOcmH03i+/v5zJ+c5oad/+vnI0EwY8POce0u++tzJeU7oOR8Zut+An+fctq3wcyfnOaHnfGbOXBBuBqw55znhgzGjIG9YnUPhqOSC6Mekoiro" & _
                                                "fa3GJt+hCTSYDe2sEBoxqSgfes5HBit+aPaqSnn4A2MlEBtK7wohCtOaKjsrhLDiJ+H+7KqkB1YYK4G4UPpWCAnorFYIjcn44oLRvoA1dYUQGDNz679ympJ9VY84WSEEPQ35+chgJY2x8scb0ICsECra9bcJTlYIQc/5yBiPc2ZGJRdEf16tsvM4N24smOBkhRD0nE9GxvIL6sLZ9BVCz760ePG/Y8woqBvCMWpyYS7NihVcKCm8HtTVYGKT71AGWsyICCneU6oYa269de2TaeIN6JrcgePWerggYRx1J2trsWY2/r7s3IT7cjx4T+nH/TnlqOvX1haNGG2G7LnC0XbW1mINrtmDfk4LaPtN8ef77imv+n5nY2trUf/8i333QM/5qBi/+DQj+tbQetfTSvjW3DpZs/qbbdt23rNn75c7S8rKf8ZhqwrGUd+2o9jZ45w5J5cei4cL5dQZGR7UnTzODz/8+J5t23bsbGxtLeofffTxvY09" & _
                                                "zhdmzR1hBhRrZu2srcUa3PqeuXWvu/dHADfTuK2DwN383M4mfyvFKsgu5zYI3Asz540zw+YEfCtFDrL3RwA309jF+33OMZN3pdYtgteA73Pi+5/o4/xcWg3tXpw5N9VcBK8D3+fE9z/RJ/t4fwRwk81dvGAWdK+E0PbALNjsKyHYwZs7vtZUXE++1lRcT77WVM62JzvIcbYfqF1cT77WVFxPvtZUnHiygy6B5YGimPZPftwj7LF13ROINIYE1KHj+l1anqKYmPafdO4Rtimse0Jep+5pKhhHHTquvyVgB10CQ7pIb/fYh92SHl/ffeJj6yLTdUD3xIbuiejj/FxagPT0dvmh3ZIofBPzQiPTdUC3Oax7IvpYvwDCDro0nyc29Ax5bH3keC6EWqgP/ZyvS+DY3LVnSF5Y5HguhFqoD/2cb6BgBx/5IDzq0XWRGY9uiCh+dH2k8IfGqT7hva62T+6bno+sj8ifsC5CqFC9wKnnDZM7Rt84" & _
                                                "pWPGiOcvKx7xfEfhz6XFqEPH9XMM/fN/RV/z1MUZwyZeXDzsqYuFH09fXIz61X+4qNHH+dj6HkPMsD2xvvuoxz+ODLI6bMU46o9t6Da6PqA96i70xLExNDx6U0hkxubQiOJNoZHCHxqnel6I/eezzjM4oiAvJEL4ERxe4NRz3n91jMq4JChj7iXBxXMvDRF+0Djqr/6X/ddoOmlf6hA09+UOwcUvXxws/KBx1F+66NJGH2deaI8hZtjyw7qP+jgyMsjqsBXjqOeFdhtdH9LGX6PvNl4RfWhzv4zDm2OKCaFyaFO/z1A/mNebfZx+A94QRVRMWB9RywfTC+qEx06Y4Dnhg65bHl7fdfwDy7t0INozdEAdIbXjicCNmNqxYvTLl9fe8kqQuHV2sB8YRx06OwFFMCl8Fdf/7yW1N065lAl7R4Fx1K99+mKPVUDT89LOf3xdt2eMYG7o7mgZF/RGODd0s7zAF0JEYamgANb6h1KG6iERHjth" & _
                                                "gufG4K5b8oK7jv/ooi4diPYMHVDfSCG144lgZl4aUvH6pWG1b3YME29d1skPjKNOQfXYCSiCOePioIpXLg6pzbgk2D/sBMZRn3FxsMcqoHlpaefnhXV7xhsyZ68R9OjbHGb9GhnB3NSv4vCmfrVcMOug+qHNMR4uoA12AM1iWVwYrYBe9VDBjOgLJhfKBkAHPecjc+PUS7MoeGwoVaCDnvORGfZUh6wbnr2EDaUKdNBzPn/K6x1szoB/fD/K0QJo6Ot68/jevNCILD6MPNBzPjLGjOgNJhfKBkAHPecjk3lJcNaCjqFsKFWgg57zkaFZMesVi1CqQAc955MX3ju4bgbs5Ow1gl7Xe3hzvyw2jJb083ucDXYAzYYlXAitmLA+XHsBJcyGFDyrGVOlwyPrwrUXj6KAlFrNmCrQkX4v5yNDYSu1mjFVoBv2dAf2v/2Jj3rEmwGjWdTZhZ5Ib/bCh9PQjFjChdCKzSH61wiHrBQ8qxlThWbQ" & _
                                                "cO1XxjIvDS6xmjFVoJt7SYj2cc7oEFRiNWOqQEfhZD3zOvWIrw+Ys9cI+vpe/jWiWbOUD6EV/fweZ4MdwAVQh+qhgveUTAgtgV71UEFAuCBaAb3qoYL3lHIAdUCvegDv6RFvwLi6DrMXPlydC6AO1UMF7ymZEFoCveqhgoBwQbQCetVDBe8p5QDqgF71AMYpEl/AuLqO+nDyrxEfwMZRPRrsAC58OlQPFTecznDDaU2bCqe6YoELnw65H36qZ3PDyXkiHFwIrYBe7uc8mxtO0zOQ4eQeJxc+HXI/59nccHKeCAcXQiugl/s5z+aG0/QMZDi5x8mFT4fcb3jKA4ALnw7VQ8WdOZ0hh5Orc+HToXqouDOnM+RwcnUufDpUjwY7gAufDtVDxQ2nM9xwWtOmwzlhfcQ+LoBWQK96qDj+tJb0nI/MiKkdyxx9Wkt6zkdm2NMdyhx+Wst6tng4QyLLuABaQnrOR6YlPq2de0lwmbNPa4O1j3NG" & _
                                                "h6AyJ5/WQs/5nIFwlqnh0+D3OBvsAJq13POcTBhVGjvP+fj6qHAzYM05zwkfTuOe57QXzkbPc3aNCq8PWDPOc5IPp2mZ85zvde05YV1kJRdEFZo1q+ys5oEGs6EvoIFcIVSF4FnNoBhHHTo7K4Sw4odmw0oEz2oGxTjq1z7docp6hVDUBWbAmrxCaF3ks3/Y0Nu40JMKVudsComo4oLoB+nsrOaBxlj54w1oQFYIYcVP5qXBVQie1QyKcSOYpLOzQggrfmg2rELwrGZQjKMOneUKoaioC+oC1sQVQsSzu3vzrxFW/Bze1K+SD6JKvypbK4QAwvHouojcCesjPWogwYQNER6jbiNEJl7PyAwjfPSeUgVrblF34ukLaC7h4YI0EuNUtxNMEwTu2qcvzr124sUevKdUwTjq2rW1GyJHmAHFmlk7a2uxBre+J9LyAl/AF9BcmsE8XCg3h0R4ULcTIhOvJ2ZQhM9/bS3W3Bp1B54IHB2u5s67" & _
                                                "JMTDBckYp7qTtbUI3MsXB+VieR7eU6p4x4NytWtrO0WOqA9ot9F21tZiDW5dD/VzWhME7tCmfrnGMj4mlIc296tA3fbaWpfAMB+B2xA5ri5sTlgfOR79nK9L4EDgNoVFjqsPqAPCIsdbBTlQsIMugQHfy3z8w+6pj2/wLoLXYXzv88NuSe73Oc8g+D5np+6p9YvgGwff58T3P93vc/5CwGGreyWEcxvMgq32SgjqCohA4HrytabievK1pnK2PdlBjrP9QO3ievK1puJ68rWm4sSTHWzT0KHl8U8Sw45+0j/hcEFsmsrRTf0TUIeO7XdpcR4oEu3ve//nsLtW/Zhwz+rjaSp3rTqagDp0XH9rgR1sm6S3O5Yfm3QkP3bi4c2x6TqgO7o5LhF9vJ9LoEkXot3da35Muiv3xMS7co+n6zkx8d53fkxEH+d3rsMOtjUoZCFHNseO50KoA33o53xdAse971SH3L3q2Hg+hI2DPvRzvucy7GDd" & _
                                                "hYkK+uYf3twPKxgacGhT34LGLkzEsf618Oi/vhaZkfd6ZPHHr3cTKhtpHPUNmfZPcK+cGR79zivhGWvnhO9aMztC+DEnohj1Va807nmkIHaIGbaj+bGjjmxNCLI8bKVx1A/lx46uCyj1s1oft608Gn3XymNz71p1vJh+WYQfK2mc6mNzjtn+bzc978w9kT925TGhcufKYwVOPQe9+l30oHlHMq6cf7R48LwjQmUQjaOelnnQtmfSzKrolFf2ZwycfeCzlNn7hcrA2fs/Qz3xla8b9bzrnRND6Lnyhm31iVF3vXskyOqwFeOo08w52uxBP6c1qampiTp5sjajpvZUMf6mqUpNTW2xUa+psb1YwvT8uaam4OeTNUKl+ueagsY8/QYQzMP5fQuObOo3/uin4b8l2jP8FnWE1E5AEcyNC7pVFLzZvfbTrB5i+9IoPzCOOuk8dgKKYK6dE+FZnxlZS6H2CzvAOOrQWQc07fxDm+OeMYJW0N/R" & _
                                                "Mi7ofeG0vNATQjR25QkPBbDWCKIVVB+78rjHTpjgeceKY1uI8WOWH+tAtGfogDpCascTwaQAVgx5/Ujt0IVHxDVvHPUD46gPmn/YYyegRjBnH6hIzfi+dlDmYb+wA4yjTjqPVUDT88T5d+cef4aep/S7Vh1z9BpBj767V52wfI0QIgpLBYWwVg2lQi10dgIKz+qff95CwZxw7NixYOI/GIJRJ10B59lgB2BG9AWTC2UDjICSXvVQyVsQuWTLm93ZUKpABz3nI7NmdngWBY8NpQp00HM+Rz8dEGzOgIfy4hwtgIbe7P1pB987NvdYVoMQaoCe85HBjOgLJhfKBkAHPecjM3j+kSVXvc6HUgU66DkfGQpcVlrGQTaUKtBBz/nc/0F1MD03xgx4z/KfHL1G0Nf1vs/3UuaWMEG0BHrOR8Y3YyKYXCgbAB30qkeDHYDZkIJnNWOq/BZ61UOFDllLrGZMFeg2vhahvdDT2tkRJVYzpgp0a2fz" & _
                                                "F7k6umVAvBmwCocX44Le7IUPp6FDzFKEzi6k1/63Yzak4FnNmCod7lxxTHvBNApIidWMqQLd4HlHtY+TwlZiNWOqQEeHvqznPWt/iKfnxhuwPOHoNYK+rpd8OA0FrkQNYGNQOLX/7ThkpeBZzZgqwTRz+r1GDXYA3lMyIbQEetVDBQHhgmgF9KqHCt5TygHUAb3qAYxTJL6AcXUdZi98uDoC5xTVQwXvKZkQWgK96qGCgHBBtAJ61UMF7ynlAOqAXvUAxukRX8C4ug6zFz5cnQugDtVDBe8pmRBaAr3q0WAHuOF0hhtOa9xw8kHkYMOprlhobjjhp3oiHFwIrYBe7uc8mxtO0zOQ4eQeJ8LmFLmf82xuODlPhIMLoRXQy/2cZ3PDaXoGMpzc4+TCp0Pu5zybG07DUx4A7szpDDmcXB1hc4rqoeLOnM4we1vdzKkOuOF0hhtOa9xw8kHksBXOlvm0tluZo09rSc/5yKydHb7P4ae17IXI" & _
                                                "WjqcY3OPlyFwdoGe85Fx/Gkt6TkfmcHzj5Q5+rSW9JyPzMDZB8ocflrLerZ0OCkY+7gAWgE95yPj9NNa6FWPBjugrZ3n/GFrQrgZsOac54QPp3HPczb/POd9a34Ip+fGG7BmnOeED6dpNec5vdc9CewKIaz4oZBUIXhWMyjGUd+4oFuVnRVCWPFDs2EVgmc1g2Ic9bVzwqssVwh9kXaBGbCmrhAinhW7r2Yv9ITVOXRYWYXg6YDOzmoeaDAb+gIakBVCWPEzeN7hKgTPagbFuC+YVXZWCGHFz8DZ+6sQPKsZFOOo06xZZbVCaFyeuICeHyNgTV0hRDz7hw2CfY2wOufkyZpKLogqNGtW2VwhFI2VP76ABmaFEDACSjMiwof3lH4U9M1H3cnaWgRu44LIXCzj44JEs6UHdSdra42Azuma++6cCA/eU6pgHHXt2tr8/iPMgGLNrJ21tViDa/agn9X68Ab0eC4dsnoQQgaPUbcRIhNoMSMa" & _
                                                "4aP3lCpYc4u6E09vQI/kEhVygEwGvXbMQ7e5TtbWInD0XjI3Zc4BD95T+jHnW4zn6tbW0iHpCHqevAHFmlkba2uNNbi+HvRzWhMjoDU1uRRAjxpIUGOM1+TaCaaJ4UkzIg5Z8Z7SjxqHa2vbJEUPtD9SEDfODJsT8K0U9LO+LgHDCNw7P4yrD6h98K0UqyCfy7CDbZP0dsfz41KPmIvgNeD7nPj+p/t9zjNHuhDt7l3199S7V//gXQSv5cREfP8TfZzfuQ472Kahw1bjSgibLK6E8Il7JYSzDWZB75UQjlpcCeHHtnUlBHUFRCBwPflaU3E9+VpTOdue7CDH2X6gdnE9+VpTaS2ev0TYwVYEDlvCTp06lfDzqVNpKtU0jrpPx/W7tDAxDzzQPnHw4LDExCsT4pPT0lT60zjq0HH9bRl2sBXQjgKZVHPq1MSfa2rSdUBH+kT0KT4uLUR6enq7+IGDkvonDZrYPzEtXQvpElIGJ6KP82uL" & _
                                                "sIPnOCEna2vHcyHUgT70K34uASYl5eqQAclp49kQakAf+jnftgY7uOfLquh9Ffvnfln1bfGXlQcEQzHqe79q/MSxzGd7voz+osyTsXdfZfEXZRVCZc++Chr3ZOze+1WjnidPnRpihq3mn/8cVV1dHUTjVodE7VGvrqkZXRdQ6md0dfQdMDAqLj41IzYhbVdcYppg+Az1uLhBtk9G13nGp+bHJqQKlZiE1AKnnqGh4dEhncMzQjuHF4d2CRd+0LhRJx3Xz9GxY8eojpcHZVx2eWhxx8tDhMqllwfTeFAG6Rr1pMPVIXVhSxw0KmHw4CCrw1aMoz4gMW202YN+TtvW8BtAML+qPFBR8fXB2q8PfC+++faQHxhHnULqsRNQBHNPWUVFaXlVLYWaC7vAOOoUVE8jAT2fAvaMEbKTJx0t44LeCPSpU/9D++yFnowQJaRWUGBqmVDWgTrhsRMmeMbEp26JSUgZHxMT04Foz9AB9dj4lHw7nghc" & _
                                                "aJeuFZ0jutV27RYlwrtH+4Fx1CmoHjsB9QYzpCIotFNtSOeu/mEnMI466TxWAU1LSzs/PjHtGSNoyWmOXiPofQG1fI3aEn4DFJSsim++Y0OpAh30qofKnrLKJaXlX/sFkgM66DkfAmsRjRnwp5+cLYCG3uylfbY3NjFtCRdGK6DnfGQwI/qCyYWyIaSDnvORoVlxSeeI7mwoVaCDnvORufSy0Kyg0M5sKFWgg57zGZCWFmzOgHFpaY5eI+jNXgq5o95fIn4DX1Z+W2I1Y6pAR4e+2osd0SFridWMqQId6VnP2traeClgji70RPyL2Qsfpn4eHcqWcCG0gmZE7X87DlkpeFYzpkqHmAGp2q93UdhKrGZMFehIv5fzkel4eXCJ1YypAl3Hy0JYzwFJafFSwBy9RtCbvfDhNG0JvwEEhAuiFdCrHip4TykHUAf0qgcwTpHUh9OvrsPshQ9X5wKoQ/VQMd5T8kFkgV71UEFAuCBaAb3qoYL3"
Private Const STR_RES_PNG2          As String = "lHIAdUCvegDv6RFvwLi6DrMXPly9LeE3gHBwIbQCetVDxQ0nH0QON5xuOE38VmsgHFwIrYBe7oef6tnccJqegQwn9zi58OmQ+znP5oaT80Q4uBBaAb3cz3k2N5ymZyDDqT7GtobfAMLBhdAK6FUPFXfm5IPI4c6c7sxp4jeAcHAhtAJ61UPFDScfRA43nG44TfwGvqo8UObk01roVQ+VPWUVZY4+rSU959Pi4YxPLeMCaAnpOR8Zx5/Wkp7zkQnpEl7m5NNa6DkfmY6XB5c5+rSW9JyPG87A4TdAATlnz3OeFCLcDFhzznPCh9O45zmbf54zIS0t3AxYc85zwofTtCX8BrDi56vKbysRPKsZFOOo06xZZWeFEFb80GxYieBZzaAYN4K5r6KqkRVCF9QFrIkrhIhnaZ+90BNW59BhZRUXRBVDZ2M1j+EZn5LvC6jVDNqEFULhlQie1QyKcdTDuoRX2VwhFE2HqlUIntUMinHUoWtkhdAF" & _
                                                "ZsCasULo2auv5i+Y1pZgB30Bzf2y6lsPF6Tyr7/1oO5kbS0CR+8lc7/YV+nBe0qVvd7xXN3a2pp//nOEGVCsmbWzthZrcM0e9DO6OhCOuITUXMLDhjJxULlRtxEiE8PTmEFTC/Ce0o/41HzUnXgicCFduuZSaDxqiLxEVBh1G8E0MQIaFJJL4fPgPSWDx6hr1tb2H5g2wgwo1szaW1s7aJTZg35O29ZgB89x2p+srR1nhs0Jvm+luN8bbGEeoMANSB40ri5sDsC3UtDP+bY12MFWAL7PmVpz6pSxCF6H7/ucSehTfFxaCHwvM2HgoNT4ZN8ieB1Jgybi+5/u9znrYQdbEfg/bBiueGB8kquAKySg7tNx/S4tDA5bcaUDXPEAn8Cq4AoJ7pUQeNhBjpZYreF68rWm0pY9f4mwgyqBfDJdL2e4Xs44V72aAjvo4tJWufol8e/J06tDk6aL6MSpIgG3KVNECMY5fUvCDrq4tDXShWiXPE10" & _
                                                "SZgqUuOfO5WmgvHkaSe7pKefuQ8V2UEXl7ZEWoa4IGGK6M+FUiVpSm1s0gviQs4n0LCDq1evvuyddz/Ysyx39am3l60UJtjHOOpcH8fChQsvXLpsZebSZSs8shf2MY4618cB7eKs7MzFWTmexUtzRB3Yp3EnXlu2bLlwy/aizC3bCz2fbC8SJtg3xqnO9XEMXzj8wuHLbsocvnyk58blI4UJ9o1xqnN9HN1+l3RhwpMDMxOeTPEkPpkiTLCPcdS5Po5uw7td2HdsbGafsXGevnfFCRNjn8ZR5/o4/tSh24XpQT0y04N7eNKDo0Q9tE/jqHN9HONDul32VI+YPX/s2f/UH/skiDpoH+Ooc30c913ePXxcl57H7u/U4/TvwnoIE+xjHHWuz2TMcvHr5KliABdEK6A/EzOo38D69R/+96o179UWf75XHPjukDjx4z/qwH7x7j1i5Zp3a9/bsOE2tVclO3tF4vLcNeWNeeXkri5fkpODv13P" & _
                                                "epi8mZWVtDR7RfnHeQVie+FOUfz5njqw//HGfIH6W9nZfbl+mfxthUlbd+ws91QdEN8fOSFO/L26DuyXV+4XqG8tLNR63Zh9Y+KIFSPLb/n4NnH71jvF2KK76sD+LX+9TaB+fc5I7X9j/GMDk5KfSim/7q0bxKgNo8WtW2+rA/vXvnW9SPpzSnnCEwnax9XnjtikmHv7lydPTBVpUweLwTOvqgP7yRNTBOq974zRPq7/e3m3pEnBUeV/CY4WmUE9xevB9WAf46inX95D6/WnyL4PTuyVULss/mqRP/B6UZw2oo7NtJ+dcJV4uld87VMRfbW/Xw936fnSg52iTz8X1ku80qm3yOx0RR1/6dRLTO7UU6A+oUuviVw/SJgmwtXwJUwWvaLSxQXnUXBxi31Vg0Ngzi+Q+A2sfW/d5p2ffd4gSCpFxbsFdGqvyts5KyfZ8YKO65eh2XHSRxs3i8/+tld8/kWJ+HxPifjb3lLjFvsY/yhvM2ZR" & _
                                                "rdfWHUWTvqr4ukEoVb70VFFAi7ReN+bcNGnMh7eIOwvHWjJmwy0COq5fJuHxpEnD3rhO3PoJBdKCYQuvF9Bx/TJXjI2ZlPjUQCOMVyKUs4i/eG+NfSLx6RQBHdcvMymox+RZQdFiAYXRillBUQI6rl/mqe798pfGX0lhvNGSrPghAjquX4Zmxqpnw6LFvM5XiPkUSBWMp4f1FNBx/UOmi/+QAxc3RaTEpItQTtt3ys9hqMt69HPaQOE3QIeuNftpVjtOwbHim2+/FzQjVqu9KlnLVpTa8crKWfkZ1y9DoSst3FlsBHJv2ZeitOwrUbqv3LjFPsZ37NyFcGq96NC19NDRhjOmCmZQOrzVeg3PGVF62ye3izt3UBAtuG3LHWJ4zk1ar4THkktHrh8tbtlyqyUj140S8Y8na736jI0tTZ1+pRHIK2dfLYZIYB/jadMGi75j47RezwX1KJ1P4VtIIbRiHoV3clCU1uuP0XE1m5OvFbtSh1uy" & _
                                                "iep/jI7V/n7h0PUvNGsiiK936iMW0q0J9jGO+u87R/0/rn/gVBElh63/ZHEFpzMZ8LzoLesTporunC5Q+A3g/eDxHyg4GrKXrzyt9qoszVl5kutVWZqz4gTXL/PW0pyTe0r3GWHcV14pMPOVE7jFPsZRp/egWi8K3UkukCr0HlTrdUP2iJN3bL9T6BiefZPWK+HRgSdvyacQaqDDX61Xn7tiT175ii+Qrw4VV2XUg30jpFTve1es1mtKcNTJN4J7CR2k03o92StO7Eq5QcsfesVpf79+F9adQugN4xsUxjdB2BXGLfa9Ib1C3Ec6tRckTRGxcth6TRMXcToT1GV9/PNCexjfHPwGKFBsgFSgU3tVAumFD34QQBxuVlR9I77+5ltRtf874xb7GEcdOrVXBR/8cGFUgU7tVRn+9k3ijk8pgBqgU3tVaOYUN2++RQt0aq8KPvgxgolAzh0qrp53jbhq/jXeW9rHOOrQqb0qU+g95SIKnw7o" & _
                                                "1F6VP/QeIHYOvEELdGqvCj74QQAXURAXE1nEEt8t9jGOOnRqL0ieIpLlsMXMb3yZJ+qyHv2cLlD4rYJoaqDg05JeCN2+ryqEp/IbI5T76XD424OHjVvsYxx1NZycV1PDyXndsHSEuH3bHVqgk/s4r4RHKZybKIAaoJP7OC+E7irMmBREI5SvXSOGvj7MuMW+MU51NZyc15QQCmcIBVADdHIf52WEM/l6LWo4OS+EDgF8iwKIUC4lsn232Mc46mo4TS85aEDWWKH2qI8pkBgPVB7Iylkhjv3wkxbo5D74tKQXQofZseqbA0Ygv/v+iDh46KhxawSUxlE/K+HcSgHUYCec8RS6MXk3a4FO7uO8jJmTZkfMlGYwhy4gfAHFOOp2wvk8he5NCp8O6OQ+zuvJ3v1FUfJ1WqCT+zgvM5z1wewrcgjcmgF1w0nApyW9EDq8v8RhLE7DIJiHjhw3brGPcdTPdDivz7pR4AMhHdDJfZxX/AQK50YK" & _
                                                "oAbo5D7Oy5g5feEc+po3mNcsvNYbUNo3Dm9thnMqhe4tCp8O6OQ+zuvJXhTOJAqgBujkPs4LocP7SxzGvk0gmMt8AcU+xlH/ZYXzBIVGg+1wMr0qdsNZToeuX9MhLA5n5XBiH+Oon/FwLqFwbqEAaoBO7uO84ickidEfj9ECndzHeRnhnEszJB3CmrNmXTgxexqHttfYCue0kJ5icUhvLdDJfZwXQleYeK0Wp+HE4awcTuy74fQBn5b0MsOJ95cI4/eHjxnhxC32MX42wnnd4uHi1oLbtEAn93Fe8Y9QOP9KAdQAndzHeZnhxPtLhBHBNMG+932nvXBOp9AtofDpgE7u47ye7BknChOGaYFO7uO8zHDi8BVhRDBNsI/xX1w4j1JgdNgNJ9er0qrD+RaFM58CqAE6uY/zMsL5EQVQwxkPZ2hPkRXaWwt0ch/nZYQz/hotbjiNx4nb+gE3nPXYC+cN4pbNt2qBTu7jvAY8nCRGfThaC3Ry" & _
                                                "H+cVyHC+QKFbSuHTAZ3cx3k9QaHbQeHTAZ3cx3m54bTADaeX696kcG6iAGqATu7jvAaMp3CupwBqgE7u47wCGc4XQ3uJt0Ov0AKd3Md5PREdK3YMGKoFOrmP82p74cymQB2n0GiATu6DT0t6navhvPaNG8TNG2/RAp3cx3kNGJ9oLM/TAZ3cx3kFMpwvUeiyKXw6oJP7OC+Ebnv/q7W44TQeJ27rB5ZQUI4c/7sW6OQ++LSk1zkbzoUUzo8pgBqgk/s4rwHjKJwfUAA1QCf3cV6BDOfLFLocCp8O6OQ+zuuJqBixPe4qLdDJfZyXG04L3HB6wbdExvz1Zi3QyX2cV38K3U3vj9QCndzHeQU6nMsofDrshPNxCt2nsVdpgU7u47zccFrghtPLsAUUzo8ogBqgk/s4r/4PUTjfowBqgE7u47wCGs4wCmcYBVADdHIf5+UN5xAtbjiNx4nb+gE3nPXYCec1r18nRn84Rgt0ch/n1f+hBDHi" & _
                                                "3Zu0QCf3cV6BDOcMCt3ysD5aoJP7OK/He/QT22Ku1AKd3Md5tblwLn57hTh87O9aoJP74NOiXudqOF+jcG6gAGqATu7jvPo/mCBGrKUAaoBO7uO8AhnOmRS6FRQ+HdDJfZyXEc5+g7W44TQeJ27rB9xw1mMnnEPnX8ue7lCBTu7jvOIodDeuGaEFOrmP8wpkOGdR6FZS+HRAJ/dxXo917ye29h2sBTq5j/Nyw2mBG04vRjjXUQA12Arnf1M436EAaoBO7uO8AhlOXE0gl8KnAzq5j/N6rHtfsbXPIC3QyX2clxtOC9xwehk6bxh7ukMFOrmP84p7IF4MX3WjFujkPs4rsOHsLVaF9dUCndzHeSF0n/RJ0+KG03icuK0fWPz2cnHo2I9aoJP74NOiXudqODMpnO9TADVAJ/dxXkY4cymAGs50OHFlu9X0C68DOrmP83qsG4XzCgqgBujkPs6rbYbzKIVGg+1wMr0qrTmcV88dxp7uUIFO" & _
                                                "7uO8Yn8fL25YOVwLdHIf5xXIcM6m0L1Dv/A6oJP7OK9Hu/URW3qnaoFO7uO83HBa4IbTy9UZFM53KYAaoJP7OC8jnCsogBrOdDjnUOjW0C+8DujkPs7r0UgKZ68ULdDJfZyXG04L3HB6uSrjGvZ0hwp0ch/nFXv/AHH98hu0QCf3cV6BDOerna4Qazv10wKd3Md5PRp5hSjoOVALdHIf59XmwvkWBeV7CowO6OQ++LSk1zkbzlcpnGsogBqgk/s4LyOcyyiAGs5GON+l8OmwFc4ICmc0BVADdHIf5+WG0wI3nF6GzBnKnu5QgU7u47xi7hsgrsu5Xgt0ch/nFfBwdqYAarATzgkUuvzoZC3QyX2cV9sL51IK1BEKjQbo5D74tKTXORvO2RTO1RRADdDJfZyXEc5sCqCGMx3ODArdexQ+HdDJfZzXhIjeIj8qSQt0ch/n5YbTAjecXnDVdO5cpAp0ch/nFfO7/uLat6/TAp3cx3kFMpxz" & _
                                                "KXTvU/h0QCf3cV4TwnuLzT0StUAn93FebS6cby5dJg4e+UELdHIffFrS65wOJ3MuUsVWOO+lcGZRADVAJ/dxXoEMZyb9gn/QOUYLdHIf52WEszsFUIMbTuNx4rZ+wA1nPbbC+Zer2HORKtDJfZxXPwrdsCXXaoFO7uO8AhnOefQLvo7CpwM6uY/zeqRrL7GpW4IW6OQ+zssNpwVuOL1cOYvCyZyLVIFO7uO8+t1D4VxMAdQAndzHeQU6nOspfDrshzNeixtO43Hitn7gzSwK1GEKjQbo5D74tKTXuRpO/J1L7lykCnRyH+fV7544cc1bFCQN0Ml9nFcgwzm/cx+xoUuMFujkPs4LocuLjNfihtN4nLitH0BQvqPA6LAbTq5XpVWHcwaFkzkXqQKd3Md59bubwvkmBUkDdHIf5xXIcL7Wua/4sEusFujkPs7rkS49RV7EAC3QyX2clxtOC9xwehk0Ywh7LlIFOrmP8zLCuYiCpOFMh/N1" & _
                                                "Ct1HFD4d0Ml9nJcRzvD+WtxwGo8Tt/UDbjjrsRXOlymczLlIFejkPs4LQRn6BoVJg51ABTKcCyh0f6Xw6YBO7uO8Hu4cLTZ2jdMCndzHef3iw6kONDVQHIH0MsNp/iEjNZxWf8iIo6nh5Eh76Ur2XKQKdGqvihHOhRQmDWqgOMxwmn/ISA2n1R8y4kDoPqbw6VDDydHUcHKY4TT/kJEaTqs/ZGSSMEUkykFz+sdz0c/pAoXfwJKc3BrPN9+zITLxfH1QQKf2qizKWlZqxws6rl9mcVZO6d6yr4wQsn8CkMb3ln4poOP6ZbZsLyw9fOwHNpAmqEPH9cukvTC4FB/SXLuUQmgB6tBx/TJ9xsaWDpl7FRtIkyEZVwnouH4ZaAbPGuINJ/cnAGl88MwhtrwWdOpTuqFzPzaQJqhDx/XLPN4pqub9LjFsIE3eozp0XL/M/WHdT88P62WE0OpPAKIOHdc/8HnRWw6b0z87j35OFyj8Bpbnri0r" & _
                                                "2r2XDZJJ4Wd7BHRqr8qbS3IW2PGCjuuXeSsre+GnO4qMv8Np/vFczJrmH8/F+LbthQI6rl9my/aiBRWk50Jpgjp0XL9M2gtXLrj69aFsKE1Qh47rl+k7NnZB6tRBbChNUqakCei4fhloUiYPavDHc+tmTd8fzx2Ybs9rYee+C9ZoAoU6dFy/zFOdossWdenLepi8QXXouH6ZB8N6nHixkzec5h/P9c6afY19jL9Adei4/uRpoosctgGasPWfLK6Q9ejndIHCb2DFinc756xYc+rzkq/YMGE8Z+WamlWrVoWpvSpz5y69KCtnVUVjXlk5uSWLFi36T65fZtGi1f+5NGdlxc7dfxNfHzhozJYIpjFr0j7GqW7Lq6Cg4KJPdxRX7KdeLpj7vzsstu7YWZKXl6f1Sp6WfFHay4MrEBwumEMX0qz50uCStPQ0rVev23td1O/e/hWk9wslGPTiYBFzb1xJnxF9bHnF3DugImXaYJolvbMlgumd" & _
                                                "NYeJVBqnui2vuaG9LlrUuU+F1Yz3Ps2cb3TqWzKrk97rf0MjOv+5U9Spt7v0Y72yjWD2rPlTp27a36/HQ7v0eZBmxVlGQL2zJYLpnTX7Ghcce7BT1OnHLo/Ap0t+/dfPF/+WOEWkyIGLSa8O5bTx6T93knUJU0+lXp8u/o3TBgp2cPXq1Zcty127Jyt75Sm8HzTBPsZR5/o4Fi5ceOGiJTnz3lyyzCN7YR/jqHN9HNAuzsrOpENXD95b1oF9GnfitWXLlgtpZsykQ1cP3luaYN8YpzrXx5H0QtKFdNiamfrilZ60F68UJsY+jaPO9XF0G97tQprNMvuMjfPg/aCJsU/jqHN9HIH0yrg46oIFXfrOW9C5j2chBcgE+xhHnevjeLpD945Pd+q5hw5dT+G9pQn2MY4618fxcIfuXceFRh27j0KK95Ym2Mc46lyfSfw00SB0IG5ybVxcuugYPlv8BrfYVzXJ0wUb4kDCDrq4tBXS00W7hCmi" & _
                                                "vxq+xkiaUou/svQr1SvQsIMuLm2JtFniP7kQciRMFalJLwjbRxzNgR10cWlrYAZNnnayC8JnFUp8AAQd198SsIMuLm2Vq18S/54yRYQkTRfRiVNFAm6Tp1eHYpzTtyTsoEogV0G4Xs5wvZxxrno1BXaQoyUeqOvJ15pKW/b8JcIOurgEEhwS4tDQPFTEezjzkBGHkGfjkLE1wA+ed96v7ML1c6h9jcH1c6h9jcH1c6h9jcH1c6h9jcH1c6h9jcH1c6h9jcH1q6QLfMgiLD9kMfF+2HLyjH7Y0hpouMO8CHaRfWRUnRM4P6DqnMD5AVXnBM4PqDoncH5A1TmB8wOqzgmcH0jLEBc05fzhmTpN0Rqo/4f0hFd+c/D2ym8OF1btP1xddeCI8IPGUS+vPHi73Ccbq54zckrmjJ+9+6dbny8Soyb7g3HUXyCd3NeY5xvbn5jz3MZrfnpyQ0/x6PpIPzCO+mvb/tCopwlqx957ac6xtx766Ye5" & _
                                                "I8WPr97gB8ZRP/TuS4Yn5wNQM9n18BNztqZe89Pm8J5iU2ikHxhHvfDhxh+nXDu46+k5328b9dOh/ARxeHOMHxhH/cCu/7Ht+WHmd3OWT/zup7ceOiAW/f5bPzCO+nrSyX2q55jl4tfJU8UALoA60OfOoF68P6Qnurzyu3H7Dx4T1T/XiNOnsZjff8M46tBBL/dT2c9z6tK92VwgrYBe7uc8M7dNyOYCaQX0cr/pKXPonanZXCCtgJ7zke9nxwMTsrlAWgG93M95Htj5p2wukFZAL/dznh/M+S6bC6QV0Mv9pidImCbC1dAlTBa9otLFBedRcI37plvsY1zVtvSC8taC90f9k9yu4ptDu0/W1NKwfoMOevSZHjTs5/nQK7uruRBaAT36TA/Oc1LeVdVcCK2AHn2mh+kpc/zNB6q5EFoBPedj3gfR7pPkq6q5EFoBPfpMD87z+60jqrkQWgE9+kwPznPZU99VcyG0Anr0mR6m55Dp4j/k" & _
                                                "oMVNESkx6Y2vQ+075ecw6OQ++HDatoT3R/2L9OvK/UdOWM2Y6gYd9OgzPWjYz/O2qTtPcyG0Anr0mR6c55Mf9j7NhdAK6NFnepieMj/MG3WaC6EV0HM+5n0Qv86P7H2aC6EV0KPP9OA8D+UnnuZCaAX06DM9OM/F4749zYXQCujRZ3qYngOniig5ZPialVlrDHxdS+5LmCq6c7q2hPeH9wnG/wXPx3tKJxv06PP1qy/8Oe8p0xRP2vx84O+7nzb3fCZNEbFyyHRfYDZRv8gc/7zox+naEt4fzX+R2vv6A/nCnxFPmaZ40ubnA3/f/bS55zN5ikiWQ6a79IeJegkQ+HC6tgQ2/DRfpPbNfZEI7LcaT9rqnoymeNImP5kt9jh9/z7nPeWAAZLWPT861F74qZq2hPdHAF8kGmpVnjJN8aTNzwf+vvtpc8+n04tmmZzpi2e1Brw/WskLj3/7xgLmKdMUT9r8fODvu58293w6vWiWyZm+eFZr" & _
                                                "wPtDeZGcfFpr94U/Vz1lmuJJm58P/H330+aeT6cXzTI50xfPag14f9S/SOc38Txno58EnsueMk3xpH/6+cDfdz9t7vl0ctEsk7Nx8azWgPeH90UC55fsq3rEyQoh6NFnelC5VXnKNMWThvx8zPsg2uTzqbtoFjRn++JZrQHvj/oXCSeV/8/uL7662/P1oaLG1tai/vkX++6B3tfX4EUy94lz2lMGGqeeFj7Nepwt4enra/Dfbu4TAX0+m3LRLJMzdfGs1oD3R/2LZBziEL8h/pX4t0ZAHTrojUMbUGfcSjxbAvM+CNuPk/ORoS3g/+3mPhHw59PJRbNM8NUx91sp9dT/o+ELhf8j4snHG34rUIfO8gUyx4lz2rMlMO+LsPU4OQ8V2gL+326OEwF/PnUXzTLxfp/zzF48qzVQ/4/6F0l+sXQ06JGNW5NnS6DeJ8E9rjo4DxXaHHn6aNBzJjxV1ItmIZBn++JZrQH/AeWJt4PqoaLq7cD5" & _
                                                "yKh6O3A+LY36GKzgeq1Qe+3A+cioejtwPi6Bgx00sfti2NGYyFoVTm8H1UeG058t1Mcmw+ntoPrIcHo7qD4ynN6lZWAHOWijG77WVFxPvtZU2rLnLw9x3v8HFzOfyZOzXlwAAAAASUVORK5CYII="
Private Const LNG_RES_WIDTH         As Long = 231
Private Const LNG_FLAT_TOP          As Long = 0
Private Const LNG_FLAT_WIDTH        As Long = 21
Private Const LNG_FLAT_HEIGHT       As Long = 21
Private Const LNG_BUTTON_TOP        As Long = 84
Private Const LNG_BUTTON_WIDTH      As Long = 19
Private Const LNG_BUTTON_HEIGHT     As Long = 52
Private Const LNG_CARD_TOP          As Long = 136
Private Const LNG_CARD_WIDTH        As Long = 21
Private Const LNG_CARD_HEIGHT       As Long = 23

'--- design-time
Private m_eStyle                As UcsNineButtonStyleEnum
Private m_sngOpacity            As Single
Private m_sngAnimationDuration  As Single
Private m_uButton(0 To ucsBstLast) As UcsNinePatchType
Private m_sCaption              As String
Private WithEvents m_oFont      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_clrFore               As Long
'--- run-time
Private m_eState                As UcsNineButtonStateEnum
Private m_hPrevBitmap           As Long
Private m_hPrevAttributes       As Long
Private m_hBitmap               As Long
Private m_hAttributes           As Long
Private m_hFocusBitmap          As Long
Private m_hFocusAttributes      As Long
Private m_nDownButton           As Integer
Private m_nDownShift            As Integer
Private m_sngDownX              As Single
Private m_sngDownY              As Single
Private m_dblAnimationStart     As Double
Private m_dblAnimationEnd       As Double
Private m_sngAnimationOpacity1  As Single
Private m_sngAnimationOpacity2  As Single
Private m_hFont                 As Long
Private m_bShown                As Boolean
#If ImplHasTimers Then
    Private m_uTimer            As FireOnceTimerData
#End If

Private Type UcsNinePatchType
    ImageArray()        As Byte
    ImagePatch          As cNinePatch
    ImageOpacity        As Single
    TextFont            As StdFont
    TextFlags           As UcsNineButtonTextFlagsEnum
    TextColor           As OLE_COLOR
    TextOpacity         As Single
    TextOffsetX         As Single
    TextOffsetY         As Single
    ShadowColor         As OLE_COLOR
    ShadowOpacity       As Single
    ShadowOffsetX       As Single
    ShadowOffsetY       As Single
End Type

Private Enum UcsNineButtonResIndex
    ucsIdxFlatPrimaryNormal = 0
    ucsIdxFlatPrimaryOutline
    ucsIdxFlatPrimaryHover
    ucsIdxFlatPrimaryPressed
    ucsIdxFlatPrimaryFocus
    ucsIdxFlatSecondaryNormal
    ucsIdxFlatSecondaryOutline
    ucsIdxFlatSecondaryHover
    ucsIdxFlatSecondaryHoverOutline
    ucsIdxFlatSecondaryPressed
    ucsIdxFlatSecondaryFocus
    ucsIdxFlatSuccessNormal
    ucsIdxFlatSuccessOutline
    ucsIdxFlatSuccessHover
    ucsIdxFlatSuccessPressed
    ucsIdxFlatSuccessFocus
    ucsIdxFlatDangerNormal
    ucsIdxFlatDangerOutline
    ucsIdxFlatDangerHover
    ucsIdxFlatDangerPressed
    ucsIdxFlatDangerFocus
    ucsIdxFlatWarningNormal
    ucsIdxFlatWarningOutline
    ucsIdxFlatWarningHover
    ucsIdxFlatWarningPressed
    ucsIdxFlatWarningFocus
    ucsIdxFlatInfoNormal
    ucsIdxFlatInfoOutline
    ucsIdxFlatInfoHover
    ucsIdxFlatInfoPressed
    ucsIdxFlatInfoFocus
    ucsIdxFlatLightNormal
    ucsIdxFlatLightOutline
    ucsIdxFlatLightHover
    ucsIdxFlatLightPressed
    ucsIdxFlatLightFocus
    ucsIdxFlatDarkNormal
    ucsIdxFlatDarkOutline
    ucsIdxFlatDarkHover
    ucsIdxFlatDarkPressed
    ucsIdxFlatDarkFocus
    ucsIdxButtonDefNormal = 0
    ucsIdxButtonDefHover
    ucsIdxButtonDefPressed
    ucsIdxButtonDisabled
    ucsIdxButtonGreenNormal
    ucsIdxButtonGreenHover
    ucsIdxButtonGreenPressed
    ucsIdxButtonRedNormal
    ucsIdxButtonRedHover
    ucsIdxButtonRedPressed
    ucsIdxButtonFocus
    ucsIdxCardDefault = 0
    ucsIdxCardPrimary
    ucsIdxCardSuccess
    ucsIdxCardOrange
    ucsIdxCardDanger
    ucsIdxCardWarning
    ucsIdxCardPurple
    ucsIdxCardFocus
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Function PrintError(sFunction As String) As VbMsgBoxResult
    Debug.Print STR_MODULE_NAME & "." & sFunction & ": " & Err.Description, Timer
End Function

'Private Function RaiseError(sFunction As String) As VbMsgBoxResult
'    Err.Raise Err.Number, STR_MODULE_NAME & "." & sFunction & vbCrLf & Err.Source, Err.Description
'End Function

'=========================================================================
' Properties
'=========================================================================

Property Get Style() As UcsNineButtonStyleEnum
    Style = m_eStyle
End Property

Property Let Style(ByVal eValue As UcsNineButtonStyleEnum)
    If m_eStyle <> eValue Then
        m_eStyle = eValue
        pvSetStyle eValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Property Let Enabled(ByVal bValue As Boolean)
    If UserControl.Enabled <> bValue Then
        UserControl.Enabled = bValue
        pvState(ucsBstDisabled) = Not bValue
    End If
    PropertyChanged
End Property

Property Get Opacity() As Single
    Opacity = m_sngOpacity
End Property

Property Let Opacity(ByVal sngValue As Single)
    If m_sngOpacity <> sngValue Then
        m_sngOpacity = sngValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get AnimationDuration() As Single
    AnimationDuration = m_sngAnimationDuration
End Property

Property Let AnimationDuration(sngValue As Single)
    m_sngAnimationDuration = sngValue
    PropertyChanged
End Property

Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = m_sCaption
End Property

Property Let Caption(sValue As String)
    If m_sCaption <> sValue Then
        m_sCaption = sValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
    Set Font = m_oFont
End Property

Property Set Font(oValue As StdFont)
    If Not m_oFont Is oValue Then
        Set m_oFont = oValue
        pvPrepareFont m_oFont, m_hFont
        Repaint
    End If
    PropertyChanged
End Property

Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_clrFore
End Property

Property Let ForeColor(ByVal clrValue As OLE_COLOR)
    If m_clrFore <> clrValue Then
        m_clrFore = clrValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get ButtonState() As UcsNineButtonStateEnum
    ButtonState = m_eState
End Property

Property Let ButtonState(ByVal eState As UcsNineButtonStateEnum)
    pvState(m_eState And Not eState) = False
    pvState(eState And Not m_eState) = True
    PropertyChanged
End Property

Property Get ButtonImageArray(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Byte()
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonImageArray = m_uButton(eState).ImageArray
End Property

Property Let ButtonImageArray(Optional ByVal eState As UcsNineButtonStateEnum = -1, baValue() As Byte)
    Dim oPatch          As cNinePatch
    
    If eState < 0 Then
        eState = m_eState
    End If
    m_uButton(eState).ImageArray = baValue
    Set oPatch = New cNinePatch
    If oPatch.LoadFromByteArray(baValue) Then
        Set m_uButton(eState).ImagePatch = oPatch
    Else
        Set m_uButton(eState).ImagePatch = Nothing
    End If
    Repaint
    PropertyChanged
End Property

Property Get ButtonImageOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonImageOpacity = m_uButton(eState).ImageOpacity
End Property

Property Let ButtonImageOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).ImageOpacity <> sngValue Then
        m_uButton(eState).ImageOpacity = sngValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get ButtonTextFont(Optional ByVal eState As UcsNineButtonStateEnum = -1) As StdFont
    If eState < 0 Then
        eState = m_eState
    End If
    Set ButtonTextFont = m_uButton(eState).TextFont
End Property

Property Set ButtonTextFont(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal oValue As StdFont)
    If eState < 0 Then
        eState = m_eState
    End If
    If Not m_uButton(eState).TextFont Is oValue Then
        Set m_uButton(eState).TextFont = oValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get ButtonTextFlags(Optional ByVal eState As UcsNineButtonStateEnum = -1) As UcsNineButtonTextFlagsEnum
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonTextFlags = m_uButton(eState).TextFlags
End Property

Property Let ButtonTextFlags(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal eValue As UcsNineButtonTextFlagsEnum)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).TextFlags <> eValue Then
        m_uButton(eState).TextFlags = eValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get ButtonTextColor(Optional ByVal eState As UcsNineButtonStateEnum = -1) As OLE_COLOR
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonTextColor = m_uButton(eState).TextColor
End Property

Property Let ButtonTextColor(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal clrValue As OLE_COLOR)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).TextColor <> clrValue Then
        m_uButton(eState).TextColor = clrValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get ButtonTextOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonTextOpacity = m_uButton(eState).TextOpacity
End Property

Property Let ButtonTextOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).TextOpacity <> sngValue Then
        m_uButton(eState).TextOpacity = sngValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get ButtonTextOffsetX(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonTextOffsetX = m_uButton(eState).TextOffsetX
End Property

Property Let ButtonTextOffsetX(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).TextOffsetX <> sngValue Then
        m_uButton(eState).TextOffsetX = sngValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get ButtonTextOffsetY(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonTextOffsetY = m_uButton(eState).TextOffsetY
End Property

Property Let ButtonTextOffsetY(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).TextOffsetY <> sngValue Then
        m_uButton(eState).TextOffsetY = sngValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get ButtonShadowColor(Optional ByVal eState As UcsNineButtonStateEnum = -1) As OLE_COLOR
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonShadowColor = m_uButton(eState).ShadowColor
End Property

Property Let ButtonShadowColor(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal clrValue As OLE_COLOR)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).ShadowColor <> clrValue Then
        m_uButton(eState).ShadowColor = clrValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get ButtonShadowOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonShadowOpacity = m_uButton(eState).ShadowOpacity
End Property

Property Let ButtonShadowOpacity(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).ShadowOpacity <> sngValue Then
        m_uButton(eState).ShadowOpacity = sngValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get ButtonShadowOffsetX(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonShadowOffsetX = m_uButton(eState).ShadowOffsetX
End Property

Property Let ButtonShadowOffsetX(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).ShadowOffsetX <> sngValue Then
        m_uButton(eState).ShadowOffsetX = sngValue
        Repaint
    End If
    PropertyChanged
End Property

Property Get ButtonShadowOffsetY(Optional ByVal eState As UcsNineButtonStateEnum = -1) As Single
    If eState < 0 Then
        eState = m_eState
    End If
    ButtonShadowOffsetY = m_uButton(eState).ShadowOffsetY
End Property

Property Let ButtonShadowOffsetY(Optional ByVal eState As UcsNineButtonStateEnum = -1, ByVal sngValue As Single)
    If eState < 0 Then
        eState = m_eState
    End If
    If m_uButton(eState).ShadowOffsetY <> sngValue Then
        m_uButton(eState).ShadowOffsetY = sngValue
        Repaint
    End If
    PropertyChanged
End Property

Private Property Get pvState(ByVal eState As UcsNineButtonStateEnum) As Boolean
    pvState = (m_eState And eState) <> 0
End Property

Private Property Let pvState(ByVal eState As UcsNineButtonStateEnum, ByVal bValue As Boolean)
    Dim ePrevState      As UcsNineButtonStateEnum
    
    ePrevState = m_eState
    If bValue Then
        If (m_eState And eState) <> eState Then
            m_eState = m_eState Or eState
        End If
    Else
        If (m_eState And eState) <> 0 Then
            m_eState = m_eState And Not eState
        End If
    End If
    If ePrevState <> m_eState And m_bShown Then
        pvStartAnimation m_sngAnimationDuration, _
            m_sngOpacity * m_uButton(pvGetEffectiveState(ePrevState)).ImageOpacity, _
            m_sngOpacity * m_uButton(pvGetEffectiveState(m_eState)).ImageOpacity
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Repaint()
    Const FUNC_NAME     As String = "Repaint"
    
    On Error GoTo EH
    If m_bShown Then
        pvPrepareBitmap m_eState, m_hFocusBitmap, m_hBitmap
        pvPrepareAttribs m_sngOpacity * m_uButton(pvGetEffectiveState(m_eState)).ImageOpacity, m_hAttributes
        Refresh
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Public Sub CancelMode()
    pvState(ucsBstHoverPressed) = False
End Sub

Friend Sub frTimer()
    pvAnimateState DateTimer - m_dblAnimationStart, m_sngAnimationOpacity1, m_sngAnimationOpacity2
End Sub

Private Function pvGetEffectiveState(ByVal eState As UcsNineButtonStateEnum) As UcsNineButtonStateEnum
    If (eState And ucsBstDisabled) <> 0 Then
        If Not m_uButton(ucsBstDisabled).ImagePatch Is Nothing Then
            pvGetEffectiveState = ucsBstDisabled
            Exit Function
        End If
    End If
    eState = eState And Not ucsBstFocused
    If m_uButton(eState).ImagePatch Is Nothing Then
        eState = eState And Not ucsBstHover
    End If
    If m_uButton(eState).ImagePatch Is Nothing Then
        eState = eState And Not ucsBstPressed
    End If
    If m_uButton(eState).ImagePatch Is Nothing Then
        eState = ucsBstNormal
    End If
    pvGetEffectiveState = eState
End Function

Private Function pvPrepareBitmap( _
            ByVal eState As UcsNineButtonStateEnum, _
            hFocusBitmap As Long, _
            hBitmap As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareBitmap"
    Dim hGraphics       As Long
    Dim hNewFocusBitmap As Long
    Dim hNewBitmap      As Long
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim hBrush          As Long
    Dim hShadowBrush    As Long
    Dim lOffset         As Long
    Dim uRect           As RECTF
    Dim hStringFormat   As Long
    Dim hFont           As Long
    
    On Error GoTo EH
    If (eState And ucsBstFocused) <> 0 And (eState And ucsBstHoverPressed) <> ucsBstHoverPressed Then
        If hFocusBitmap = 0 Then
            With m_uButton(ucsBstFocused)
                If Not .ImagePatch Is Nothing Then
                    If GdipCreateBitmapFromScan0(ScaleWidth, ScaleHeight, ScaleWidth * 4, PixelFormat32bppARGB, ByVal 0, hNewFocusBitmap) <> 0 Then
                        GoTo QH
                    End If
                    If GdipGetImageGraphicsContext(hNewFocusBitmap, hGraphics) <> 0 Then
                        GoTo QH
                    End If
                    If Not .ImagePatch.DrawToGraphics(hGraphics, 0, 0, ScaleWidth, ScaleHeight) Then
                        GoTo QH
                    End If
                    Call GdipDeleteGraphics(hGraphics)
                    hGraphics = 0
                End If
            End With
        Else
            hNewFocusBitmap = hFocusBitmap
        End If
    End If
    With m_uButton(pvGetEffectiveState(eState))
        If Not .TextFont Is Nothing Then
            If Not pvPrepareFont(.TextFont, hFont) Then
                GoTo QH
            End If
        Else
            hFont = m_hFont
        End If
        If GdipCreateBitmapFromScan0(ScaleWidth, ScaleHeight, ScaleWidth * 4, PixelFormat32bppARGB, ByVal 0, hNewBitmap) <> 0 Then
            GoTo QH
        End If
        If GdipGetImageGraphicsContext(hNewBitmap, hGraphics) <> 0 Then
            GoTo QH
        End If
        If GdipSetTextRenderingHint(hGraphics, TextRenderingHintClearTypeGridFit) <> 0 Then
            GoTo QH
        End If
        If Not .ImagePatch Is Nothing Then
            If Not .ImagePatch.DrawToGraphics(hGraphics, 0, 0, ScaleWidth, ScaleHeight) Then
                GoTo QH
            End If
            .ImagePatch.CalcClientRect ScaleWidth, ScaleHeight, lLeft, lTop, lWidth, lHeight
        Else
            lWidth = ScaleWidth
            lHeight = ScaleHeight
        End If
        RaiseEvent OwnerDraw(hGraphics, hFont, eState, lLeft, lTop, lWidth, lHeight)
        If hFont <> 0 Then
            If GdipCreateSolidFill(pvTranslateColor(IIf(.TextColor = DEF_TEXTCOLOR, m_clrFore, .TextColor), .TextOpacity), hBrush) <> 0 Then
                GoTo QH
            End If
            If GdipCreateStringFormat(0, 0, hStringFormat) <> 0 Then
                GoTo QH
            End If
            If GdipSetStringFormatAlign(hStringFormat, .TextFlags And 3) <> 0 Then
                GoTo QH
            End If
            If GdipSetStringFormatLineAlign(hStringFormat, (.TextFlags \ 4) And 3) <> 0 Then
                GoTo QH
            End If
            If GdipSetStringFormatFlags(hStringFormat, .TextFlags \ 16) <> 0 Then
                GoTo QH
            End If
            lOffset = .TextOffsetX * -((eState And ucsBstHoverPressed) = ucsBstHoverPressed)
            uRect.Left = lLeft + lOffset
            lOffset = .TextOffsetY * -((eState And ucsBstHoverPressed) = ucsBstHoverPressed)
            uRect.Top = lTop + lOffset
            uRect.Right = lWidth
            uRect.Bottom = lHeight
            If .ShadowOffsetX <> 0 Or .ShadowOffsetY <> 0 Or .ImagePatch Is Nothing Then
                If GdipCreateSolidFill(pvTranslateColor(.ShadowColor, .ShadowOpacity), hShadowBrush) <> 0 Then
                    GoTo QH
                End If
                If GdipSetTextRenderingHint(hGraphics, TextRenderingHintAntiAlias) <> 0 Then
                    GoTo QH
                End If
                uRect.Left = uRect.Left + .ShadowOffsetX
                uRect.Top = uRect.Top + .ShadowOffsetY
                If GdipDrawString(hGraphics, StrPtr(m_sCaption), -1, hFont, uRect, hStringFormat, hShadowBrush) <> 0 Then
                    GoTo QH
                End If
                uRect.Left = uRect.Left - .ShadowOffsetX
                uRect.Top = uRect.Top - .ShadowOffsetY
            End If
            If GdipDrawString(hGraphics, StrPtr(m_sCaption), -1, hFont, uRect, hStringFormat, hBrush) <> 0 Then
                GoTo QH
            End If
        End If
    End With
    '--- commit
    If hNewFocusBitmap <> hFocusBitmap Then
        If hFocusBitmap <> 0 Then
            Call GdipDisposeImage(hFocusBitmap)
            hFocusBitmap = 0
        End If
        hFocusBitmap = hNewFocusBitmap
    End If
    hNewFocusBitmap = 0
    If hNewBitmap <> hBitmap Then
        If hBitmap <> 0 Then
            Call GdipDisposeImage(hBitmap)
            hBitmap = 0
        End If
        hBitmap = hNewBitmap
    End If
    hNewBitmap = 0
    '-- success
    pvPrepareBitmap = True
QH:
    If hFont <> 0 And hFont <> m_hFont Then
        Call GdipDeleteFont(hFont)
        hFont = 0
    End If
    If hStringFormat <> 0 Then
        Call GdipDeleteStringFormat(hStringFormat)
        hStringFormat = 0
    End If
    If hShadowBrush <> 0 Then
        Call GdipDeleteBrush(hShadowBrush)
        hShadowBrush = 0
    End If
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
        hBrush = 0
    End If
    If hNewFocusBitmap <> 0 Then
        Call GdipDisposeImage(hNewFocusBitmap)
        hNewFocusBitmap = 0
    End If
    If hNewBitmap <> 0 Then
        Call GdipDisposeImage(hNewBitmap)
        hNewBitmap = 0
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvTranslateColor(ByVal clrValue As OLE_COLOR, Optional ByVal Alpha As Single = 1) As Long
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
    Call CopyMemory(pvTranslateColor, uQuad, 4)
End Function

Private Function pvPrepareAttribs(ByVal sngAlpha As Single, hAttributes As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareAttribs"
    Dim clrMatrix(0 To 4, 0 To 4) As Single
    Dim hTempAttributes As Long
    
    On Error GoTo EH
    If GdipCreateImageAttributes(hTempAttributes) <> 0 Then
        GoTo QH
    End If
    clrMatrix(0, 0) = 1
    clrMatrix(1, 1) = 1
    clrMatrix(2, 2) = 1
    clrMatrix(3, 3) = sngAlpha
    clrMatrix(4, 4) = 1
    If GdipSetImageAttributesColorMatrix(hTempAttributes, 0, 1, clrMatrix(0, 0), clrMatrix(0, 0), 0) <> 0 Then '
        GoTo QH
    End If
    '--- commit
    If hAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hAttributes)
        hAttributes = 0
    End If
    hAttributes = hTempAttributes
    hTempAttributes = 0
    '--- success
    pvPrepareAttribs = True
QH:
    If hTempAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hTempAttributes)
        hTempAttributes = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPrepareFont(oFont As StdFont, hFont As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareFont"
    Dim hFamily         As Long
    Dim eStyle          As FontStyle
    
    On Error GoTo EH
    If hFont <> 0 Then
        Call GdipDeleteFont(hFont)
        hFont = 0
    End If
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFamily) <> 0 Then
        If GdipGetGenericFontFamilySansSerif(hFamily) <> 0 Then
            GoTo QH
        End If
    End If
    eStyle = FontStyleBold * -oFont.Bold _
        Or FontStyleItalic * -oFont.Italic _
        Or FontStyleUnderline * -oFont.Underline _
        Or FontStyleStrikeout * -oFont.Strikethrough
    If GdipCreateFont(hFamily, oFont.Size, eStyle, UnitPoint, hFont) <> 0 Then
        GoTo QH
    End If
    '--- success
    pvPrepareFont = True
QH:
    If hFamily <> 0 Then
        Call GdipDeleteFontFamily(hFamily)
        hFamily = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvRegisterCancelMode(oCtl As Object) As Boolean
    On Error GoTo QH
    Parent.RegisterCancelMode oCtl
    '--- success
    pvRegisterCancelMode = True
QH:
End Function

Private Function pvHitTest(ByVal X As Single, ByVal Y As Single) As HitResultConstants
    Const FUNC_NAME     As String = "pvHitTest"
    Dim clrCurrent      As Long
    Dim lAlpha          As Long
    
    On Error GoTo EH
    pvHitTest = vbHitResultHit
    If GdipBitmapGetPixel(m_hBitmap, X, Y, clrCurrent) <> 0 Then
        GoTo QH
    End If
    Call CopyMemory(lAlpha, ByVal UnsignedAdd(VarPtr(clrCurrent), 3), 1)
    If lAlpha < 255 Then
        If lAlpha > 0 Then
            pvHitTest = vbHitResultTransparent
        Else
            pvHitTest = vbHitResultOutside
        End If
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvStartAnimation(ByVal sngDuration As Single, ByVal sngOpacity1 As Single, ByVal sngOpacity2 As Single) As Boolean
    Const FUNC_NAME     As String = "pvStartAnimation"
    Dim hNewBitmap      As Long
    
    On Error GoTo EH
    If Not pvPrepareBitmap(m_eState, m_hFocusBitmap, hNewBitmap) Then
        GoTo QH
    End If
    m_hPrevBitmap = m_hBitmap
    m_hBitmap = hNewBitmap
    hNewBitmap = 0
    m_dblAnimationStart = DateTimer
    m_dblAnimationEnd = m_dblAnimationStart + sngDuration
    m_sngAnimationOpacity1 = sngOpacity1
    m_sngAnimationOpacity2 = sngOpacity2
    pvAnimateState 0, m_sngAnimationOpacity1, m_sngAnimationOpacity2
QH:
    If hNewBitmap <> 0 Then
        Call GdipDisposeImage(hNewBitmap)
        hNewBitmap = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvAnimateState(dblElapsed As Double, ByVal sngOpacity1 As Single, ByVal sngOpacity2 As Single) As Boolean
    Const FUNC_NAME     As String = "pvAnimateState"
    Dim sngAlpha1       As Single
    Dim sngAlpha2       As Single
    Dim dblHalf         As Double

    On Error GoTo EH
    sngAlpha1 = sngOpacity1 * 0
    sngAlpha2 = sngOpacity2 * 1
    #If ImplHasTimers Then
        dblHalf = (m_dblAnimationEnd - m_dblAnimationStart) / 2
        If dblHalf > DBL_EPLISON Then
            If dblElapsed < dblHalf Then
                sngAlpha1 = sngOpacity1 * 1
                sngAlpha2 = sngOpacity2 * dblElapsed / dblHalf
            ElseIf dblElapsed < 2 * dblHalf Then
                sngAlpha1 = sngOpacity1 * (2 * dblHalf - dblElapsed) / dblHalf
                sngAlpha2 = sngOpacity2 * 1
            End If
        End If
    #End If
    If Not pvPrepareAttribs(sngAlpha1, m_hPrevAttributes) Then
        GoTo QH
    End If
    If Not pvPrepareAttribs(sngAlpha2, m_hAttributes) Then
        GoTo QH
    End If
    Refresh
    '--- success
    pvAnimateState = True
QH:
    #If ImplHasTimers Then
        If sngAlpha1 > DBL_EPLISON Then
            TerminateFireOnceTimer m_uTimer
            InitFireOnceTimer m_uTimer, ObjPtr(Me), AddressOf RedirectNineButtonTimerProc, Delay:=1
            Exit Function
        End If
    #End If
    If m_hPrevBitmap <> 0 Then
        Call GdipDisposeImage(m_hPrevBitmap)
        m_hPrevBitmap = 0
    End If
    If m_hPrevAttributes <> 0 Then
        Call GdipDisposeImageAttributes(m_hPrevAttributes)
        m_hPrevAttributes = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Sub pvSetStyle(ByVal eStyle As UcsNineButtonStyleEnum)
    Const FUNC_NAME     As String = "pvSetStyle"
    Static hResBitmap   As Long
    
    On Error GoTo EH
    pvSetEmptyStyle
    If eStyle <> ucsBtyNone Then
        If hResBitmap = 0 Then
            With New cNinePatch
                If Not .frBitmapFromByteArray(FromBase64Array(STR_RES_PNG1 & STR_RES_PNG2), hResBitmap) Then
                    GoTo QH
                End If
            End With
        End If
        Select Case eStyle
        '--- buttons
        Case ucsBtyButtonDefault
            pvSetButtonStyle hResBitmap, ucsIdxButtonDefNormal, vbBlack, _
                ShadowColor:=vbWhite, ShadowOffsetY:=1
        Case ucsBtyButtonGreen
            pvSetButtonStyle hResBitmap, ucsIdxButtonGreenNormal, vbWhite
        Case ucsBtyButtonRed
            pvSetButtonStyle hResBitmap, ucsIdxButtonRedNormal, vbWhite
        Case ucsBtyButtonTurnGreen
            pvSetButtonStyle hResBitmap, ucsIdxButtonGreenNormal, vbWhite, _
                NormalTextColor:=&H7C3F&
        Case ucsBtyButtonTurnRed
            pvSetButtonStyle hResBitmap, ucsIdxButtonRedNormal, vbWhite, _
                NormalTextColor:=&H3124CB
        '--- flat buttons
        Case ucsBtyFlatPrimary
            pvSetFlatStyle hResBitmap, ucsIdxFlatPrimaryNormal, vbWhite
        Case ucsBtyFlatSecondary
            pvSetFlatStyle hResBitmap, ucsIdxFlatSecondaryNormal, &H575049, _
                ShadowOpacity:=1, ShadowColor:=vbWhite, ShadowOffsetY:=1, PressedOffset:=1
        Case ucsBtyFlatSuccess
            pvSetFlatStyle hResBitmap, ucsIdxFlatSuccessNormal, vbWhite, _
                ShadowOpacity:=0.2
        Case ucsBtyFlatDanger
            pvSetFlatStyle hResBitmap, ucsIdxFlatDangerNormal, vbWhite
        Case ucsBtyFlatWarning
            pvSetFlatStyle hResBitmap, ucsIdxFlatWarningNormal, vbWhite, _
                ShadowOpacity:=0.2
        Case ucsBtyFlatInfo
            pvSetFlatStyle hResBitmap, ucsIdxFlatInfoNormal, vbWhite, _
                ShadowOpacity:=0.2
        Case ucsBtyFlatLight
            pvSetFlatStyle hResBitmap, ucsIdxFlatLightNormal, &H575049, _
                ShadowOpacity:=1, ShadowColor:=vbWhite, ShadowOffsetY:=1
        Case ucsBtyFlatDark
            pvSetFlatStyle hResBitmap, ucsIdxFlatDarkNormal, vbWhite, _
                ShadowOpacity:=0
        '--- outline buttons
        Case ucsBtyOutlinePrimary
            pvSetOutlineStyle hResBitmap, ucsIdxFlatPrimaryOutline, &HCF7F46
        Case ucsBtyOutlineSecondary
            pvSetOutlineStyle hResBitmap, ucsIdxFlatSecondaryOutline, &H575049, _
                ShadowOpacity:=1, HoverOffset:=2
        Case ucsBtyOutlineSuccess
            pvSetOutlineStyle hResBitmap, ucsIdxFlatSuccessOutline, &HBA5E&, _
                ShadowOpacity:=0.2
        Case ucsBtyOutlineDanger
            pvSetOutlineStyle hResBitmap, ucsIdxFlatDangerOutline, &H1F20CD, _
                ShadowOpacity:=1
        Case ucsBtyOutlineWarning
            pvSetOutlineStyle hResBitmap, ucsIdxFlatWarningOutline, &HFC4F1, _
                ShadowOpacity:=0.2
        Case ucsBtyOutlineInfo
            pvSetOutlineStyle hResBitmap, ucsIdxFlatInfoOutline, &HF2AA45, _
                ShadowOpacity:=0.2
        Case ucsBtyOutlineLight
            pvSetOutlineStyle hResBitmap, ucsIdxFlatLightOutline, &H575049, _
                ShadowColor:=vbWhite, ShadowOffsetY:=1, TextColor:=DEF_TEXTCOLOR, HoverOffset:=1
        Case ucsBtyOutlineDark
            pvSetOutlineStyle hResBitmap, ucsIdxFlatDarkOutline, &H403A34, _
                ShadowOpacity:=0
        '--- cards
        Case ucsBtyCardDefault
            pvSetCardStyle hResBitmap, ucsIdxCardDefault
        Case ucsBtyCardPrimary
            pvSetCardStyle hResBitmap, ucsIdxCardPrimary
        Case ucsBtyCardSuccess
            pvSetCardStyle hResBitmap, ucsIdxCardSuccess
        Case ucsBtyCardOrange
            pvSetCardStyle hResBitmap, ucsIdxCardOrange
        Case ucsBtyCardDanger
            pvSetCardStyle hResBitmap, ucsIdxCardDanger
        Case ucsBtyCardWarning
            pvSetCardStyle hResBitmap, ucsIdxCardWarning
        Case ucsBtyCardPurple
            pvSetCardStyle hResBitmap, ucsIdxCardPurple
        End Select
    End If
QH:
'    If hResBitmap <> 0 Then
'        Call GdipDisposeImage(hResBitmap)
'        hResBitmap = 0
'    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub pvSetButtonStyle( _
            ByVal hResBitmap As Long, _
            ByVal eIdx As UcsNineButtonResIndex, _
            ByVal clrFore As OLE_COLOR, _
            Optional ByVal ShadowColor As OLE_COLOR = vbBlack, _
            Optional ByVal ShadowOffsetY As Long = -1, _
            Optional ByVal NormalTextColor As OLE_COLOR = DEF_TEXTCOLOR)
    With m_uButton(ucsBstNormal)
        If NormalTextColor = DEF_TEXTCOLOR Then
            Set .ImagePatch = pvResExtract(hResBitmap, eIdx, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
            .ShadowColor = ShadowColor
            .ShadowOffsetY = ShadowOffsetY
        Else
            Set .ImagePatch = pvResExtract(hResBitmap, ucsIdxButtonDefNormal, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
            .TextColor = NormalTextColor
            .ShadowColor = vbWhite
            .ShadowOffsetY = 1
        End If
    End With
    With m_uButton(ucsBstHover)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 1, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstPressed)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 2, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
        .TextOffsetY = 1
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstDisabled)
        Set .ImagePatch = pvResExtract(hResBitmap, ucsIdxButtonDisabled, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
        .TextOpacity = 0.4
        .TextColor = &H2E2924
        .ShadowColor = vbWhite
        .ShadowOffsetY = 1
    End With
    With m_uButton(ucsBstFocused)
        Set .ImagePatch = pvResExtract(hResBitmap, ucsIdxButtonFocus, LNG_BUTTON_TOP, LNG_BUTTON_WIDTH, LNG_BUTTON_HEIGHT)
    End With
    AnimationDuration = 0.2
    ForeColor = clrFore
End Sub

Private Sub pvSetFlatStyle( _
            ByVal hResBitmap As Long, _
            ByVal eIdx As UcsNineButtonResIndex, _
            ByVal clrFore As OLE_COLOR, _
            Optional ByVal ShadowOpacity As Single = 1, _
            Optional ByVal ShadowColor As OLE_COLOR = vbBlack, _
            Optional ByVal ShadowOffsetY As Long = -1, _
            Optional ByVal PressedOffset As Long = 0)
    With m_uButton(ucsBstNormal)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstHover)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 2, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstPressed)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 3 + PressedOffset, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .TextOffsetY = 1
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstDisabled)
        Set .ImagePatch = m_uButton(ucsBstNormal).ImagePatch
        .ImageOpacity = 0.65
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstFocused)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 4 + PressedOffset, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
    End With
    AnimationDuration = 0.2
    ForeColor = clrFore
End Sub

Private Sub pvSetOutlineStyle( _
            ByVal hResBitmap As Long, _
            ByVal eIdx As UcsNineButtonResIndex, _
            ByVal clrFore As OLE_COLOR, _
            Optional ByVal ShadowOpacity As Single = DEF_SHADOWOPACITY, _
            Optional ByVal ShadowColor As OLE_COLOR = vbBlack, _
            Optional ByVal ShadowOffsetY As Long = -1, _
            Optional ByVal TextColor As OLE_COLOR = vbWhite, _
            Optional ByVal HoverOffset As Long = -1)
    With m_uButton(ucsBstNormal)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .ShadowOpacity = 0
        .ShadowOffsetY = 1
    End With
    With m_uButton(ucsBstHover)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + HoverOffset, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .TextColor = TextColor
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstPressed)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + HoverOffset, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
        .TextColor = TextColor
        .TextOffsetY = 1
        .ShadowOpacity = ShadowOpacity
        .ShadowColor = ShadowColor
        .ShadowOffsetY = ShadowOffsetY
    End With
    With m_uButton(ucsBstDisabled)
        Set .ImagePatch = m_uButton(ucsBstNormal).ImagePatch
        .ImageOpacity = 0.65
        .ShadowOpacity = 0
        .ShadowOffsetY = 1
    End With
    With m_uButton(ucsBstFocused)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx + 3, LNG_FLAT_TOP, LNG_FLAT_WIDTH, LNG_FLAT_HEIGHT)
    End With
    AnimationDuration = 0.2
    ForeColor = clrFore
End Sub

Private Sub pvSetCardStyle( _
            ByVal hResBitmap As Long, _
            ByVal eIdx As UcsNineButtonResIndex)
    With m_uButton(ucsBstNormal)
        Set .ImagePatch = pvResExtract(hResBitmap, eIdx, LNG_CARD_TOP, LNG_CARD_WIDTH, LNG_CARD_HEIGHT)
    End With
    With m_uButton(ucsBstDisabled)
        Set .ImagePatch = m_uButton(ucsBstNormal).ImagePatch
        .ImageOpacity = 0.65
    End With
    With m_uButton(ucsBstFocused)
        Set .ImagePatch = pvResExtract(hResBitmap, ucsIdxCardFocus, LNG_CARD_TOP, LNG_CARD_WIDTH, LNG_CARD_HEIGHT)
    End With
End Sub

Private Function pvResExtract(ByVal hResBitmap As Long, ByVal eIdx As UcsNineButtonResIndex, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long) As cNinePatch
    Const FUNC_NAME     As String = "pvResExtract"
    Dim hNewBitmap      As Long
    Dim lLeft           As Long
    Dim oRetVal         As cNinePatch
    
    On Error GoTo EH
    lLeft = eIdx * lWidth
    lTop = lTop + lHeight * (lLeft \ LNG_RES_WIDTH)
    lLeft = lLeft Mod LNG_RES_WIDTH
    If GdipCloneBitmapAreaI(lLeft, lTop, lWidth, lHeight, 0, hResBitmap, hNewBitmap) <> 0 Then
        GoTo QH
    End If
    Set oRetVal = New cNinePatch
    If Not oRetVal.LoadFromBitmap(hNewBitmap) Then
        GoTo QH
    End If
    hNewBitmap = 0
    '--- success
    Set pvResExtract = oRetVal
QH:
    If hNewBitmap <> 0 Then
        Call GdipDisposeImage(hNewBitmap)
        hNewBitmap = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Sub pvSetEmptyStyle()
    Dim lIdx            As Long
    Dim uEmpty          As UcsNinePatchType

    With uEmpty
        .ImageOpacity = DEF_IMAGEOPACITY
        .TextOpacity = DEF_TEXTOPACITY
        .TextColor = DEF_TEXTCOLOR
        .TextFlags = DEF_TEXTFLAGS
        .ShadowOpacity = DEF_SHADOWOPACITY
        .ShadowColor = DEF_SHADOWCOLOR
    End With
    For lIdx = 0 To UBound(m_uButton)
        m_uButton(lIdx) = uEmpty
    Next
End Sub

#If Not ImplUseShared Then
Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Private Property Get DateTimer() As Double
    Dim cDateTime       As Currency
    
    Call GetSystemTimeAsFileTime(cDateTime)
    DateTimer = CDbl(cDateTime - 9435304800000@) / 1000#
End Property

Private Function RedrawWindow(ByVal hWnd As Long, Optional ByVal UpdateImmediate As Boolean) As Long
    If hWnd <> 0 Then
        RedrawWindow = ApiRedrawWindow(hWnd, 0, 0, RDW_INVALIDATE Or RDW_ALLCHILDREN Or (-UpdateImmediate * (RDW_ERASE Or RDW_FRAME Or RDW_UPDATENOW)))
    End If
End Function

Private Function FromBase64Array(sText As String) As Byte()
    Dim lSize           As Long
    Dim dwDummy         As Long
    Dim baOutput()      As Byte
    
    lSize = Len(sText) + 1
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, dwDummy)
    If lSize > 0 Then
        ReDim Preserve baOutput(0 To lSize - 1) As Byte
        FromBase64Array = baOutput
    Else
        FromBase64Array = vbNullString
    End If
End Function
#End If

'=========================================================================
' Base class events
'=========================================================================

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    pvPrepareFont m_oFont, m_hFont
    Repaint
    PropertyChanged
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    If Ambient.UserMode Then
        HitResult = pvHitTest(X, Y)
    Else
        HitResult = vbHitResultHit
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "UserControl_MouseDown"
    
    On Error GoTo EH
    m_nDownButton = Button
    m_nDownShift = Shift
    m_sngDownX = X
    m_sngDownY = Y
    If (Button And vbLeftButton) <> 0 Then
        If pvHitTest(X, Y) <> vbHitResultOutside Then
            pvRegisterCancelMode Me
            pvState(ucsBstPressed Or ucsBstFocused) = True
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "UserControl_MouseMove"
    
    On Error GoTo EH
    If X >= 0 And X < ScaleWidth And Y >= 0 And Y < ScaleHeight Then
        If Not pvState(ucsBstHover) Then
            If pvRegisterCancelMode(Me) Then
                pvState(ucsBstHover) = True
            End If
        End If
    Else
        If pvState(ucsBstHover) Then
            pvState(ucsBstHover) = False
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbLeftButton) <> 0 Then
        pvState(ucsBstPressed) = False
    End If
    If X >= 0 And X < ScaleWidth And Y >= 0 And Y < ScaleHeight Then
        Call RedrawWindow(ContainerHwnd, True)
        If (m_nDownButton And Button And vbLeftButton) <> 0 Then
            RaiseEvent Click
        ElseIf (m_nDownButton And Button And vbRightButton) <> 0 Then
            RaiseEvent ContextMenu
        End If
    End If
    m_nDownButton = 0
End Sub

Private Sub UserControl_DblClick()
    UserControl_MouseDown vbLeftButton, m_nDownShift, m_sngDownX, m_sngDownY
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Paint()
    Const FUNC_NAME     As String = "UserControl_Paint"
    Dim hGraphics       As Long
    
    On Error GoTo EH
    If Not m_bShown Then
        m_bShown = True
        pvPrepareBitmap m_eState, m_hFocusBitmap, m_hBitmap
    End If
    If m_hBitmap <> 0 Then
        If GdipCreateFromHDC(hDC, hGraphics) <> 0 Then
            GoTo QH
        End If
        If m_hFocusBitmap <> 0 Then
            If GdipDrawImageRectRect(hGraphics, m_hFocusBitmap, 0, 0, ScaleWidth, ScaleHeight, 0, 0, ScaleWidth, ScaleHeight, , m_hFocusAttributes) <> 0 Then
                GoTo QH
            End If
        End If
        If m_hPrevBitmap <> 0 Then
            If GdipDrawImageRectRect(hGraphics, m_hPrevBitmap, 0, 0, ScaleWidth, ScaleHeight, 0, 0, ScaleWidth, ScaleHeight, , m_hPrevAttributes) <> 0 Then
                GoTo QH
            End If
        End If
        If GdipDrawImageRectRect(hGraphics, m_hBitmap, 0, 0, ScaleWidth, ScaleHeight, 0, 0, ScaleWidth, ScaleHeight, , m_hAttributes) <> 0 Then
            GoTo QH
        End If
    Else
        Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &HE0FFFF, BF
    End If
QH:
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub UserControl_EnterFocus()
    pvState(ucsBstFocused) = True
End Sub

Private Sub UserControl_ExitFocus()
    pvState(ucsBstFocused) = False
End Sub

Private Sub UserControl_InitProperties()
    Const FUNC_NAME     As String = "UserControl_InitProperties"
    
    On Error GoTo EH
    Style = DEF_STYLE
    Enabled = DEF_ENABLED
    Opacity = DEF_OPACITY
    AnimationDuration = DEF_ANIMATIONDURATION
    Caption = Ambient.DisplayName
    Set Font = Ambient.Font
    ForeColor = DEF_FORECOLOR
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    With PropBag
        Style = .ReadProperty("Style", DEF_STYLE)
        Enabled = .ReadProperty("Enabled", DEF_ENABLED)
        Opacity = .ReadProperty("Opacity", DEF_OPACITY)
        AnimationDuration = .ReadProperty("AnimationDuration", DEF_ANIMATIONDURATION)
        Caption = .ReadProperty("Caption", Ambient.DisplayName)
        Set Font = .ReadProperty("Font", Ambient.Font)
        ForeColor = .ReadProperty("ForeColor", DEF_FORECOLOR)
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_WriteProperties"
    
    On Error GoTo EH
    With PropBag
        .WriteProperty "Style", Style, DEF_STYLE
        .WriteProperty "Enabled", Enabled, DEF_ENABLED
        .WriteProperty "Opacity", Opacity, DEF_OPACITY
        .WriteProperty "AnimationDuration", AnimationDuration, DEF_ANIMATIONDURATION
        .WriteProperty "Caption", Caption, Ambient.DisplayName
        .WriteProperty "Font", Font, Ambient.Font
        .WriteProperty "ForeColor", ForeColor, DEF_FORECOLOR
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_Resize()
    Repaint
End Sub

Private Sub UserControl_Hide()
    m_bShown = False
End Sub

Private Sub UserControl_Initialize()
    Dim uStartup        As GdiplusStartupInput
    
    If GetModuleHandle("gdiplus") = 0 Then
        uStartup.GdiplusVersion = 1&
        Call GdiplusStartup(0, uStartup)
    End If
    pvSetEmptyStyle
End Sub

Private Sub UserControl_Terminate()
    If m_hAttributes <> 0 Then
        Call GdipDisposeImageAttributes(m_hAttributes)
        m_hAttributes = 0
    End If
    If m_hBitmap <> 0 Then
        Call GdipDisposeImage(m_hBitmap)
        m_hBitmap = 0
    End If
    If m_hPrevBitmap <> 0 Then
        Call GdipDisposeImage(m_hPrevBitmap)
        m_hPrevBitmap = 0
    End If
    If m_hPrevAttributes <> 0 Then
        Call GdipDisposeImageAttributes(m_hPrevAttributes)
        m_hPrevAttributes = 0
    End If
    If m_hFocusBitmap <> 0 Then
        Call GdipDisposeImage(m_hFocusBitmap)
        m_hFocusBitmap = 0
    End If
    If m_hFocusAttributes <> 0 Then
        Call GdipDisposeImageAttributes(m_hFocusAttributes)
        m_hFocusAttributes = 0
    End If
    If m_hFont <> 0 Then
        Call GdipDeleteFont(m_hFont)
        m_hFont = 0
    End If
    #If ImplHasTimers Then
        TerminateFireOnceTimer m_uTimer
    #End If
End Sub

