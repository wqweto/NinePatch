VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9936
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17172
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9936
   ScaleWidth      =   17172
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   468
      Left            =   1092
      TabIndex        =   35
      Top             =   1428
      Width           =   348
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      Height          =   468
      Left            =   672
      TabIndex        =   34
      Top             =   1428
      Width           =   348
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Enabled"
      Height          =   216
      Left            =   4536
      TabIndex        =   33
      Top             =   1596
      Value           =   1  'Checked
      Width           =   1608
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Enabled"
      Height          =   216
      Left            =   4536
      TabIndex        =   18
      Top             =   1260
      Value           =   1  'Checked
      Width           =   1608
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enabled"
      Height          =   216
      Left            =   4536
      TabIndex        =   17
      Top             =   924
      Value           =   1  'Checked
      Width           =   1608
   End
   Begin VB.TextBox Text1 
      Height          =   960
      Left            =   420
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   5712
      Width           =   3624
   End
   Begin VB.ComboBox Combo1 
      Height          =   312
      Left            =   420
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   5292
      Width           =   3624
   End
   Begin VB.PictureBox picTab1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FBF7F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasDC           =   0   'False
      Height          =   1608
      Left            =   420
      ScaleHeight     =   1608
      ScaleWidth      =   3456
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6804
      Width           =   3456
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   216
         Left            =   1092
         TabIndex        =   2
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   1332
         WordWrap        =   -1  'True
      End
   End
   Begin Project1.ctxNineButton ctxNineButton28 
      Height          =   1524
      Left            =   11508
      TabIndex        =   36
      Top             =   3192
      Width           =   1776
      _ExtentX        =   3133
      _ExtentY        =   2688
      Style           =   28
      Caption         =   "ctxNineButton28"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ctxNineButton ctxNineButton27 
      Height          =   600
      Left            =   14448
      TabIndex        =   32
      Top             =   2352
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   21
      AnimationDuration=   0.2
      Caption         =   "Dark Outline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4209204
   End
   Begin Project1.ctxNineButton ctxNineButton26 
      Height          =   600
      Left            =   14448
      TabIndex        =   31
      Top             =   1680
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   13
      AnimationDuration=   0.2
      Caption         =   "Dark"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin Project1.ctxNineButton ctxNineButton25 
      Height          =   600
      Left            =   12432
      TabIndex        =   30
      Top             =   2352
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   20
      AnimationDuration=   0.2
      Caption         =   "Light Outline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   5722185
   End
   Begin Project1.ctxNineButton ctxNineButton24 
      Height          =   600
      Left            =   12432
      TabIndex        =   29
      Top             =   1680
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   12
      AnimationDuration=   0.2
      Caption         =   "Light"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   5722185
   End
   Begin Project1.ctxNineButton ctxNineButton23 
      Height          =   600
      Left            =   10416
      TabIndex        =   28
      Top             =   2352
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   19
      AnimationDuration=   0.2
      Caption         =   "Info Outline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   15903301
   End
   Begin Project1.ctxNineButton ctxNineButton22 
      Height          =   600
      Left            =   10416
      TabIndex        =   27
      Top             =   1680
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   11
      AnimationDuration=   0.2
      Caption         =   "Info"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin Project1.ctxNineButton ctxNineButton21 
      Height          =   600
      Left            =   8400
      TabIndex        =   26
      Top             =   2352
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   18
      AnimationDuration=   0.2
      Caption         =   "Warning Outline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   1033457
   End
   Begin Project1.ctxNineButton ctxNineButton20 
      Height          =   600
      Left            =   8400
      TabIndex        =   25
      Top             =   1680
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   10
      AnimationDuration=   0.2
      Caption         =   "Warning"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin Project1.ctxNineButton ctxNineButton19 
      Height          =   600
      Left            =   14448
      TabIndex        =   24
      Top             =   840
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   17
      AnimationDuration=   0.2
      Caption         =   "Danger Outline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   2040013
   End
   Begin Project1.ctxNineButton ctxNineButton18 
      Height          =   600
      Left            =   14448
      TabIndex        =   23
      Top             =   168
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   9
      AnimationDuration=   0.2
      Caption         =   "Danger"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin Project1.ctxNineButton ctxNineButton17 
      Height          =   600
      Left            =   12432
      TabIndex        =   22
      Top             =   840
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   16
      AnimationDuration=   0.2
      Caption         =   "Success Outline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   47710
   End
   Begin Project1.ctxNineButton ctxNineButton16 
      Height          =   600
      Left            =   12432
      TabIndex        =   21
      Top             =   168
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   8
      AnimationDuration=   0.2
      Caption         =   "Success"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin Project1.ctxNineButton ctxNineButton14 
      Height          =   600
      Left            =   10416
      TabIndex        =   20
      Top             =   168
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   7
      AnimationDuration=   0.2
      Caption         =   "Secondary"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   5722185
   End
   Begin Project1.ctxNineButton ctxNineButton15 
      Height          =   600
      Left            =   10416
      TabIndex        =   19
      Top             =   840
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   15
      AnimationDuration=   0.2
      Caption         =   "Secondary Outline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   5722185
   End
   Begin Project1.ctxNineButton ctxNineButton11 
      Height          =   1524
      Left            =   5964
      TabIndex        =   16
      Top             =   3192
      Width           =   1776
      _ExtentX        =   3133
      _ExtentY        =   2688
      Style           =   25
      Caption         =   "ctxNineButton11"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ctxNineButton ctxNineButton12 
      Height          =   1524
      Left            =   7812
      TabIndex        =   15
      Top             =   3192
      Width           =   1776
      _ExtentX        =   3133
      _ExtentY        =   2688
      Style           =   26
      Caption         =   "ctxNineButton12"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ctxNineButton ctxNineButton13 
      Height          =   1524
      Left            =   9660
      TabIndex        =   14
      Top             =   3192
      Width           =   1776
      _ExtentX        =   3133
      _ExtentY        =   2688
      Style           =   27
      Caption         =   "ctxNineButton13"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ctxNineButton ctxNineButton10 
      Height          =   1524
      Left            =   4116
      TabIndex        =   13
      Top             =   3192
      Width           =   1776
      _ExtentX        =   3133
      _ExtentY        =   2688
      Style           =   24
      Caption         =   "ctxNineButton10"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ctxNineButton ctxNineButton9 
      Height          =   1524
      Left            =   2268
      TabIndex        =   12
      Top             =   3192
      Width           =   1776
      _ExtentX        =   3133
      _ExtentY        =   2688
      Style           =   23
      Caption         =   "ctxNineButton9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ctxNineButton ctxNineButton8 
      Height          =   1524
      Left            =   420
      TabIndex        =   11
      Top             =   3192
      Width           =   1776
      _ExtentX        =   3133
      _ExtentY        =   2688
      Style           =   22
      Caption         =   "ctxNineButton8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ctxNineButton ctxNineButton7 
      Height          =   600
      Left            =   8400
      TabIndex        =   10
      Top             =   840
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   14
      AnimationDuration=   0.2
      Caption         =   "Source code"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   13598534
   End
   Begin Project1.ctxNineButton ctxNineButton6 
      Height          =   600
      Left            =   8400
      TabIndex        =   9
      Top             =   168
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   6
      AnimationDuration=   0.2
      Caption         =   "Send data"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin Project1.ctxNineButton ctxNineButton5 
      Height          =   600
      Left            =   6384
      TabIndex        =   8
      Top             =   840
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   5
      AnimationDuration=   0.2
      Caption         =   "Revoke all"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin Project1.ctxNineButton ctxNineButton4 
      Height          =   600
      Left            =   6384
      TabIndex        =   7
      Top             =   168
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      Style           =   2
      AnimationDuration=   0.2
      Caption         =   "Update profile"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin Project1.ctxNineButton ctxNineButton3 
      Height          =   600
      Left            =   4368
      TabIndex        =   6
      Top             =   168
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1058
      AnimationDuration=   0.2
      Caption         =   "ctxNineButton3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin Project1.ctxNineButton ctxNineButton2 
      Height          =   936
      Left            =   2352
      TabIndex        =   4
      Top             =   168
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1651
      Style           =   0
      AnimationDuration=   0.2
      Caption         =   "ctxNineButton2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin Project1.ctxNineButton ctxNineButton1 
      Height          =   936
      Left            =   336
      TabIndex        =   5
      Top             =   168
      Width           =   1944
      _ExtentX        =   3429
      _ExtentY        =   1651
      Style           =   16
      AnimationDuration=   0.2
      Caption         =   "ctxNineButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4209204
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--- for GdipCreateFont
Private Const UnitPoint                     As Long = 3

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal hFontCollection As Long, hFontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "gdiplus" (hFontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal hFontFamily As Long) As Long
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal hFontFamily As Long, ByVal emSize As Single, ByVal Style As Long, ByVal unit As Long, createdfont As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, hBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal hGraphics As Long, ByVal str As Long, ByVal Length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal hStringFormat As Long, ByVal hBrush As Long) As Long

Private Type RECTF
   Left             As Single
   Top              As Single
   Right            As Single
   Bottom           As Single
End Type

Private Const DEF_OFFSETX           As Long = 10
Private Const DEF_OFFSETY           As Long = 10

Private m_oPatch                As cNinePatch
Private m_oCtlCancelMode        As Object
Private m_hFont                 As Long

Private Enum FontStyle
   FontStyleRegular = 0
   FontStyleBold = 1
   FontStyleItalic = 2
   FontStyleBoldItalic = 3
   FontStyleUnderline = 4
   FontStyleStrikeout = 8
End Enum

Private Sub Check1_Click()
    ctxNineButton1.Enabled = Check1.Value = vbChecked
    ctxNineButton2.Enabled = Check1.Value = vbChecked
    ctxNineButton3.Enabled = Check1.Value = vbChecked
End Sub

Private Sub Check2_Click()
    ctxNineButton4.Enabled = Check2.Value = vbChecked
    ctxNineButton5.Enabled = Check2.Value = vbChecked
End Sub

Private Sub Check3_Click()
    ctxNineButton6.Enabled = Check3.Value = vbChecked
    ctxNineButton7.Enabled = Check3.Value = vbChecked
    ctxNineButton8.Enabled = Check3.Value = vbChecked
    ctxNineButton9.Enabled = Check3.Value = vbChecked
    ctxNineButton10.Enabled = Check3.Value = vbChecked
    ctxNineButton11.Enabled = Check3.Value = vbChecked
    ctxNineButton12.Enabled = Check3.Value = vbChecked
    ctxNineButton13.Enabled = Check3.Value = vbChecked
    ctxNineButton14.Enabled = Check3.Value = vbChecked
    ctxNineButton15.Enabled = Check3.Value = vbChecked
    ctxNineButton16.Enabled = Check3.Value = vbChecked
    ctxNineButton17.Enabled = Check3.Value = vbChecked
    ctxNineButton18.Enabled = Check3.Value = vbChecked
    ctxNineButton19.Enabled = Check3.Value = vbChecked
    ctxNineButton20.Enabled = Check3.Value = vbChecked
    ctxNineButton21.Enabled = Check3.Value = vbChecked
    ctxNineButton22.Enabled = Check3.Value = vbChecked
    ctxNineButton23.Enabled = Check3.Value = vbChecked
    ctxNineButton24.Enabled = Check3.Value = vbChecked
    ctxNineButton25.Enabled = Check3.Value = vbChecked
    ctxNineButton26.Enabled = Check3.Value = vbChecked
    ctxNineButton27.Enabled = Check3.Value = vbChecked
End Sub

Private Sub Command1_Click()
    ctxNineButton1.Style = ctxNineButton1.Style - 1
End Sub

Private Sub Command2_Click()
    ctxNineButton1.Style = ctxNineButton1.Style + 1
End Sub

Private Sub ctxNineButton1_Click()
    Screen.MousePointer = vbHourglass
    pvLongRunningTask
    MsgBox "Done", vbExclamation
    Screen.MousePointer = vbDefault
End Sub

Private Sub pvLongRunningTask()
    Dim dblTimer        As Double

    dblTimer = Timer
    Do While Timer < dblTimer + 1
        Call Sleep(1)
    Loop
End Sub

Private Sub ctxNineButton8_OwnerDraw(ByVal hGraphics As Long, ByVal hFont As Long, ByVal ButtonState As UcsNineButtonStateEnum, ClientLeft As Long, ClientTop As Long, ClientWidth As Long, ClientHeight As Long, Caption As String, ByVal hPicture As Long)
    Dim hBrush      As Long
    Dim lOffset     As Long
    Dim uRect       As RECTF

    If m_hFont = 0 Then
        m_hFont = pvGetFont(ctxNineButton8.Font)
    End If
    If GdipCreateSolidFill(&HEE0000FF, hBrush) <> 0 Then
        GoTo QH
    End If
    lOffset = -((ButtonState And ucsBstHoverPressed) = ucsBstHoverPressed)
    uRect.Left = ClientLeft ' + lOffset
    uRect.Top = ClientTop + lOffset
    uRect.Right = uRect.Left + ClientWidth
    uRect.Bottom = uRect.Top + ClientHeight
    If GdipDrawString(hGraphics, StrPtr("Total: $0.72"), -1, m_hFont, uRect, 0, hBrush) <> 0 Then
        GoTo QH
    End If
QH:
    Call GdipDeleteBrush(hBrush)
End Sub

Private Function pvGetFont(oFont As StdFont) As Long
    Dim hFamily     As Long
    Dim hFont       As Long

    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFamily) <> 0 Then
        If GdipGetGenericFontFamilySansSerif(hFamily) <> 0 Then
            GoTo QH
        End If
    End If
    If GdipCreateFont(hFamily, oFont.Size, FontStyleRegular, UnitPoint, hFont) <> 0 Then
        GoTo QH
    End If
    '--- success
    pvGetFont = hFont
QH:
    If hFamily <> 0 Then
        Call GdipDeleteFontFamily(hFamily)
        hFamily = 0
    End If
End Function


Private Sub Form_Click()
    On Error Resume Next
    BackColor = QBColor(Rnd * 16)
End Sub

Private Sub Form_DblClick()
    Text1.Text = ToBase64Array(ReadBinaryFile(App.Path & "\res\" & Combo1.Text))
End Sub

Private Sub Form_Load()
    Dim sFile           As String
    Dim oBold           As StdFont
    
    sFile = Dir$(App.Path & "\res\*.png")
    Do While LenB(sFile) <> 0
        Combo1.AddItem sFile
        sFile = Dir$
    Loop
    Combo1.ListIndex = 0
    Text1.Text = "Lorem ipsum dolor sit amet, nunc lorem viverra morbi, diam leo curabitur eu libero odio, orci dapibus donec, donec dui convallis dolor metus ac in."
''    ctxNineButton1.ButtonImageArray(ucsBstNormal) = ReadBinaryFile(App.Path & "\res\a9p_09_11_00002.9.png")
''    ctxNineButton1.ButtonImageArray(ucsBstHover) = ReadBinaryFile(App.Path & "\res\a9p_09_11_00007.9.png")
''    ctxNineButton1.ButtonImageArray(ucsBstPressed) = ReadBinaryFile(App.Path & "\res\a9p_09_11_00009.9.png")
''    ctxNineButton1.ButtonImageArray(ucsBstPressed Or ucsBstHover) = ReadBinaryFile(App.Path & "\res\a9p_09_11_00010.9.png")
'    With ctxNineButton1
''        .ButtonImageArray(ucsBstNormal) = ReadBinaryFile(App.Path & "\res\toast_c.9.png")
''        .ButtonImageArray(ucsBstHover) = ReadBinaryFile(App.Path & "\res\toast_d.9.png")
''        .ButtonImageArray(ucsBstFocused) = ReadBinaryFile(App.Path & "\res\toast_e.9.png")
'        .AnimationDuration = 0.2
'    End With
    With ctxNineButton2
        .ButtonImageArray(ucsBstNormal) = ReadBinaryFile(App.Path & "\res\a9p_09_11_00086.9.png")
        .ButtonImageArray(ucsBstHover) = ReadBinaryFile(App.Path & "\res\a9p_09_11_00087.9.png")
        .ButtonImageArray(ucsBstPressed) = ReadBinaryFile(App.Path & "\res\a9p_09_11_00088.9.png")
        .AnimationDuration = 0.5
        .Opacity = 0.8
    End With
    With ctxNineButton3
        Set oBold = CloneFont(.Font)
        oBold.Bold = True
        Set .ButtonTextFont(ucsBstHover) = oBold
        Set .ButtonTextFont(ucsBstPressed) = oBold
    End With
'    With ctxNineButton3
'        Set oBold = CloneFont(.Font)
'        oBold.Bold = True
'
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-def-normal.png")
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-def-hover.png")
'        Set .ButtonTextFont = oBold
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-def-pressed.png")
'        Set .ButtonTextFont = oBold
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-def-disabled.png")
'        .ButtonTextOpacity = 0.4
'        .ButtonTextColor = &H2E2924
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-def-focus.png")
'
'        .AnimationDuration = 0.2
'        .ButtonState = ucsBstNormal
''        .Opacity = 0.8
'    End With
'
'    With ctxNineButton4
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-green-normal.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-green-hover.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-green-pressed.png")
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-def-disabled.png")
'        .ButtonTextOpacity = 0.4
'        .ButtonTextColor = &H2E2924
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-def-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = vbWhite
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton5
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-def-normal.png")
'        .ButtonTextColor = &H3124CB
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-red-hover.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-red-pressed.png")
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-def-disabled.png")
'        .ButtonTextOpacity = 0.4
'        .ButtonTextColor = &H2E2924
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-def-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = vbWhite
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton6
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-primary-normal.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-primary-hover.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-primary-pressed.png")
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-primary-normal.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-primary-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = vbWhite
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton7
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-primary-outline.png")
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-primary-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-primary-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-primary-outline.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-primary-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = &HCF7F46
'        .ButtonState = ucsBstNormal
'    End With
'
    With ctxNineButton8
        .ButtonTextOpacity(ucsBstNormal) = 1
    End With
'    With ctxNineButton8
'        .ButtonImageArray(ucsBstNormal) = ReadBinaryFile(App.Path & "\res\card-normal.png")
'        .ButtonImageArray(ucsBstHover) = EmptyByteArray
'        .ButtonImageArray(ucsBstPressed) = EmptyByteArray
'        .ButtonImageArray(ucsBstFocused) = ReadBinaryFile(App.Path & "\res\card-focus.png")
'    End With
'    With ctxNineButton9
'        .ButtonImageArray(ucsBstNormal) = ReadBinaryFile(App.Path & "\res\card-blue.png")
'        .ButtonImageArray(ucsBstHover) = EmptyByteArray
'        .ButtonImageArray(ucsBstPressed) = EmptyByteArray
'        .ButtonImageArray(ucsBstFocused) = ReadBinaryFile(App.Path & "\res\card-focus.png")
'    End With
'    With ctxNineButton10
'        .ButtonImageArray(ucsBstNormal) = ReadBinaryFile(App.Path & "\res\card-green.png")
'        .ButtonImageArray(ucsBstHover) = EmptyByteArray
'        .ButtonImageArray(ucsBstPressed) = EmptyByteArray
'        .ButtonImageArray(ucsBstFocused) = ReadBinaryFile(App.Path & "\res\card-focus.png")
'    End With
'    With ctxNineButton11
'        .ButtonImageArray(ucsBstNormal) = ReadBinaryFile(App.Path & "\res\card-orange.png")
'        .ButtonImageArray(ucsBstHover) = EmptyByteArray
'        .ButtonImageArray(ucsBstPressed) = EmptyByteArray
'        .ButtonImageArray(ucsBstFocused) = ReadBinaryFile(App.Path & "\res\card-focus.png")
'    End With
'    With ctxNineButton12
'        .ButtonImageArray(ucsBstNormal) = ReadBinaryFile(App.Path & "\res\card-red.png")
'        .ButtonImageArray(ucsBstHover) = EmptyByteArray
'        .ButtonImageArray(ucsBstPressed) = EmptyByteArray
'        .ButtonImageArray(ucsBstFocused) = ReadBinaryFile(App.Path & "\res\card-focus.png")
'    End With
'    With ctxNineButton13
'        .ButtonImageArray(ucsBstNormal) = ReadBinaryFile(App.Path & "\res\card-purple.png")
'        .ButtonImageArray(ucsBstHover) = EmptyByteArray
'        .ButtonImageArray(ucsBstPressed) = EmptyByteArray
'        .ButtonImageArray(ucsBstFocused) = ReadBinaryFile(App.Path & "\res\card-focus.png")
'    End With
'
'    With ctxNineButton14
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-secondary-normal.png")
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-secondary-hover.png")
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-secondary-pressed.png")
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-secondary-normal.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-secondary-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = &H575049
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton15
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-secondary-outline.png")
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-secondary-outline-hover.png")
'        .ButtonTextColor = vbWhite
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-secondary-outline-hover.png")
'        .ButtonTextColor = vbWhite
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-secondary-outline.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-secondary-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = &H575049
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton16
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-success-normal.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-success-hover.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-success-pressed.png")
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-success-normal.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-success-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = vbWhite
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton17
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-success-outline.png")
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-success-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-success-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-success-outline.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-success-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = &HBA5E&
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton18
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-danger-normal.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-danger-hover.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-danger-pressed.png")
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-danger-normal.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-danger-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = vbWhite
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton19
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-danger-outline.png")
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-danger-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-danger-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-danger-outline.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-danger-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = &H1F20CD
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton20
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-warning-normal.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-warning-hover.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-warning-pressed.png")
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-warning-normal.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-warning-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = vbWhite
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton21
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-warning-outline.png")
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-warning-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-warning-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-warning-outline.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-warning-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = &HFC4F1
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton22
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-info-normal.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-info-hover.png")
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-info-pressed.png")
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-info-normal.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-info-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = vbWhite
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton23
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-info-outline.png")
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-info-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-info-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbBlack
'        .ButtonShadowOpacity = 0.2
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-info-outline.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-info-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = &HF2AA45
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton24
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-light-normal.png")
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-light-hover.png")
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-light-pressed.png")
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-light-normal.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-light-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = &H575049
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton25
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-light-outline.png")
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-light-hover.png")
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-light-hover.png")
'        .ButtonTextOffsetY = 1
'        .ButtonShadowColor = vbWhite
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-light-outline.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-light-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = &H575049
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton26
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-dark-normal.png")
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-dark-hover.png")
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-dark-pressed.png")
'        .ButtonTextOffsetY = 1
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-dark-normal.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-dark-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = vbWhite
'        .ButtonState = ucsBstNormal
'    End With
'
'    With ctxNineButton27
'        .ButtonState = ucsBstNormal
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-dark-outline.png")
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstHover
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-dark-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstPressed
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-dark-normal.png")
'        .ButtonTextColor = vbWhite
'        .ButtonTextOffsetY = 1
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = -1
'
'        .ButtonState = ucsBstDisabled
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-dark-outline.png")
'        .ButtonImageOpacity = 0.65
'        .ButtonShadowOpacity = 0
'        .ButtonShadowOffsetY = 1
'
'        .ButtonState = ucsBstFocused
'        .ButtonImageArray = ReadBinaryFile(App.Path & "\res\button-flat-dark-focus.png")
'
'        .AnimationDuration = 0.2
'        .ForeColor = &H403A34
'        .ButtonState = ucsBstNormal
'    End With
End Sub

Property Get EmptyByteArray() As Byte()
    EmptyByteArray = vbNullString
End Property

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_oCtlCancelMode Is Nothing Then
        m_oCtlCancelMode.CancelMode
        Set m_oCtlCancelMode = Nothing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    ShutdownGdip
End Sub

Private Sub Combo1_Click()
    Set m_oPatch = New cNinePatch
'    m_oPatch.LoadFromByteArray ReadBinaryFile(App.Path & "\res\" & Combo1.Text)
    m_oPatch.LoadFromFile App.Path & "\res\" & Combo1.Text
    Text1_Change
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = vbMinimized Then
        Exit Sub
    End If
    Combo1.Width = ScaleWidth - 2 * Combo1.Left
    Text1.Width = ScaleWidth - 2 * Text1.Left
    picTab1.Move picTab1.Left, picTab1.Top, ScaleWidth - 2 * picTab1.Left, ScaleHeight - picTab1.Top - picTab1.Left
    If Not m_oPatch Is Nothing Then
        picTab1.Cls
        m_oPatch.DrawToDC picTab1.hDC, DEF_OFFSETX, DEF_OFFSETY, _
            picTab1.ScaleWidth / Screen.TwipsPerPixelX - 2 * DEF_OFFSETX, _
            picTab1.ScaleHeight / Screen.TwipsPerPixelY - 2 * DEF_OFFSETY
    End If
End Sub

Public Sub RegisterCancelMode(oCtl As Object)
    If Not m_oCtlCancelMode Is Nothing And Not m_oCtlCancelMode Is oCtl Then
        m_oCtlCancelMode.CancelMode
    End If
    Set m_oCtlCancelMode = oCtl
End Sub

Private Sub Text1_Change()
    Dim lBoxWidth As Long
    Dim lBoxHeight As Long
    Dim lClientX As Long
    Dim lClientY As Long

    Label1.Caption = Text1.Text
    m_oPatch.CalcBoundingBox Label1.Width / Screen.TwipsPerPixelX, Label1.Height / Screen.TwipsPerPixelY, lBoxWidth, lBoxHeight, lClientX, lClientY
    Label1.Move (DEF_OFFSETX + lClientX) * Screen.TwipsPerPixelX, (DEF_OFFSETY + lClientY) * Screen.TwipsPerPixelY
    picTab1.Move picTab1.Left, picTab1.Top, _
        (DEF_OFFSETX + lBoxWidth + DEF_OFFSETX) * Screen.TwipsPerPixelX, _
        (DEF_OFFSETY + lBoxHeight + DEF_OFFSETY) * Screen.TwipsPerPixelX
    picTab1.Cls
    m_oPatch.DrawToDC picTab1.hDC, DEF_OFFSETX, DEF_OFFSETY, _
        picTab1.ScaleWidth / Screen.TwipsPerPixelX - 2 * DEF_OFFSETX, _
        picTab1.ScaleHeight / Screen.TwipsPerPixelY - 2 * DEF_OFFSETY
End Sub

Private Function CloneFont(pFont As IFont) As StdFont
    If Not pFont Is Nothing Then
        pFont.Clone CloneFont
    Else
        Set CloneFont = New StdFont
    End If
End Function
