VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   10704
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10752
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   10704
   ScaleWidth      =   10752
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   756
      TabIndex        =   0
      Top             =   756
      Width           =   5724
   End
   Begin Project1.ctxTouchKeyboard ctxTouchKeyboard2 
      Height          =   4044
      Left            =   1176
      Top             =   5964
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   7133
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "PT Sans Narrow"
         Size            =   13.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layout          =   "1 2 3 4|||N 5 6 7|||N 8 9 .|||N 0 <=||D Cancel||D|N Done||B"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   192
      Left            =   6972
      TabIndex        =   1
      Top             =   588
      UseMnemonic     =   0   'False
      Width           =   504
      WordWrap        =   -1  'True
   End
   Begin Project1.ctxTouchKeyboard ctxTouchKeyboard1 
      Height          =   3288
      Left            =   0
      Top             =   2016
      Width           =   9924
      _ExtentX        =   17505
      _ExtentY        =   5800
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "PT Sans Narrow"
         Size            =   13.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STR_EN_LAYOUT1         As String = "q w e r t y u i o p <=||D " & _
                                                "|0.5|S|N a s d f g h j k l Done|1.75|B " & _
                                                "^^|||N z x c v b n m ! ? ^^|1.25| " & _
                                                "?!123|1.5|D|N BG|1.5|D space|6 ?!123|1.25|D keyb||D"
Private Const STR_EN_LAYOUT2         As String = "Q W E R T Y U I O P <=||D " & _
                                                "|0.5|S|N A S D F G H J K L Done|1.75|B " & _
                                                "^^||L|N Z X C V B N M ! ? ^^|1.25|L " & _
                                                "?!123|1.5|D|N BG|1.5|D space|6 ?!123|1.25|D keyb||D"
Private Const STR_EN_LAYOUT3         As String = "1 2 3 4 5 6 7 8 9 0 <=||D " & _
                                                "|0.5|S|N - / : ; ( ) € & @ Done|1.75|B " & _
                                                "#+=|||N . , ? ! ' "" #+=|1.25 " & _
                                                "ABC|1.5|D|N BG|1.5|D space|6 ABC|1.25|D keyb||D"
Private Const STR_EN_LAYOUT4         As String = "[ ] { } # % ^ * + = <=||D " & _
                                                "|0.5|S|N _ \ | ~ < > $ & · Done|1.75|B " & _
                                                "#+=||L|N . , ? ! ' "" #+=|1.25|L " & _
                                                "ABC|1.5|D|N BG|1.5|D space|6 ABC|1.25|D keyb||D"
Private Const STR_BG_LAYOUT1        As String = "я в е р т ъ у и о п ю <=|1.75|D " & _
                                                "а|||N с д ф г х й к л ш щ Done|1.75|B " & _
                                                "^^|||N з ь ц ж б н м ч . , ^^|1.25| " & _
                                                "?!123|1.5|D|N EN|1.5|D интервал|6 ?!123|1.25|D keyb||D"
Private Const STR_BG_LAYOUT2        As String = "Я В Е Р Т Ъ У И О П Ю <=|1.75|D " & _
                                                "А|||N С Д Ф Г Х Й К Л Ш Щ Done|1.75|B " & _
                                                "^^||L|N З Ь Ц Ж Б Н М Ч . , ^^|1.25|L " & _
                                                "?!123|1.5|D|N EN|1.5|D интервал|6 ?!123|1.25|D keyb||D"
Private Const STR_BG_LAYOUT3        As String = "1 2 3 4 5 6 7 8 9 0 <=||D " & _
                                                "|0.5|S|N - / : ; ( ) € & @ Done|1.75|B " & _
                                                "#+=|||N . , ? ! ' "" #+=|1.25 " & _
                                                "АБВ|1.5|D|N EN|1.5|D интервал|6 ABC|1.25|D keyb||D"
Private Const STR_BG_LAYOUT4        As String = "[ ] { } # % ^ * + = <=||D " & _
                                                "|0.5|S|N _ \ | ~ < > $ & · Done|1.75|B " & _
                                                "#+=||L|N . , ? ! ' "" #+=|1.25|L " & _
                                                "АБВ|1.5|D|N EN|1.5|D интервал|6 ABC|1.25|D keyb||D"
                                                
Private m_oCtlCancelMode        As Object
Private m_sNextLayout           As String
Private m_bIsIntl               As Boolean
                                                
Public Sub RegisterCancelMode(oCtl As Object)
    If Not m_oCtlCancelMode Is Nothing And Not m_oCtlCancelMode Is oCtl Then
        m_oCtlCancelMode.CancelMode
    End If
    Set m_oCtlCancelMode = oCtl
End Sub

Private Sub ctxTouchKeyboard1_ButtonClick(ByVal Index As Long)
    Dim sText           As String
    
    Select Case ctxTouchKeyboard1.ButtonCaption(Index)
    Case "EN", "BG", "keyb"
        If ctxTouchKeyboard1.ButtonCaption(Index) = "keyb" Then
            m_bIsIntl = Not m_bIsIntl
        Else
            m_bIsIntl = (ctxTouchKeyboard1.ButtonCaption(Index) <> "EN")
        End If
        ctxTouchKeyboard1.Layout = IIf(m_bIsIntl, STR_BG_LAYOUT1, STR_EN_LAYOUT1)
        m_sNextLayout = vbNullString
    Case "^^"
        If InStr(ctxTouchKeyboard1.ButtonTag(Index), "|L") > 0 Then
            ctxTouchKeyboard1.Layout = IIf(m_bIsIntl, STR_BG_LAYOUT1, STR_EN_LAYOUT1)
            m_sNextLayout = vbNullString
        Else
            ctxTouchKeyboard1.Layout = IIf(m_bIsIntl, STR_BG_LAYOUT2, STR_EN_LAYOUT2)
            m_sNextLayout = IIf(m_bIsIntl, STR_BG_LAYOUT1, STR_EN_LAYOUT1)
        End If
    Case "#+="
        If InStr(ctxTouchKeyboard1.ButtonTag(Index), "|L") > 0 Then
            ctxTouchKeyboard1.Layout = IIf(m_bIsIntl, STR_BG_LAYOUT3, STR_EN_LAYOUT3)
            m_sNextLayout = vbNullString
        Else
            ctxTouchKeyboard1.Layout = IIf(m_bIsIntl, STR_BG_LAYOUT4, STR_EN_LAYOUT4)
            m_sNextLayout = vbNullString
        End If
    Case "?!123"
        ctxTouchKeyboard1.Layout = IIf(m_bIsIntl, STR_BG_LAYOUT3, STR_EN_LAYOUT3)
        m_sNextLayout = vbNullString
    Case "ABC", "АБВ"
        ctxTouchKeyboard1.Layout = IIf(m_bIsIntl, STR_BG_LAYOUT1, STR_EN_LAYOUT1)
        m_sNextLayout = vbNullString
    Case "<="
        sText = "{BKSP}"
    Case "Done"
        sText = "{ENTER}"
    Case "space", "интервал"
        sText = " "
    Case "(", ")", "[", "]", "{", "}", "+", "^", "%", "~"
        sText = "{" & ctxTouchKeyboard1.ButtonCaption(Index) & "}"
    Case Else
        If Len(ctxTouchKeyboard1.ButtonCaption(Index)) = 1 Then
            sText = ctxTouchKeyboard1.ButtonCaption(Index)
        End If
    End Select
    pvSendKeys ctxTouchKeyboard1, sText
End Sub

Private Sub ctxTouchKeyboard2_ButtonClick(ByVal Index As Long)
    Dim sText           As String
    
    Select Case ctxTouchKeyboard2.ButtonCaption(Index)
    Case "<="
        sText = "{BKSP}"
    Case "(", ")", "[", "]", "{", "}", "+", "^", "%", "~"
        sText = "{" & ctxTouchKeyboard2.ButtonCaption(Index) & "}"
    Case Else
        If Len(ctxTouchKeyboard2.ButtonCaption(Index)) = 1 Then
            sText = ctxTouchKeyboard2.ButtonCaption(Index)
        End If
    End Select
    pvSendKeys ctxTouchKeyboard1, sText
End Sub

Private Sub Form_Load()
    ctxTouchKeyboard1.Layout = STR_EN_LAYOUT1
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ctxTouchKeyboard1.Move 120, ctxTouchKeyboard1.Top, ScaleWidth - 2 * 120
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_oCtlCancelMode Is Nothing Then
        m_oCtlCancelMode.CancelMode
        Set m_oCtlCancelMode = Nothing
    End If
End Sub

Private Sub pvSendKeys(oCtl As ctxTouchKeyboard, sText As String)
    If LenB(sText) <> 0 Then
        If sText Like "[а-яА-Я]" Then
            KeybLanguage = "bg"
        ElseIf sText Like "[a-zA-Z]" Then
            KeybLanguage = "en"
        End If
        With CreateObject("WScript.Shell")
            .SendKeys sText, False
        End With
        If LenB(m_sNextLayout) <> 0 Then
            oCtl.Layout = m_sNextLayout
            m_sNextLayout = vbNullString
        End If
    End If
End Sub
