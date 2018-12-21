VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   8100
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10752
   LinkTopic       =   "Form3"
   ScaleHeight     =   8100
   ScaleWidth      =   10752
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   432
      Left            =   756
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   756
      Width           =   1524
   End
   Begin Project1.ctxTouchKeyboard ctxTouchKeyboard1 
      Height          =   3456
      Left            =   0
      Top             =   2016
      Width           =   9672
      _extentx        =   17060
      _extenty        =   6096
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oCtlCancelMode        As Object

Private Const DEF_LAYOUT1           As String = "q w e r t y u i o p <=|1.25|D " & _
                                                "|0.1|S|N a s d f g h j k l Done|1.85|B " & _
                                                "^|||N z x c v b n m ! ? ^|1.25| " & _
                                                "?!123|3|D|N _|6 ?!123|1.25|D keyb||D"
Private Const DEF_LAYOUT2           As String = "Q W E R T Y U I O P <=|1.25|D " & _
                                                "|0.5|S|N A S D F G H J K L Done|1.85|B " & _
                                                "^||L|N Z X C V B N M ! ? ^|1.25|L " & _
                                                "?!123|3|D|N _|6 ?!123|1.25|D keyb||D"
                                                
Public Sub RegisterCancelMode(oCtl As Object)
    If Not m_oCtlCancelMode Is Nothing And Not m_oCtlCancelMode Is oCtl Then
        m_oCtlCancelMode.CancelMode
    End If
    Set m_oCtlCancelMode = oCtl
End Sub

Private Sub ctxTouchKeyboard1_ButtonClick(ByVal Index As Long)
    If Left$(ctxTouchKeyboard1.ButtonCaption(Index), 1) = "^" Then
        If InStr(ctxTouchKeyboard1.ButtonTag(Index), "|L") > 0 Then
            ctxTouchKeyboard1.Layout = DEF_LAYOUT1
        Else
            ctxTouchKeyboard1.Layout = DEF_LAYOUT2
        End If
    Else
        Text1.SelText = ctxTouchKeyboard1.ButtonCaption(Index)
    End If
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

