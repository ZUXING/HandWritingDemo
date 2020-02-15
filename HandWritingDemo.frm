VERSION 5.00
Begin VB.Form HandWritingDemo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   18
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   5910
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label UpX 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label DownX 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "HandWritingDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Print "现在请在程序上划一段距离."
End Sub

Private Sub Form_DblClick()
If Me.Width = 6000 Then
    Me.Width = 7200
Else
    Me.Width = 6000
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cls
DownX.Caption = X
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpX.Caption = X

Select Case UpX - DownX
    Case Is > 3000
        MsgBox "你做出了向右划过的动作, 要我做什么? "
    Case Is < -3000
        MsgBox "你做出了向左划过的动作, 要我做什么? "
End Select

End Sub
