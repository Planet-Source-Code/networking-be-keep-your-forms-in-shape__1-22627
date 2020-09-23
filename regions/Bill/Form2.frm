VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   ClientHeight    =   1305
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   1425
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1305
   ScaleWidth      =   1425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   600
      Top             =   840
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Direction As String
Private First As Boolean

Private Sub Form_Load()

    Direction = "UP"
    First = True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Timer1.Enabled = False
    If Button = 1 Then
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    ElseIf Button = 2 Then
        Direction = "DOWN"
    End If
    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

    If Direction = "UP" Then
        If Me.Top > (-1 * Me.Height) Then
            Me.Top = Me.Top - 10
        Else
            Me.Top = Screen.Height
            Me.Left = CLng(Rnd * (Screen.Width - Me.Width))
        End If
    Else
        If Me.Top < Screen.Height Then
            Me.Top = Me.Top + 100
        Else
            End
        End If
    End If

End Sub
