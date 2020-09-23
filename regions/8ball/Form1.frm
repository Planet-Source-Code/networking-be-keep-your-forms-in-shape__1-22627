VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1080
      Top             =   2520
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "end"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   1080
      Shape           =   2  'Oval
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FW As Long ' form width (in pixels)
Dim FH As Long ' form height (in pixels)
Dim FillMode As Long
Dim rgn As Long
Dim rgn2 As Long
Dim StartX As Long, StartY As Long
Dim EndX As Long, EndY As Long
Dim Shaked As Integer
Dim ShakedEnough As Boolean

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&

End Sub

Private Sub Form_Load()

    FW = Me.Width / Screen.TwipsPerPixelX
    FH = Me.Height / Screen.TwipsPerPixelY
    
    rgn = CreateEllipticRgn(0, 0, FW, FH)

    Call SetWindowRgn(Me.hwnd, rgn, True)
    
    GetPosition StartX, StartY

End Sub

Private Function GetPosition(X As Long, Y As Long)

    Dim Loc As RECT

    Call GetWindowRect(Me.hwnd, Loc)
    X = Loc.Left
    Y = Loc.Top

End Function

Private Sub Label4_Click()
    End
End Sub

Private Sub Timer1_Timer()

    GetPosition EndX, EndY
        
    If EndX <> StartX Or EndY <> StartY Then
        If Shaked < 1 Then Shaked = 0
        Shaked = Shaked + 1
        StartX = EndX
        StartY = EndY
    Else
        If Shaked < 1 Then
            Shaked = Shaked - 1
        Else
            Shaked = 0
        End If
    End If

    If Shaked > 10 Then ShakedEnough = True

    If Shaked < -10 And ShakedEnough Then
        ShakedEnough = False
    
        Randomize
    
        If (Int(Rnd * 10000) Mod 2) = 0 Then
            FadeLabel Label3, 0, 0, 0, 255, 255, 255
            Label3.Visible = False
            Label2.Visible = True
            FadeLabel Label2, 255, 255, 255, 0, 0, 0
            Sleep 2000
            FadeLabel Label2, 0, 0, 0, 255, 255, 255
            Label2.Visible = False
            Label3.Visible = True
            FadeLabel Label3, 255, 255, 255, 0, 0, 0
        Else
            FadeLabel Label3, 0, 0, 0, 255, 255, 255
            Label3.Visible = False
            Label1.Visible = True
            FadeLabel Label1, 255, 255, 255, 0, 0, 0
            Sleep 2000
            FadeLabel Label1, 0, 0, 0, 255, 255, 255
            Label1.Visible = False
            Label3.Visible = True
            FadeLabel Label3, 255, 255, 255, 0, 0, 0
        End If
        
    End If


End Sub

Private Sub FadeLabel(lbl As Label, ByVal r1, ByVal g1, ByVal b1, ByVal r2, ByVal g2, ByVal b2)

    Dim PlusR As Single, PlusG As Single, PlusB As Single

    PlusR = (r2 - r1) / 100
    PlusG = (g2 - g1) / 100
    PlusB = (b2 - b1) / 100

    For t = 1 To 100
        lbl.ForeColor = RGB(r1, g1, b1)
        r1 = r1 + PlusR
        g1 = g1 + PlusG
        b1 = b1 + PlusB
        DoEvents
        Sleep 10
    Next t

End Sub
