VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Select your region or effect"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "Search"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Eclips"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Bite"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Donut"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Holes"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Suck Effect"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Mickey"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "TV Effect"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Corners"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Circle"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Plus sign"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Triangle"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      Height          =   135
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Shape5 
      Height          =   375
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   375
   End
   Begin VB.Line Line32 
      X1              =   2280
      X2              =   2280
      Y1              =   3480
      Y2              =   120
   End
   Begin VB.Shape Shape4 
      Height          =   135
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape Shape3 
      Height          =   135
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   255
   End
   Begin VB.Line Line31 
      X1              =   1800
      X2              =   1800
      Y1              =   2040
      Y2              =   2160
   End
   Begin VB.Line Line30 
      X1              =   1680
      X2              =   1680
      Y1              =   2040
      Y2              =   2160
   End
   Begin VB.Line Line29 
      X1              =   1680
      X2              =   1800
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line28 
      X1              =   1680
      X2              =   1800
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line27 
      X1              =   1560
      X2              =   1680
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Line Line26 
      X1              =   1800
      X2              =   1920
      Y1              =   1920
      Y2              =   2040
   End
   Begin VB.Line Line25 
      X1              =   1800
      X2              =   1920
      Y1              =   2280
      Y2              =   2160
   End
   Begin VB.Line Line24 
      X1              =   1560
      X2              =   1680
      Y1              =   2160
      Y2              =   2280
   End
   Begin VB.Line Line23 
      X1              =   1800
      X2              =   1920
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line22 
      X1              =   1920
      X2              =   1920
      Y1              =   2280
      Y2              =   2160
   End
   Begin VB.Line Line21 
      X1              =   1560
      X2              =   1680
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line20 
      X1              =   1680
      X2              =   1560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line19 
      X1              =   1920
      X2              =   1920
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Line Line18 
      X1              =   1800
      X2              =   1920
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line17 
      X1              =   1560
      X2              =   1560
      Y1              =   2160
      Y2              =   2280
   End
   Begin VB.Line Line16 
      X1              =   1560
      X2              =   1560
      Y1              =   1920
      Y2              =   2040
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   1560
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   375
   End
   Begin VB.Line Line15 
      X1              =   1680
      X2              =   1800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line14 
      X1              =   1560
      X2              =   1680
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line13 
      X1              =   1800
      X2              =   1800
      Y1              =   1080
      Y2              =   960
   End
   Begin VB.Line Line12 
      X1              =   1680
      X2              =   1680
      Y1              =   1080
      Y2              =   960
   End
   Begin VB.Line Line11 
      X1              =   1920
      X2              =   1800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line10 
      X1              =   1560
      X2              =   1560
      Y1              =   960
      Y2              =   840
   End
   Begin VB.Line Line9 
      X1              =   1560
      X2              =   1680
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line8 
      X1              =   1680
      X2              =   1680
      Y1              =   840
      Y2              =   720
   End
   Begin VB.Line Line7 
      X1              =   1920
      X2              =   1920
      Y1              =   960
      Y2              =   840
   End
   Begin VB.Line Line6 
      X1              =   1920
      X2              =   1800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line5 
      X1              =   1800
      X2              =   1800
      Y1              =   840
      Y2              =   720
   End
   Begin VB.Line Line4 
      X1              =   1680
      X2              =   1800
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line3 
      X1              =   1560
      X2              =   1920
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   1560
      Y1              =   480
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   1560
      Y1              =   480
      Y2              =   120
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

Private Sub Command1_Click()
    
    Dim hRgn As Long
    Dim RetVal As Long
    
    Dim Points(2) As POINTAPI

    Points(0).x = 0
    Points(0).y = 0

    Points(1).x = FW
    Points(1).y = FH

    Points(2).x = 0
    Points(2).y = FH

    hRgn = CreatePolygonRgn(Points(0), 3, FillMode)

    Call SetWindowRgn(Form2.hWnd, hRgn, True)

End Sub

Private Sub Command10_Click()

    Dim r(4) As Long
    Dim RAll As Long
    
    Dim t As Integer
    
    RAll = CreateRectRgn(0, 0, FW, FH)
    RBite = CreateRectRgn(0, 0, FW, FH)
    
    For t = 0 To 4
        r(t) = CreateEllipticRgn(FW - 20, FH / 4 + (20 * t), FW + 20, FW / 4 + (20 * t) + 20)
    Next t

    Call CombineRgn(RBite, r(0), r(1), RGN_OR)

    For t = 2 To 4
        Call CombineRgn(RBite, RBite, r(t), RGN_OR)
    Next t

    Call CombineRgn(RAll, RAll, RBite, RGN_XOR)
    Call SetWindowRgn(Form2.hWnd, RAll, True)

End Sub

Private Sub Command11_Click()

    Dim R1 As Long, R2 As Long
    Dim RAll As Long
    
    For t = -FW To 2 * FW
        RAll = CreateEllipticRgn(0, 0, FW, FH)
        R1 = CreateEllipticRgn(t + 5, 5, t + FW - 5, FH - 5)
        R2 = CreateEllipticRgn(0, 0, FW, FH)
        Call CombineRgn(R2, R2, R1, RGN_AND)
        Call CombineRgn(RAll, RAll, R2, RGN_XOR)
        Call SetWindowRgn(Form2.hWnd, RAll, True)
        DoEvents
        Sleep 5
    Next t

End Sub

Private Sub Command12_Click()

    Dim r(2) As Long, RAll As Long
    
    ' used to store the center of each spotlight
    Dim P(2) As POINTAPI
    ' used to store offsets per light
    Dim O(2) As POINTAPI

    O(0).x = 2
    O(0).y = 1
    
    O(1).x = -1
    O(1).y = 2
    
    O(2).x = 3
    O(2).y = -1

    ' set startup postition of lights
    P(0).x = 0
    P(0).y = FH / 2

    P(1).x = FW / 2
    P(1).y = FH / 2
    
    P(1).x = FW
    P(1).y = FH / 2
    

    Dim t As Integer
    Dim s As Integer
    
    For t = 1 To 1000
        RAll = CreateRectRgn(0, 0, FW, FH)
        For s = 0 To 2
            P(s).x = P(s).x + O(s).x
            P(s).y = P(s).y + O(s).y
            If P(s).x < 0 Or P(s).x > FW Then O(s).x = O(s).x * -1
            If P(s).y < 0 Or P(s).y > FW Then O(s).y = O(s).y * -1
            r(s) = CreateEllipticRgn(P(s).x - 25, P(s).y - 25, P(s).x + 25, P(s).y + 25)
        Next s
        Call CombineRgn(RAll, r(0), r(1), RGN_OR)
        Call CombineRgn(RAll, RAll, r(2), RGN_OR)
        Call SetWindowRgn(Form2.hWnd, RAll, True)
        DoEvents
        Sleep 5
    Next t

End Sub

Private Sub Command2_Click()

    Dim hRgn As Long
    Dim RetVal As Long
    
    Dim Points(11) As POINTAPI

    Points(0).x = FW / 3
    Points(0).y = 0

    Points(1).x = FW / 3 * 2
    Points(1).y = 0

    Points(2).x = FW / 3 * 2
    Points(2).y = FH / 3

    Points(3).x = FW
    Points(3).y = FH / 3

    Points(4).x = FW
    Points(4).y = FH / 3 * 2

    Points(5).x = FW / 3 * 2
    Points(5).y = FH / 3 * 2

    Points(6).x = FW / 3 * 2
    Points(6).y = FH

    Points(7).x = FW / 3
    Points(7).y = FH

    Points(8).x = FW / 3
    Points(8).y = FH / 3 * 2

    Points(9).x = 0
    Points(9).y = FH / 3 * 2

    Points(10).x = 0
    Points(10).y = FH / 3

    Points(11).x = FW / 3
    Points(11).y = FH / 3


    hRgn = CreatePolygonRgn(Points(0), 12, FillMode)

    If hRgn = 0 Then
        MsgBox "Failed to create region"
    Else
        RetVal = SetWindowRgn(Form2.hWnd, hRgn, True)
        If RetVal = 0 Then
            MsgBox "Failed to set the region"
        End If
    End If

End Sub

Private Sub Command3_Click()

    Dim hRgn As Long
    Dim RetVal As Long
    
    FillMode = GetPolyFillMode(Form2.hdc)

    hRgn = CreateEllipticRgn(0, 0, Form2.Width / Screen.TwipsPerPixelX, Form2.Height / Screen.TwipsPerPixelY)
    
    If hRgn = 0 Then
        MsgBox "Failed to create region"
    Else
        RetVal = SetWindowRgn(Form2.hWnd, hRgn, True)
        If RetVal = 0 Then
            MsgBox "Failed to set the region"
        End If
    End If

End Sub

Private Sub Command4_Click()

    Dim hRgn As Long
    Dim RetVal As Long
    
    ' 4 * 3
    Dim Points(15) As POINTAPI
    Dim PointsPerPoly(4) As Long
    Dim lNrOfRegions As Long
    lNrOfRegions = 5
    
    ' nr of points per region
    PointsPerPoly(0) = 3
    PointsPerPoly(1) = 3
    PointsPerPoly(2) = 3
    PointsPerPoly(3) = 3
    PointsPerPoly(4) = 4
    
    ' top left
    Points(0).x = 0
    Points(0).y = 0

    Points(1).x = FW / 3
    Points(1).y = 0

    Points(2).x = 0
    Points(2).y = FH / 3
    
    ' top right
    Points(3).x = FW / 3 * 2
    Points(3).y = 0

    Points(4).x = FW
    Points(4).y = 0

    Points(5).x = FW
    Points(5).y = FH / 3
    
    ' bottom right
    Points(6).x = FW / 3 * 2
    Points(6).y = FH

    Points(7).x = FW
    Points(7).y = FH

    Points(8).x = FW
    Points(8).y = FH / 3 * 2
    
    ' bottom left
    Points(9).x = 0
    Points(9).y = FH

    Points(10).x = FW / 3
    Points(10).y = FH

    Points(11).x = 0
    Points(11).y = FH / 3 * 2
    
    ' middle box
    Points(12).x = FW / 3
    Points(12).y = FH / 3

    Points(13).x = FW / 3 * 2
    Points(13).y = FH / 3

    Points(14).x = FW / 3 * 2
    Points(14).y = FH / 3 * 2
    
    Points(15).x = FW / 3
    Points(15).y = FH / 3 * 2
    
    hRgn = CreatePolyPolygonRgn(Points(0), PointsPerPoly(0), lNrOfRegions, FillMode)
    
    If hRgn = 0 Then
        MsgBox "Failed to create region"
    Else
        RetVal = SetWindowRgn(Form2.hWnd, hRgn, True)
        If RetVal = 0 Then
            MsgBox "Failed to set the region"
        End If
    End If

End Sub

Private Sub Command5_Click()

    Dim hRgn As Long
    Dim RetVal As Long
    
    Dim Points(3) As POINTAPI

    Points(0).x = 0
    Points(0).y = 0

    Points(1).x = Form2.Width / Screen.TwipsPerPixelX
    Points(1).y = 0
    
    Points(2).x = Form2.Width / Screen.TwipsPerPixelX
    Points(2).y = Form2.Height / Screen.TwipsPerPixelY

    Points(3).x = 0
    Points(3).y = Form2.Height / Screen.TwipsPerPixelY

    Dim t As Integer
    
    For t = 0 To ((Form2.Height / Screen.TwipsPerPixelY) / 2) - 5 Step 5
        Points(0).y = Points(0).y + 5
        Points(1).y = Points(1).y + 5
        Points(2).y = Points(2).y - 5
        Points(3).y = Points(3).y - 5
        
        hRgn = CreatePolygonRgn(Points(0), 4, FillMode)
        Call SetWindowRgn(Form2.hWnd, hRgn, True)
    
        Sleep 5
        DoEvents
    
    Next t

    For t = 0 To ((Form2.Width / Screen.TwipsPerPixelX) / 2) - 5 Step 5
        Points(0).x = Points(0).x + 5
        Points(1).x = Points(1).x - 5
        Points(2).x = Points(2).x - 5
        Points(3).x = Points(3).x + 5
        
        hRgn = CreatePolygonRgn(Points(0), 4, FillMode)
        Call SetWindowRgn(Form2.hWnd, hRgn, True)
    
        Sleep 5
        DoEvents
    
    Next t

End Sub

Private Sub Command6_Click()

    Dim R1 As Long, R2 As Long, R3 As Long, RAll As Long
    
    ShowForm
    
    ' left ear
    R1 = CreateEllipticRgn(0, 0, FW / 3, FH / 3)
    ' right ear
    R2 = CreateEllipticRgn(FW / 3 * 2, 0, FW, FH / 3)
    ' face
    R3 = CreateEllipticRgn(FW / 4, FH / 4, FW / 4 * 3, FH / 4 * 3)

    RAll = CreateRectRgn(0, 0, FW, FH)

    Call CombineRgn(RAll, R1, R2, RGN_OR)
    Call CombineRgn(RAll, RAll, R3, RGN_OR)

    SetWindowRgn Form2.hWnd, RAll, True

End Sub

Private Sub Command7_Click()

    Dim Points(3) As POINTAPI
    Dim hRgn As Long

    Points(0).x = 0
    Points(0).y = 0

    Points(1).x = FW
    Points(1).y = 0
    
    Points(2).x = FW
    Points(2).y = FH

    Points(3).x = 0
    Points(3).y = FH
    
    Dim iTo As Integer
    Dim iX As Integer
    Dim iY As Integer

    iTo = (FW / 20)
    iX = FW / iTo / 2
    iY = FH / iTo / 2
    

    For t = i To iTo
        Points(0).x = Points(0).x + iX
        Points(0).y = Points(0).y + iY
        Points(1).x = Points(1).x - iX
        Points(1).y = Points(1).y + iY
        Points(2).x = Points(2).x - iX
        Points(2).y = Points(2).y - iY
        Points(3).x = Points(3).x + iX
        Points(3).y = Points(3).y - iY
        
        hRgn = CreatePolygonRgn(Points(0), 4, FillMode)
        Call SetWindowRgn(Form2.hWnd, hRgn, True)
    
        Sleep 5
        DoEvents
    
    Next t

End Sub

Private Sub Command8_Click()

    ' you must always recreate the region in order to work
    
    Dim RAll As Long
    Dim R1 As Long, R2 As Long, R3 As Long
    
    R1 = CreateEllipticRgn(FW / 8, FH / 8, FW / 8 * 3, FH / 8 * 3)
    R2 = CreateEllipticRgn(FW / 8 * 3, FH / 8 * 3, FW / 8 * 5, FH / 8 * 5)
    R3 = CreateEllipticRgn(FW / 8 * 5, FH / 8 * 5, FW / 8 * 7, FH / 8 * 7)
    
    RAll = CreateRectRgn(0, 0, FW, FH)

    Call CombineRgn(RAll, RAll, R1, RGN_XOR)
    Call SetWindowRgn(Form2.hWnd, RAll, True)

    DoEvents
    Sleep 500

    RAll = CreateRectRgn(0, 0, FW, FH)
    Call CombineRgn(RAll, RAll, R1, RGN_XOR)
    Call CombineRgn(RAll, RAll, R2, RGN_XOR)
    Call SetWindowRgn(Form2.hWnd, RAll, True)

    DoEvents
    Sleep 500

    RAll = CreateRectRgn(0, 0, FW, FH)
    Call CombineRgn(RAll, RAll, R1, RGN_XOR)
    Call CombineRgn(RAll, RAll, R2, RGN_XOR)
    Call CombineRgn(RAll, RAll, R3, RGN_XOR)
    Call SetWindowRgn(Form2.hWnd, RAll, True)

End Sub

Private Sub Command9_Click()

    Dim R1 As Long, R2 As Long, RAll As Long

    R1 = CreateEllipticRgn(0, 0, FW, FH)
    R2 = CreateEllipticRgn(FW / 4, FH / 4, FW / 4 * 3, FH / 4 * 3)
    RAll = CreateRectRgn(0, 0, FW, FH)

    Call CombineRgn(RAll, R1, R2, RGN_XOR)
    Call SetWindowRgn(Form2.hWnd, RAll, True)

End Sub

Private Sub Form_Load()

    ShowForm
    
    FW = Form2.Width / Screen.TwipsPerPixelX
    FH = Form2.Height / Screen.TwipsPerPixelY

    FillMode = GetPolyFillMode(Form2.hdc)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' unload form2
    On Error Resume Next
    Unload Form2

End Sub

Private Sub ShowForm()

    Form2.Show
    Form2.Left = Form1.Left + Form1.Width
    Form2.Top = Form1.Top

End Sub
