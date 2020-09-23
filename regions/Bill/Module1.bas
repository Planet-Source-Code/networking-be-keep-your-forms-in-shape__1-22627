Attribute VB_Name = "Module1"
' window always on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40

' region handling
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3

' image handling
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest

' handles the moving of the borderless form
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' other
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' vars
Private FW As Long
Private FH As Long

Private Const BGColor = 65280

Public Function CreateRegionFromFile(strFile) As Long

    Dim T As Integer
    Dim R As Integer
    
    Form2.Image1.Picture = LoadPicture(strFile)
    Form2.Width = Form2.Image1.Width
    Form2.Height = Form2.Image1.Height
    Form2.Image1.Visible = False
    Form2.Picture = LoadPicture(strFile)
    Form2.ScaleMode = 3 ' pixels
    
    FW = Form2.ScaleWidth
    FH = Form2.ScaleHeight
    
    Dim DC As Long
    Dim BMP As Long
    Dim Pix As Long
    Dim rgnInv As Long
    Dim rgn As Long
    Dim rgnTotal As Long
    
    DC = CreateCompatibleDC(Form2.hdc)
    BMP = SelectObject(DC, LoadPicture(strFile))
    
    rgnTotal = CreateRectRgn(0, 0, FW, FH)
    rgnInv = CreateRectRgn(0, 0, FW, FH)
    
    CombineRgn rgnTotal, rgnTotal, rgnTotal, RGN_XOR
    
    For T = 0 To FW
        For R = 0 To FH
            Pix = GetPixel(DC, T, R)
            If Pix = BGColor Then
                rgn = CreateRectRgn(T, R, T + 1, R + 1)
                CombineRgn rgnTotal, rgnTotal, rgn, RGN_OR
                DeleteObject rgn
            End If
        Next R
    Next T

    CombineRgn rgnTotal, rgnTotal, rgnInv, RGN_XOR
    
    CreateRegionFromFile = rgnTotal

End Function
