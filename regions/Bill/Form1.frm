VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Loading..."
   ClientHeight    =   435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2115
   LinkTopic       =   "Form1"
   ScaleHeight     =   435
   ScaleWidth      =   2115
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Bill is coming, please wait..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Face As Long

Private Sub Form_Load()

    Me.Show
    DoEvents
    
    Face = CreateRegionFromFile(App.Path & "\bill.bmp")
    
    Form2.Show
    SetWindowRgn Form2.hwnd, Face, True
    ' make window stay on top
    SetWindowPos Form2.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW
    Me.Visible = False

End Sub
