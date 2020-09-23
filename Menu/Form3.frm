VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   LinkTopic       =   "Form3"
   ScaleHeight     =   1785
   ScaleWidth      =   1500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "REGULAR"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ITALIC"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "BOLD"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 On Error Resume Next
    Dim alngTempRegions(6) As Long
    Dim lngFormWidthInPixels As Long
    Dim lngFormHeightInPixels As Long
    Dim varA
    lngFormWidthInPixels = Form3.Width / Screen.TwipsPerPixelX
    lngFormHeightInPixels = Form3.Height / Screen.TwipsPerPixelY
    varA = CreateRoundRectRgn(0, 0, lngFormWidthInPixels, lngFormHeightInPixels, 24, 24)
    varA = SetWindowRgn(Form3.hwnd, varA, True)
    
    Dim e As Integer
        W = 0
        For e = Form3.Height + 150 To 0 Step -15
        Line (0, e)-(Form3.Width + 150, e), RGB(0, 0, W)
        W = W + 2
        Next
End Sub
