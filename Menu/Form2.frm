VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1380
   DrawWidth       =   2
   LinkTopic       =   "Form2"
   ScaleHeight     =   1695
   ScaleWidth      =   1380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PASTE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COPY"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CUT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "Form2"
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
    lngFormWidthInPixels = Form2.Width / Screen.TwipsPerPixelX
    lngFormHeightInPixels = Form2.Height / Screen.TwipsPerPixelY
    varA = CreateRoundRectRgn(0, 0, lngFormWidthInPixels, lngFormHeightInPixels, 24, 24)
    varA = SetWindowRgn(Form2.hwnd, varA, True)
    
    Dim q As Integer
        W = 0
        For q = Form2.Height + 150 To 0 Step -15
        Line (0, q)-(Form2.Width + 150, q), RGB(0, 0, W)
        W = W + 2
        Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.BackColor = &HFFC0C0
    Label2.BackColor = &HFFC0C0
    Label3.BackColor = &HFFC0C0
End Sub

Private Sub Label1_Click()
      MsgBox "Put code here or...", vbInformation, "Menu By Nobody"
     Unload Form2
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.BackColor = vbYellow
End Sub

Private Sub Label2_Click()
     MsgBox "Put code here or...", vbInformation, "Menu By Nobody"
     Unload Form2
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2.BackColor = vbYellow
End Sub

Private Sub Label3_Click()
  MsgBox "Put code here or...", vbInformation, "Menu By Nobody"
     Unload Form2
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = vbYellow
End Sub
