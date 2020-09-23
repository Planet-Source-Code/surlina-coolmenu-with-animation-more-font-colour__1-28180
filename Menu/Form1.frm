VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5040
      Top             =   3720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   150
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   300
      Left            =   5400
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   45
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   4095
      Left            =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Menu By Nobody 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   840
      MouseIcon       =   "Form1.frx":030A
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Dim A As Integer
Dim P As Integer
Private Sub Form_Load()
   A = 0: W = 0
   Label6.Caption = Label6.Caption & "Please vote." & _
       " If you have some" & vbCrLf & " suggestion on this programm, send" & _
       vbCrLf & " me" & _
       " e-mail or leave on Planet..." & vbCrLf & "Today is: " & Date & _
       ". Time: " & Time
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label1.BackColor = &H8000000F
   Label2.BackColor = &H8000000F
   Label3.BackColor = &H8000000F
End Sub

Private Sub Label1_Click()
   Form2.Top = Label1.Top + Form1.Top + 400
   Form2.Left = Label1.Left + Form1.Left
   X = 1: P = 1
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.BackColor = vbWhite
    Label2.BackColor = &H8000000F
    Label3.BackColor = &H8000000F
End Sub

Private Sub Label2_Click()
   Form3.Top = Label2.Top + Form1.Top + 400
   Form3.Left = Label2.Left + Form1.Left
   X = 1: P = 2
     
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label1.BackColor = &H8000000F
   Label2.BackColor = vbWhite
End Sub

Private Sub Label3_Click()
   Form4.Top = Label3.Top + Form1.Top + 400
   Form4.Left = Label3.Left + Form1.Left
     P = 3: X = 1
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.BackColor = &H8000000F
    Label3.BackColor = vbWhite
End Sub

Private Sub Label5_Click()
        End
End Sub

Private Sub Timer1_Timer()
 If X = 1 Then
   Select Case P
     Case 1
         A = A + 150
         Form2.Width = A
         Form2.Height = A
         Form2.Show
         If A > 2000 Then X = 0: A = 0
     Case 2
         A = A + 150
         Form3.Width = 2000
         Form3.Height = A
         Form3.Show
         If A > 2000 Then X = 0: A = 0
     Case 3
         A = A + 150
         Form4.Width = 2000
         Form4.Height = A
         Form4.Show
         If A > 2000 Then X = 0: A = 0
     End Select
     End If
End Sub
