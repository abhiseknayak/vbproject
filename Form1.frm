VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "A  Student"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   8160
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A  Teacher"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   5040
      TabIndex        =   3
      Top             =   1560
      Width           =   7455
   End
   Begin VB.Image Image4 
      Height          =   1905
      Left            =   9960
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1995
   End
   Begin VB.Image Image3 
      Height          =   1830
      Left            =   9960
      Picture         =   "Form1.frx":7CD7
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1950
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "YOU ARE......"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   3840
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim c As Integer

Private Sub Command1_Click()
Unload Me
Form2.Show

End Sub

Private Sub Command2_Click()
Unload Me
Form3.Show

End Sub

Private Sub Command3_Click()
Unload Me
Form5.Show


End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

End Sub

Private Sub Label2_Click()
c = c + 1
If c > 10 Then
c = 0
Form6.Show
Unload Me

End If


End Sub
