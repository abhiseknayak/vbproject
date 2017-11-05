VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000080&
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "LOGIN"
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   6240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4200
      Width           =   6975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "FOR VERIFICATION"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "ENTER PASS-CODE "
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "WELCOME ADMINISTRATOR"
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
      Height          =   615
      Left            =   6480
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "hello_world" Then
Unload Me

Form14.Show
End If




End Sub

Private Sub Form_Load()
Text1.Text = "PLEASE ENTER VERIFICATION CODE"
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub
