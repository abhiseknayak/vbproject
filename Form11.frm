VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00000080&
   Caption         =   "Form11"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form11"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12960
      MaskColor       =   &H00000080&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   3600
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   6120
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   13320
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   13320
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   8400
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   4200
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   8400
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000080&
      Caption         =   "YOUR GP"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1920
      TabIndex        =   17
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000080&
      Caption         =   "END SEMESTER"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   11760
      TabIndex        =   16
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000080&
      Caption         =   "CLASS TEST"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11760
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000080&
      Caption         =   "END SEMESTER"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   6840
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000080&
      Caption         =   "CLASS TEST"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6840
      TabIndex        =   13
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000080&
      Caption         =   "END SEMESTER"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Caption         =   "CLASS TEST"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000080&
      Caption         =   "DIGTAL LOGIC AND DESIGN"
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   13440
      TabIndex        =   10
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000080&
      Caption         =   "LOGIC BUILDING USING C"
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   8400
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Caption         =   "COMPUTER FUNDAMENTALS"
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   3720
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "YOUR MARKS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   7920
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset



Private Sub Command1_Click()
Form8.Show
Unload Me

End Sub

Private Sub Form_Load()
con1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\abhisek\Desktop\VB6 project\Database2.mdb;Persist Security Info=False"
rs1.Open "Select * from STUDENT1 where username='" + Form8.user_imp + "'", con1, adOpenDynamic, adLockPessimistic
display
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Dim i As Double
i = Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text)
i = (i / 300) * 100

i = i / 10
Text7.Text = i
Text7.Enabled = False






End Sub
Sub display()
Text1.Text = rs1!ct_sub1

Text2.Text = rs1!ese_sub1

Text3.Text = rs1!ct_sub2

Text4.Text = rs1!ese_sub2
Text5.Text = rs1!ct_sub3
Text6.Text = rs1!ese_sub3

End Sub
