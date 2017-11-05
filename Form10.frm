VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form10"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "FIRST STUDENT"
      Height          =   615
      Left            =   13800
      TabIndex        =   25
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LAST STUDENT"
      Height          =   615
      Left            =   3480
      TabIndex        =   24
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "BACK"
      Height          =   615
      Left            =   8640
      TabIndex        =   23
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT"
      Height          =   615
      Left            =   11280
      TabIndex        =   11
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPDATE"
      Height          =   615
      Left            =   6000
      TabIndex        =   10
      Top             =   8040
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   12000
      TabIndex        =   9
      Text            =   "Text8"
      Top             =   6360
      Width           =   3255
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   12000
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   5280
      Width           =   3135
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   7080
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   6360
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   7080
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   5280
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   6480
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2520
      Width           =   5055
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   12720
      ScaleHeight     =   2595
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   495
      Left            =   10800
      TabIndex        =   22
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   375
      Left            =   10800
      TabIndex        =   21
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   840
      TabIndex        =   18
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "DIGITAL LOGIC AND DESIGN"
      Height          =   495
      Left            =   12480
      TabIndex        =   16
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "LOGIC BUILDING IN C"
      Height          =   495
      Left            =   7560
      TabIndex        =   15
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "COMPUTER FUNDAMENTALS"
      Height          =   495
      Left            =   2520
      TabIndex        =   14
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "ROLL NUMBER"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "NAME"
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ALLOCATE MARKS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub Command1_Click()
rs.Fields("ct_sub1") = Val(Text3.Text)
rs.Fields("ct_sub2") = Val(Text5.Text)
rs.Fields("ct_sub3") = Val(Text7.Text)
rs.Fields("ese_sub1") = Val(Text4.Text)
rs.Fields("ese_sub2") = Val(Text6.Text)
rs.Fields("ese_sub3") = Val(Text8.Text)
MsgBox "records updated"
rs.Update


End Sub

Private Sub Command2_Click()
rs.MoveNext
If Not rs.EOF Then
display
Else
rs.MoveFirst
display
End If

End Sub

Private Sub Command3_Click()
rs.MoveLast
display


End Sub
Sub display()
Text1.Text = rs!student_name
Text2.Text = rs!roll_no
Text3.Text = rs!ct_sub1
Text4.Text = rs!ese_sub1
Text5.Text = rs!ct_sub2
Text6.Text = rs!ese_sub2
Text7.Text = rs!ct_sub3
Text8.Text = rs!ese_sub3


End Sub

Private Sub Command4_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
display
Else
display
End If

End Sub

Private Sub Command5_Click()
rs.MoveFirst
display
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\abhisek\Desktop\VB6 project\Database2.mdb;Persist Security Info=False"
rs.Open "Select * from STUDENT1", con, adOpenDynamic, adLockPessimistic
display
If Form2.sub_imp = "COMPUTER FUNDAMENTALS" Then
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
ElseIf Form2.sub_imp = "PSLBC" Then
Text3.Enabled = False
Text4.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
ElseIf Form2.sub_imp = "DIGITAL LOGIC AND DESIGN" Then
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
End If





End Sub
