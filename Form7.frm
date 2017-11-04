VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00000080&
   Caption         =   "Form7"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2 
      Caption         =   "FEMALE"
      Height          =   255
      Left            =   12240
      TabIndex        =   15
      Top             =   5400
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      Height          =   255
      Left            =   10200
      TabIndex        =   14
      Top             =   5400
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   15600
      Top             =   8760
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\abhisek\Desktop\VB6 project\Database2.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\abhisek\Desktop\VB6 project\Database2.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "STUDENT1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   12240
      TabIndex        =   13
      Top             =   9480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   9000
      TabIndex        =   12
      Top             =   9480
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      DataField       =   "address"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   10080
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   7680
      Width           =   4095
   End
   Begin VB.TextBox Text5 
      DataField       =   "phone_number"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   10080
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   6480
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      DataField       =   "password"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   10080
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   3840
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      DataField       =   "username"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   10080
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      DataField       =   "student_name"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   10080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   7920
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "PHONE NUMBER"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "GENDER"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "STUDENT REGISTRATION"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.Fields("student_name") = Text1.Text
Adodc1.Recordset.Fields("username") = Text2.Text
Adodc1.Recordset.Fields("password") = Text3.Text
If Option1 = True Then
Adodc1.Recordset.Fields("gender") = Option1.Caption
Else
Adodc1.Recordset.Fields("gender") = Option2.Caption
End If




Adodc1.Recordset.Fields("phone_number") = Text5.Text
Adodc1.Recordset.Fields("address") = Text6.Text
If Text1.Text = "" Then
MsgBox "NAME FIELD CANNOT BE EMPTY", vbCritical
ElseIf Text2.Text = "" Then
MsgBox "USER NAME FIELD CANNOT BE EMPTY", vbCritical
ElseIf Text3.Text = "" Then
MsgBox "PASSWORD CANNOT BE EMPTY"
ElseIf Option1 <> True And Option2 <> True Then
MsgBox "PLEASE ENTER A GENDER"
ElseIf Len(Text5.Text) <> 10 Or Not IsNumeric(Text5.Text) Then
MsgBox "PLEASE ENTER A VALID PHONE NUMBER"
ElseIf Text6.Text = "" Then
MsgBox "PLEASE ENTER A VALID ADDRESS"
Else
Adodc1.Recordset.Update
MsgBox "registration successfull", vbInformation
Form5.Show
Unload Me
End If




















End Sub

Private Sub Command2_Click()
Form9.Show
Unload Me

End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew


End Sub
