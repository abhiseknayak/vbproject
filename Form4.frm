VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00000080&
   Caption         =   "Form4"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPLOAD"
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
      Left            =   840
      TabIndex        =   18
      Top             =   4800
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   600
      ScaleHeight     =   2235
      ScaleWidth      =   2715
      TabIndex        =   17
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "subject_alloted"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Form4.frx":0000
      Left            =   7200
      List            =   "Form4.frx":000D
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   4680
      Width           =   6255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11880
      Top             =   9120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
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
      RecordSource    =   "TEACHER"
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
      BackColor       =   &H000080FF&
      Caption         =   "CANCEL"
      Height          =   735
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton regidterbtn 
      BackColor       =   &H0000FF00&
      Caption         =   "REGISTER"
      Height          =   735
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9120
      Width           =   2055
   End
   Begin VB.TextBox txtaddress 
      DataField       =   "address"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   7200
      TabIndex        =   11
      Text            =   "Text7"
      Top             =   7560
      Width           =   6495
   End
   Begin VB.TextBox txtphone 
      DataField       =   "phone number"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   7200
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   6600
      Width           =   6375
   End
   Begin VB.TextBox txtqual 
      DataField       =   "qualification"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   5520
      Width           =   6375
   End
   Begin VB.TextBox txtpassword 
      DataField       =   "password"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   3600
      Width           =   6375
   End
   Begin VB.TextBox txtusername 
      DataField       =   "username"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   7200
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2640
      Width           =   6375
   End
   Begin VB.TextBox txtname 
      DataField       =   "teacher_name"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   7200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1800
      Width           =   6375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "ADDRESS"
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   4320
      TabIndex        =   13
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "PHONE NUMBER"
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   4320
      TabIndex        =   12
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "QUALIFICATION"
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   4320
      TabIndex        =   9
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "SUBJECT ALLOTED"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "PASSWORD"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "USER NAME"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "NAME"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "TEACHER REGISTRATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   7200
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim str As String

Private Sub Command1_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "jpeg|*.jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)
End Sub

Private Sub Command2_Click()
Form3.Show
Unload Me

End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
End Sub


Private Sub regidterbtn_Click()
Adodc1.Recordset.Fields("teacher_name") = txtname.Text
Adodc1.Recordset.Fields("username") = txtusername.Text
Adodc1.Recordset.Fields("password") = txtpassword.Text
Adodc1.Recordset.Fields("subject_alloted") = Combo1.Text
Adodc1.Recordset.Fields("qualification") = txtqual.Text
Adodc1.Recordset.Fields("phone number") = txtphone.Text
Adodc1.Recordset.Fields("address") = txtaddress.Text
Adodc1.Recordset.Fields("photo") = str

If txtname.Text = "" Then
MsgBox "NAME field cannot be empty"
ElseIf txtusername.Text = "" Then
MsgBox "USERNAME field cannot be empty"
ElseIf Not Picture1.Picture Then
MsgBox "PLEASE UPLOAD A PICTURE"
ElseIf Len(txtpassword.Text) < 8 Then
MsgBox "PASSWORD should be atleast 8 characters long"
ElseIf Combo1.Text = "" Then
MsgBox "registration cannot be completed until subject is allotted"
ElseIf txtqual.Text = "" Then
MsgBox "Please provide your qualification for completing registration"
ElseIf txtphone.Text = "" Then
MsgBox "PHONE NUMBER is not valid"
ElseIf Len(txtphone.Text) <> 10 Then
MsgBox "PHONE NUMBER is invalid"
ElseIf Not IsNumeric(txtphone.Text) Then
MsgBox "PHONE NUMBER is invalid"

ElseIf txtaddress.Text = "" Then
MsgBox "enter a valid address"
Else
Adodc1.Recordset.Update
MsgBox "successful", vbInformation
Me.Hide
Form3.Show

End If
End Sub
