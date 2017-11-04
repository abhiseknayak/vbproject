VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00000080&
   Caption         =   "Form8"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "LOG OUT"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "MESSAGES"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "CHECK MARKS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   7920
      TabIndex        =   18
      Top             =   7680
      Width           =   3615
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7920
      TabIndex        =   16
      Top             =   6840
      Width           =   3615
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   5760
      TabIndex        =   15
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7920
      TabIndex        =   14
      Top             =   6000
      Width           =   3615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   5760
      TabIndex        =   13
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7920
      TabIndex        =   12
      Top             =   5160
      Width           =   3615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   4200
      Width           =   3615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7920
      TabIndex        =   8
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7920
      TabIndex        =   6
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Caption         =   "Label2"
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
      Height          =   615
      Left            =   9960
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000080&
      Caption         =   "WELCOME"
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
      Height          =   615
      Left            =   7560
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Form9.Show
MsgBox "LOGOUT SUCCESSFULL"
Unload Me

End Sub
