VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form splash 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9660
   LinkTopic       =   "Form5"
   ScaleHeight     =   5985
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   7920
      Top             =   720
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   4920
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      Max             =   105
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Caption         =   "VERSION 1.0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "EXAMINATION SYSTEM "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label lblstat 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblstatus 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   5775
      Left            =   120
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer1.Enabled = True


End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
lblstatus.Caption = "Loading..Please Wait..."
lblstat.Caption = ProgressBar1.Value & "%"



If ProgressBar1.Value = ProgressBar1.Max Then
Timer1.Enabled = False
Unload Me
Form1.Show

End If




End Sub
