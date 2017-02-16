VERSION 5.00
Begin VB.Form frmabout 
   Caption         =   "About/Help"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Domains"
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Custom Tags"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Custom Tags"
      Height          =   3495
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   2295
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "frmabout.frx":0000
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   $"frmabout.frx":04AE
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Hosting Domains"
      Height          =   3495
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3495
      Begin VB.Label Label7 
         Caption         =   $"frmabout.frx":053D
         Height          =   3135
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "About"
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3495
      Begin VB.Label Label5 
         Caption         =   "Thanks To: SnapperHead Software, MasterMac, PlanetSourceCode, And my friends from the Bg3 team"
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Email: Support@Novaslp.net "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Need help, Info, questions or want to comment?"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   $"frmabout.frx":0749
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Desktop Server By Nova1313"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
End Sub

Private Sub Command3_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
End Sub

Private Sub Command4_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()

End Sub
