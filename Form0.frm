VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "USER NAME"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   7875
   Begin VB.CommandButton cmdhar 
      Caption         =   "Hard"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdmed 
      Caption         =   "Medium"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdkid 
      Caption         =   "Kid"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   11055
      Left            =   0
      ScaleHeight     =   10995
      ScaleWidth      =   15075
      TabIndex        =   1
      Top             =   -360
      Width           =   15135
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   6360
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtuname 
         DataField       =   "NAME"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   0
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lbluname 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enter Your name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdhar_Click()
frmmainhar.Uname = txtuname.Text
Unload Me
Load frmmainhar
End Sub
Private Sub cmdkid_Click()
frmmainkid.Uname = txtuname.Text
Unload Me
Load frmmainkid
End Sub
Private Sub cmdmed_Click()
frmmain.Uname = txtuname.Text
Unload Me
Load frmmain
End Sub
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
txtuname.Text = Uname
End Sub
