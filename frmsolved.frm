VERSION 5.00
Begin VB.Form frmsolved 
   Caption         =   "Form3"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11340
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   11055
      Left            =   -600
      ScaleHeight     =   10995
      ScaleWidth      =   11640
      TabIndex        =   0
      Top             =   0
      Width           =   11700
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   960
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "You won the game"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   840
         TabIndex        =   3
         Top             =   6480
         Width           =   9375
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Caption         =   "Congratulations "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   2
         Top             =   1800
         Width           =   12855
      End
   End
End
Attribute VB_Name = "frmsolved"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
lblname.Caption = lblname.Caption & Text1.Text
End Sub

