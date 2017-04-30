VERSION 5.00
Begin VB.Form frmkidstat 
   BackColor       =   &H00FFC0FF&
   Caption         =   "STATISTICS TABLES"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11430
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   11430
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   4800
      TabIndex        =   34
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "NAME"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   9600
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      DataField       =   "SCORE"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   10200
      TabIndex        =   31
      Text            =   "Text2"
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text3 
      DataField       =   "TIMEELPSED"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   10920
      TabIndex        =   30
      Text            =   "Text3"
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOP FIVE GAMMERS DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   360
      Left            =   3480
      TabIndex        =   33
      Top             =   120
      Width           =   4545
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   9
      Left            =   7080
      TabIndex        =   29
      Top             =   5400
      Width           =   2505
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   8
      Left            =   7080
      TabIndex        =   28
      Top             =   4920
      Width           =   2505
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   7
      Left            =   7080
      TabIndex        =   27
      Top             =   4440
      Width           =   2505
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   6
      Left            =   7080
      TabIndex        =   26
      Top             =   3960
      Width           =   2505
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   5
      Left            =   7080
      TabIndex        =   25
      Top             =   3480
      Width           =   2505
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   9
      Left            =   1080
      TabIndex        =   24
      Top             =   5400
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   8
      Left            =   1080
      TabIndex        =   23
      Top             =   4920
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   9
      Left            =   5280
      TabIndex        =   22
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   8
      Left            =   5280
      TabIndex        =   21
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   7
      Left            =   5280
      TabIndex        =   20
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   6
      Left            =   5280
      TabIndex        =   19
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   5
      Left            =   5280
      TabIndex        =   18
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   7
      Left            =   1080
      TabIndex        =   17
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   6
      Left            =   1080
      TabIndex        =   16
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   5
      Left            =   1080
      TabIndex        =   15
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   4
      Left            =   7080
      TabIndex        =   14
      Top             =   3000
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   4
      Left            =   5280
      TabIndex        =   13
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   4
      Left            =   1080
      TabIndex        =   12
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   3
      Left            =   7080
      TabIndex        =   11
      Top             =   2520
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   3
      Left            =   5280
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   3
      Left            =   1080
      TabIndex        =   9
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   2
      Left            =   7080
      TabIndex        =   8
      Top             =   2040
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   2
      Left            =   5280
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   2
      Left            =   1080
      TabIndex        =   6
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   1
      Left            =   7080
      TabIndex        =   5
      Top             =   1560
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   1
      Left            =   5280
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   0
      Left            =   7080
      TabIndex        =   2
      Top             =   1080
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   0
      Left            =   5280
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
   End
End
Attribute VB_Name = "frmkidstat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim cdes As Integer
cdes = MsgBox("Do you want new game", vbYesNo, "New Game")
If cdes = vbYes Then
   Unload Me
   Load frmmainkid
Else
   Unload Me
End If
End Sub
Private Sub Form_Load()
Dim con As New ADODb.Connection
con.Provider = "Microsoft.Jet.OLEDB.4.0"
con.Open App.Path + "\db2.mdb"
Dim rs2 As New ADODb.Recordset
rs2.Open "SELECT * From TABLE1 ORDER BY TABLE1.SCORE DESC", con, adOpenDynamic, adLockOptimistic
Dim i As Integer
rs2.MoveFirst
For i = 0 To 9
Text1.Text = rs2.Fields("name").Value
Text2.Text = rs2.Fields("score").Value
Text3.Text = rs2.Fields("timeelpsed").Value
If Text1.Text = "" Then
 Label1(i).Caption = "-XX-"
Else
Label1(i).Caption = Text1.Text
End If
If Text2.Text = 0 Then
 Label2(i).Caption = "-no-"
Else
Label2(i).Caption = Text2.Text
End If
Label3(i).Caption = Text3.Text
rs2.MoveNext
Next
rs2.Close
con.Close
End Sub
