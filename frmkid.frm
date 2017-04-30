VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmmainkid 
   BackColor       =   &H00FFC0C0&
   Caption         =   "INFO MAZE"
   ClientHeight    =   8040
   ClientLeft      =   720
   ClientTop       =   1845
   ClientWidth     =   10515
   Icon            =   "frmkid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   10515
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   360
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF8080&
      Height          =   2415
      Left            =   8640
      ScaleHeight     =   2355
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   3
         Left            =   360
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   5
         Top             =   1680
         Width           =   540
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   2
         Left            =   360
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   3
         Top             =   1080
         Width           =   540
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   1
         Left            =   360
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   2
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0E0FF&
      Height          =   4780
      Left            =   3720
      ScaleHeight     =   4725
      ScaleWidth      =   4560
      TabIndex        =   0
      Top             =   960
      Width           =   4620
      Begin VB.Image Image1 
         Height          =   540
         Index           =   89
         Left            =   4000
         Top             =   4200
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   88
         Left            =   3500
         Top             =   4200
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   87
         Left            =   3000
         Top             =   4200
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   86
         Left            =   2500
         Top             =   4200
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   85
         Left            =   2000
         Top             =   4200
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   84
         Left            =   1500
         Top             =   4200
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   83
         Left            =   1000
         Top             =   4200
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   82
         Left            =   500
         Top             =   4200
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   81
         Left            =   0
         Top             =   4200
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   79
         Left            =   4000
         Top             =   3675
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   55
         Left            =   2000
         Top             =   2625
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   78
         Left            =   3500
         Top             =   3675
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   77
         Left            =   3000
         Top             =   3675
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   76
         Left            =   2500
         Top             =   3675
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   75
         Left            =   2000
         Top             =   3675
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   74
         Left            =   1500
         Top             =   3675
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   73
         Left            =   1000
         Top             =   3675
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   72
         Left            =   500
         Top             =   3675
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   71
         Left            =   0
         Top             =   3675
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   69
         Left            =   4000
         Top             =   3150
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   68
         Left            =   3500
         Top             =   3150
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   67
         Left            =   3000
         Top             =   3150
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   66
         Left            =   2500
         Top             =   3150
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   65
         Left            =   2000
         Top             =   3150
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   64
         Left            =   1500
         Top             =   3150
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   63
         Left            =   1000
         Top             =   3150
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   62
         Left            =   500
         Top             =   3150
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   61
         Left            =   0
         Top             =   3150
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   59
         Left            =   4000
         Top             =   2625
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   58
         Left            =   3500
         Top             =   2625
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   57
         Left            =   3000
         Top             =   2625
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   56
         Left            =   2500
         Top             =   2625
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   54
         Left            =   1500
         Top             =   2625
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   53
         Left            =   1000
         Top             =   2625
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   52
         Left            =   500
         Top             =   2625
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   51
         Left            =   0
         Top             =   2625
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   49
         Left            =   4000
         Top             =   2100
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   48
         Left            =   3500
         Top             =   2100
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   47
         Left            =   3000
         Top             =   2100
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   46
         Left            =   2500
         Top             =   2100
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   45
         Left            =   2000
         Top             =   2100
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   44
         Left            =   1500
         Top             =   2100
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   43
         Left            =   1000
         Top             =   2100
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   42
         Left            =   500
         Top             =   2100
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   41
         Left            =   0
         Top             =   2100
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   39
         Left            =   4000
         Top             =   1575
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   38
         Left            =   3500
         Top             =   1575
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   37
         Left            =   3000
         Top             =   1575
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   36
         Left            =   2500
         Top             =   1575
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   35
         Left            =   2000
         Top             =   1575
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   34
         Left            =   1500
         Top             =   1575
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   33
         Left            =   1000
         Top             =   1575
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   32
         Left            =   500
         Top             =   1575
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   31
         Left            =   0
         Top             =   1575
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   29
         Left            =   4000
         Top             =   1050
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   28
         Left            =   3500
         Top             =   1050
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   27
         Left            =   3000
         Top             =   1050
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   26
         Left            =   2500
         Top             =   1050
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   25
         Left            =   2000
         Top             =   1050
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   24
         Left            =   1500
         Top             =   1050
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   23
         Left            =   1000
         Top             =   1050
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   22
         Left            =   500
         Top             =   1050
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   21
         Left            =   0
         Top             =   1050
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   19
         Left            =   4000
         Top             =   525
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   18
         Left            =   3500
         Top             =   525
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   17
         Left            =   3000
         Top             =   525
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   16
         Left            =   2500
         Top             =   525
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   15
         Left            =   2000
         Top             =   525
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   14
         Left            =   1500
         Top             =   525
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   13
         Left            =   1000
         Top             =   525
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   12
         Left            =   500
         Top             =   525
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   11
         Left            =   0
         Top             =   525
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   9
         Left            =   4000
         Top             =   0
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   8
         Left            =   3500
         Top             =   0
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   7
         Left            =   3000
         Top             =   0
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   6
         Left            =   2500
         Top             =   0
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   5
         Left            =   2000
         Top             =   0
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   4
         Left            =   1500
         Top             =   0
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   3
         Left            =   1000
         Top             =   0
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   2
         Left            =   500
         Top             =   0
         Width           =   550
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   550
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   7425
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7752
            Text            =   "TEST YOUR SKILLS "
            TextSave        =   "TEST YOUR SKILLS "
            Object.ToolTipText     =   "BEST OF LUCK"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "SCORE"
            TextSave        =   "SCORE"
            Object.ToolTipText     =   "YOUR SCORE"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:52 PM"
            Object.ToolTipText     =   "TIME RIGHT NOW"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "1/19/2006"
            Object.ToolTipText     =   "TODAY DATE"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "time elapsed"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INFOMAZE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   4560
      TabIndex        =   13
      Top             =   120
      Width           =   2700
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   9360
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      URL             =   "F:\vb ball game\FINAL5\6.MID"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   100
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1085
      _cy             =   873
   End
   Begin VB.Label lblelapsed 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Time Elapsed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label vrsanimation 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "CYBERPARK , V.R.S.&&.Y.R.N.College,  CHIRALA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   6600
      Width           =   6705
   End
   Begin VB.Label lblscore 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label lblname 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "frmmainkid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time1 As Integer
Dim Score As Integer
Dim pause As Boolean
Dim imagearray(1 To 3) As Integer
Dim insertarray(1 To 3) As Integer
Dim boardarray(0 To 8, 1 To 9) As Integer
Dim selection(1 To 89) As Integer
Dim imagecolorindex(1 To 90) As Integer
Dim ball(1 To 7, 1 To 100) As Integer
Dim destination As Integer
Dim Source As Integer
Dim des1 As Boolean
Dim des2 As Boolean
Dim des3, des4, des5, des6 As Boolean
Dim count1 As Integer
Dim previewcount As Integer
Dim i6, i7 As Integer
Public Uname As String
Private Sub Form_Load()
Score = 0
time1 = 0
pause = False
   Timer1.Enabled = True
previewcount = 0
For i = 1 To 89
  If (Int(i Mod 10) <> 0) Then
  Image1.Item(i).Picture = LoadPicture(App.Path + "\empty.jpg")
  End If
Next
For i = 0 To 8
 For j = 1 To 9
   boardarray(i, j) = 0
 Next
Next
preview
Position
End Sub
Private Sub preview()
Dim i As Integer
Dim p As Integer
For i = 1 To 89
  If imagecolorindex(i) <> 0 Then
    p = p + 1
  End If
Next
ganeover = p
If p >= 80 Then
    MsgBox "Game Over", vbOKOnly
    Unload frmmainkid
End If
If p >= 71 Then
  previewcount = 1
  Picture3.Item(2).Visible = False
  Picture3.Item(3).Visible = False
ElseIf p >= 54 Then
  previewcount = 2
  Picture3.Item(3).Visible = False
  Picture3.Item(2).Visible = True
Else
  previewcount = 3
   Picture3.Item(3).Visible = True
   Picture3.Item(2).Visible = True
End If
Dim a As Integer
Randomize
i = 1
While (i <= previewcount)
a = Rnd(Time) * 10
If (a < 6 And a > 0) Then
  imagearray(i) = a
  i = i + 1
End If
If a = 6 Then
 i6 = i6 + 1
  If i6 = 1 Then
    imagearray(i) = 6
    i6 = 0
    i = i + 1
  End If
End If
If a = 7 Then
  i7 = i7 + 1
 If i7 = 1 Then
    imagearray(i) = 7
    i7 = 0
    i = i + 1
  End If
End If
Wend
For i = 1 To previewcount
  Picture3.Item(i).Picture = LoadPicture(App.Path + "/ball" & imagearray(i) & ".jpg")
Next
End Sub
Private Sub Position()
Dim a As Integer
Dim i As Integer
i = 1
Randomize
While (i <= previewcount)
a = Rnd(Time) * 100
If ((a <= 89) And ((a Mod 10) <> 0)) Then
 If (boardarray(Int(a / 10), a Mod 10) <> 1) Then
  insertarray(i) = a
  boardarray(Int(a / 10), a Mod 10) = 1
   i = i + 1
 End If
End If
Wend
For i = 1 To previewcount
  Image1.Item(insertarray(i)).Picture = LoadPicture(App.Path + "/ball" & imagearray(i) & ".jpg")
   imagecolorindex(insertarray(i)) = imagearray(i)
Next
des1 = False
des2 = False
Matchingh
Matchingv
Matchingd1
Matchingd2
preview
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim con As New ADODb.Connection
con.Provider = "Microsoft.Jet.OLEDB.4.0"
con.Open App.Path + "\db2.mdb"
Dim rs As New ADODb.Recordset
rs.Open "select *from table1 where score=(select min(score) from table1)", con, adOpenDynamic, adLockOptimistic
If (rs.Fields("score") < Score) Then
 rs.Fields("name").Value = Uname
 rs.Fields("score").Value = Score
 rs.Fields("timeelpsed").Value = lblelapsed.Caption
 rs.Update
End If
rs.Close
con.Close
Load frmkidstat
End Sub
Private Sub Image1_Click(Index As Integer)
Dim i, j As Integer
i = Int(Index / 10)
j = Index Mod 10
If boardarray(i, j) = 1 Then
  Source = Index
  des1 = True
  des2 = False
End If
If des1 = True Then
If boardarray(i, j) = 0 Then
  destination = Index
  des2 = True
End If
End If
If des1 = True And des2 = True Then
   prev = Source
   Generation
   des1 = False
des2 = False
End If
End Sub
Private Sub Generation()
Dim i As Integer
Dim left, right, bottom, top As Integer
left = 0
right = 0
bottom = 0
top = 0
For i = 1 To 89
  selection(i) = 0
Next
selection(1) = Source
count1 = 1
i = 1
While (selection(i) <> 0)
  If (selection(i) = destination) Then
    Image1.Item(destination).Picture = Image1.Item(Source).Picture
    Image1.Item(Source).Picture = LoadPicture(App.Path + "\empty.jpg")
    boardarray(Int(destination / 10), destination Mod 10) = 1
    boardarray(Int(Source / 10), Source Mod 10) = 0
    imagecolorindex(destination) = imagecolorindex(Source)
    imagecolorindex(Source) = 0
    des3 = False
    des4 = False
    des5 = False
    des6 = False
    Matchingh
    Matchingv
    Matchingd1
    Matchingd2
    If (des3 = False And des4 = False And des5 = False And des6 = False) Then
         Position
    End If
    Exit Sub
  Else
     left = selection(i) - 1
     right = selection(i) + 1
     bottom = selection(i) + 10
     top = selection(i) - 10
     If ((left Mod 10) = 0 Or (left < 1 Or left > 89)) Then
       left = 0
     End If
     If ((right Mod 10) = 0 Or (right < 1 Or right > 89)) Then
       right = 0
     End If
     If ((bottom Mod 10) = 0 Or (bottom < 1 Or bottom > 89)) Then
       bottom = 0
     End If
     If ((top Mod 10) = 0 Or (top < 1 Or top > 89)) Then
       top = 0
     End If
       For j = 1 To count1
          If (selection(j) = left) Then
            left = 0
          End If
       Next
       For j = 1 To count1
          If (selection(j) = right) Then
            right = 0
          End If
       Next
       For j = 1 To count1
          If (selection(j) = bottom) Then
            bottom = 0
          End If
       Next
       For j = 1 To count1
          If (selection(j) = top) Then
            top = 0
          End If
       Next
     If left <> 0 Then
       If boardarray(Int(left / 10), left Mod 10) <> 1 Then
         count1 = count1 + 1
         selection(count1) = left
       End If
     End If
     If right <> 0 Then
         If boardarray(Int(right / 10), right Mod 10) <> 1 Then
         count1 = count1 + 1
         selection(count1) = right
         End If
     End If
     
     If bottom <> 0 Then
         If boardarray(Int(bottom / 10), bottom Mod 10) <> 1 Then
         count1 = count1 + 1
         selection(count1) = bottom
         End If
     End If
     If top <> 0 Then
         If boardarray(Int(top / 10), top Mod 10) <> 1 Then
         count1 = count1 + 1
         selection(count1) = top
         End If
     End If
   End If
i = i + 1
Wend
End Sub
Private Sub Colorgeneration()
Dim i, j As Integer
For i = 1 To 7
  For j = 1 To 89
    ball(i, j) = 0
  Next
Next
 For l = 1 To 89
   If imagecolorindex(l) <> 0 Then
    ball(imagecolorindex(l), l) = 1
     If imagecolorindex(l) = 6 Or imagecolorindex(l) = 7 Then
       For k1 = 1 To 5
          ball(k1, l) = 1
       Next
     End If
   End If
  Next
End Sub
Private Sub Matchingv()
Colorgeneration
Dim i, j As Integer
Dim c, p1, k As Integer
  For i = 1 To 7
   For j = 1 To 89
      If ball(i, j) = 1 Then
        p1 = j
         k = 0
         While (ball(i, j + k) <> 0 And j + k <= 89)
            c = c + 1
            If ((j + k) / 10 <> 8) Then
            k = k + 10
            End If
         Wend
         If c > 4 Then
            count1 = c
            Scorecard (count1)
            lblscore.Caption = Score
            Disappearv (p1)
          End If
          c = 0
      End If
        c = 0
    Next
    c = 0
  Next
End Sub
Private Sub Matchingh()
Colorgeneration
Dim i, j As Integer
Dim c, p1, k As Integer
  For i = 1 To 7
   For j = 1 To 89
      If ball(i, j) = 1 Then
        p1 = j
         k = 0
         While (ball(i, j + k) <> 0 And j + k <= 89)
           c = c + 1
           If j + k <= 89 Then
           k = k + 1
           End If
         Wend
         If c > 4 Then
            count1 = c
            Scorecard (count1)
            lblscore.Caption = Score
           Timer1_Timer
             disappearh (p1)
          End If
          c = 0
        End If
        c = 0
    Next
    c = 0
  Next
End Sub
Private Sub Matchingd1()
Colorgeneration
Dim i, j As Integer
Dim c, p1, k As Integer
  For i = 1 To 7
   For j = 1 To 89
      If ball(i, j) = 1 Then
        p1 = j
         k = 0
         While (ball(i, j + k) <> 0 And j + k <= 89)
            c = c + 1
            If ((j + k) / 10 <> 8) Then
            k = k + 11
            End If
         Wend
         If c > 4 Then
            count1 = c
            Scorecard (count1)
            lblscore.Caption = Score
              Timer1_Timer
            Disappeard1 (p1)
          End If
          c = 0
      End If
        c = 0
    Next
    c = 0
  Next
End Sub
Private Sub Matchingd2()
Colorgeneration
Dim i, j As Integer
Dim c, p1, k As Integer
  For i = 1 To 7
   For j = 1 To 89
      If ball(i, j) = 1 Then
        p1 = j
         k = 0
         While (ball(i, j + k) <> 0 And j + k <= 85)
            c = c + 1
            If ((j + k) / 10 <> 8) Then
            k = k + 9
            End If
         Wend
         If c > 4 Then
            count1 = c
            Scorecard (count1)
            lblscore.Caption = Score
            Timer1_Timer
            Disappeard2 (p1)
          End If
          c = 0
      End If
        c = 0
    Next
    c = 0
  Next
End Sub
Private Sub Scorecard(s As Integer)
Dim prev As Integer
  prev = Score
Select Case s
 Case 5
    Score = Score + 3
 Case 6
    Score = Score + 5
 Case 7
    Score = Score + 8
 Case 8
    Score = Score + 12
 Case 9
    Score = Score + 17
End Select
StatusBar1.Panels.Item(2).Text = Score
For i = 1 To Int(Score / 50)
  If (prev < (50 * i)) And (Score >= (50 * i)) Then
     For j = 1 To i
       Placefball
     Next
     d22 = True
  End If
Next
End Sub
Private Sub disappearh(pos As Integer)
Dim t1(9) As Integer
Dim desb As Boolean
desb = False
For i1 = 0 To count1 - 1
  t1(i1) = imagecolorindex(pos + i1)
Next
Dim temp As Integer
For i1 = 0 To count1 - 1
 If t1(i1) <> 0 And t1(i1) < 6 Then
    temp = t1(i1)
 End If
 If t1(i1) = 7 Then
    desb = True
 End If
Next
If desb = True Then
 For k1 = 1 To 89
      If (k1 Mod 10) <> 0 Then
       If imagecolorindex(k1) = temp Then
         imagecolorindex(k1) = 0
          boardarray(Int(k1 / 10), k1 Mod 10) = 0
         Image1.Item(k1).Picture = LoadPicture(App.Path + "\empty.jpg")
       End If
      End If
     Next
End If
For i = 0 To count1 - 1
   boardarray(Int(pos / 10), (pos + i) Mod 10) = 0
   imagecolorindex(pos + i) = 0
   Image1.Item(pos + i).Picture = LoadPicture(App.Path + "\empty.jpg")
   Image1.Item(pos + i).Picture = LoadPicture(App.Path + "\empty.jpg")
Next
count1 = 0
For i = 1 To 7
  For j = 1 To 89
    ball(i, j) = 0
  Next
Next
des3 = True
End Sub
Private Sub Disappearv(pos As Integer)
Dim t1(9) As Integer
Dim desb As Boolean
desb = False
For i1 = 0 To count1 - 1
  t1(i1) = imagecolorindex(pos + (i1 * 10))
Next
Dim temp As Integer
For i1 = 0 To count1 - 1
 If t1(i1) <> 0 And t1(i1) < 6 Then
    temp = t1(i1)
 End If
 If t1(i1) = 7 Then
    desb = True
 End If
Next
If desb = True Then
 For k1 = 1 To 89
      If (k1 Mod 10) <> 0 Then
       If imagecolorindex(k1) = temp Then
         imagecolorindex(k1) = 0
          boardarray(Int(k1 / 10), k1 Mod 10) = 0
         Image1.Item(k1).Picture = LoadPicture(App.Path + "\empty.jpg")
       End If
      End If
  Next
 End If
For i = 0 To count1 - 1
   boardarray(Int(pos / 10) + i, pos Mod 10) = 0
   imagecolorindex(pos + (i * 10)) = 0
   Image1.Item(pos + (i * 10)).Picture = LoadPicture(App.Path + "\empty.jpg")
Next
count1 = 0
For i = 1 To 7
  For j = 1 To 89
    ball(i, j) = 0
  Next
Next
des4 = True
End Sub
Private Sub Disappeard1(pos As Integer)
Dim t1(9) As Integer
Dim desb As Boolean
desb = False
For i1 = 0 To count1 - 1
  t1(i1) = imagecolorindex(pos + (i1 * 11))
Next
Dim temp As Integer
For i1 = 0 To count1 - 1
 If t1(i1) <> 0 And t1(i1) < 6 Then
    temp = t1(i1)
 End If
 If t1(i1) = 7 Then
    desb = True
 End If
Next
If desb = True Then
 For k1 = 1 To 89
      If (k1 Mod 10) <> 0 Then
       If imagecolorindex(k1) = temp Then
         imagecolorindex(k1) = 0
          boardarray(Int(k1 / 10), k1 Mod 10) = 0
         Image1.Item(k1).Picture = LoadPicture(App.Path + "\empty.jpg")
       End If
      End If
     Next
End If
For i = 0 To count1 - 1
   boardarray(Int(pos / 10) + i, (pos + i) Mod 10) = 0
   imagecolorindex(pos + (i * 11)) = 0
   Image1.Item(pos + (i * 11)).Picture = LoadPicture(App.Path + "\empty.jpg")
Next
count1 = 0
For i = 1 To 7
  For j = 1 To 89
    ball(i, j) = 0
  Next
Next
des5 = True
End Sub
Private Sub Disappeard2(pos As Integer)
Dim t1(9) As Integer
Dim desb As Boolean
desb = False
For i1 = 0 To count1 - 1
  t1(i1) = imagecolorindex(pos + (i1 * 9))
Next
Dim temp As Integer
For i1 = 0 To count1 - 1
 If t1(i1) <> 0 And t1(i1) < 6 Then
    temp = t1(i1)
 End If
 If t1(i1) = 7 Then
    desb = True
 End If
Next
If desb = True Then
 For k1 = 1 To 89
      If (k1 Mod 10) <> 0 Then
       If imagecolorindex(k1) = temp Then
         imagecolorindex(k1) = 0
          boardarray(Int(k1 / 10), k1 Mod 10) = 0
         Image1.Item(k1).Picture = LoadPicture(App.Path + "\empty.jpg")
       End If
      End If
     Next
End If
For i = 0 To count1 - 1
   boardarray(Int(pos / 10) + i, (pos - i) Mod 10) = 0
   imagecolorindex(pos + (i * 9)) = 0
   Image1.Item(pos + (i * 9)).Picture = LoadPicture(App.Path + "\empty.jpg")
Next
count1 = 0
For i = 1 To 7
  For j = 1 To 89
    ball(i, j) = 0
  Next
Next
des6 = True
End Sub
Private Sub Timer1_Timer()
  Dim s As Integer, m As Integer, h As Integer
  Dim display As String
  display = ""
  If pause = False Then
     incre (1)
  End If
  h = Int(time1 / 3600)
  m = Int((time1 Mod 3600) / 60)
  s = time1 Mod 60
  If h < 10 Then
     display = display & "0" & h & ":"
  Else
     display = display & h & ":"
  End If
  If m < 10 Then
     display = display & "0" & m & ":"
  Else
     display = display & m & ":"
  End If
  If s < 10 Then
     display = display & "0" & s
  Else
     display = display & s
  End If
  lblelapsed.Caption = display
  StatusBar1.Panels.Item(5).Text = display
End Sub
Private Sub incre(a As Integer)
   time1 = time1 + a
End Sub

Private Sub Timer2_Timer()
If vrsanimation.left = 10600 Then
  vrsanimation.left = frmmainkid.left - vrsanimation.Width
End If
vrsanimation.left = vrsanimation.left + 10
End Sub
 Public Sub fresume()
  pause = False
  For i = 1 To 89
   If i Mod 10 <> 0 Then
     Image1(i).Enabled = True
   End If
  Next
End Sub
 Public Sub fpause()
 If pause <> True Then
   pause = True
  End If
 For i = 1 To 89
   If i Mod 10 <> 0 Then
     Image1(i).Enabled = False
   End If
Next
End Sub
Private Sub Placefball()
i = 0
Dim b As Integer
Randomize
a = Rnd(Time) * 10
If (a Mod 2 = 0) Then
  b = 6
Else
  b = 7
End If
i = 0
Randomize
Do
a = Rnd(Time) * 100
If ((a <= 89) And ((a Mod 10) <> 0)) Then
 If (boardarray(Int(a / 10), a Mod 10) <> 1) Then
    boardarray(Int(a / 10), a Mod 10) = 1
    imagecolorindex(a) = b
    Image1.Item(a).Picture = LoadPicture(App.Path + "/ball" & b & ".jpg")
   i = i + 1
 End If
End If
Loop While (i < 1)
End Sub
