VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "INFOMAZE"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufnew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu muufhi 
         Caption         =   "&Hiscores"
         Begin VB.Menu mnukid 
            Caption         =   "&kid"
         End
         Begin VB.Menu mnumed 
            Caption         =   "&Medium"
         End
         Begin VB.Menu mnuhar 
            Caption         =   "&Hard"
         End
      End
      Begin VB.Menu mnufline 
         Caption         =   "-"
      End
      Begin VB.Menu mnufquit 
         Caption         =   "Q&uit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnucontrols 
      Caption         =   "Controls"
      Begin VB.Menu mnucpause 
         Caption         =   "&Pause"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnucresume 
         Caption         =   "&Resume"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu mnuhelp 
         Caption         =   "Help"
      End
      Begin VB.Menu cllege 
         Caption         =   "about our college"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cllege_Click()
 frmcollege.Show
End Sub
Private Sub MDIForm_Load()
 Load frmempty
End Sub
Private Sub mnucpause_Click()
frmmain.fpause
frmmainhar.fpause
frmmainkid.fpause
End Sub
Private Sub mnucresume_Click()
 frmmain.fresume
 frmmainhar.fresume
 frmmainkid.fresume
End Sub
Private Sub mnufnew_Click()
 Unload frmempty
 Unload Me
 Load Form1
 End Sub
Private Sub mnufquit_Click()
 Unload Me
 Unload frmmain
 Unload frmmainkid
 Unload frmmainhar
 Unload MDIForm1
End Sub
Private Sub mnuhar_Click()
Me.Hide
Load frmharstat
End Sub
Private Sub mnuhelp_Click()
 frmhelp.Show
End Sub
Private Sub mnukid_Click()
Me.Hide
Load frmstat
End Sub
Private Sub mnumed_Click()
Me.Hide
Load frmkidstat
End Sub
