VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmhelp 
   Caption         =   "Help"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtgamerules 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7215
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   11175
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   1085
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Game Rules"
            Key             =   "rules"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "How to play"
            Key             =   "How to Play"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Game controls"
            Key             =   "controls"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String



Private Sub Form_Load()
s = "Is a single player, highscore-oriented 9x9 square boardgame . The aim is to align five or more balls of the same color into one horizontal, vertical or diagonal row making them disappear and scoring points. Initial board is with three balls   ( randomly placed)." & vbNewLine
s = s & "                          " & vbNewLine


s = s & "You score points by aligning five or more balls of the same color horizontally, vertically or diagonally.To move a ball, point to it and click, then point and click on the destination square. Three new balls appear every round (round means moving a ball from one position to another position) may be Joker ball or Dynamite bar . The position of the new balls is random, as is the color. The next colors can be seen in the upper right preview window. Aligning more than five balls scores more points. " & vbNewLine
s = s & "                          " & vbNewLine


s = s & "Rules:" & vbNewLine
s = s & "                          " & vbNewLine

s = s & "1.  if the path is availabel then only ball moves" & vbNewLine
s = s & "2.  The ball is moved to the empty position only." & vbNewLine
s = s & "3.  The ball is not permitted over one another." & vbNewLine & vbNewLine


s = s & "If you complete the game with highest socre you will be in the top ten players." & vbNewLine

s = s & "Try! " & vbNewLine


txtgamerules.Text = s
End Sub

Private Sub TabStrip1_Click()



If TabStrip1.SelectedItem.Index = 1 Then
''''
s = "Is a single player, highscore-oriented 9x9 square boardgame . The aim is to align five or more balls of the same color into one horizontal, vertical or diagonal row making them disappear and scoring points. Initial board is with three balls   ( randomly placed)." & vbNewLine
s = s & "                          " & vbNewLine


s = s & "You score points by aligning five or more balls of the same color horizontally, vertically or diagonally.To move a ball, point to it and click, then point and click on the destination square. Three new balls appear every round (round means moving a ball from one position to another position) may be Joker ball or Dynamite bar . The position of the new balls is random, as is the color. The next colors can be seen in the upper right preview window. Aligning more than five balls scores more points. " & vbNewLine
s = s & "                          " & vbNewLine


s = s & "Rules:" & vbNewLine
s = s & "                          " & vbNewLine

s = s & "1.  if the path is availabel then only ball moves" & vbNewLine
s = s & "2.  The ball is moved to the empty position only." & vbNewLine
s = s & "3.  The ball is not permitted over one another." & vbNewLine & vbNewLine


s = s & "If you complete the game with highest socre you will be in the top ten players." & vbNewLine

s = s & "Try! " & vbNewLine


txtgamerules.Text = s
''''


ElseIf TabStrip1.SelectedItem.Index = 2 Then
s = " "
txtgamerules.Text = s
s = s & "How to play" & vbNewLine & vbNewLine

s = s & "1.  Click on file menu, it opens popup menu." & vbNewLine & vbNewLine

s = s & "2.  Click on new menu, opens new game." & vbNewLine & vbNewLine

s = s & "3.  When you click on new time counting automatically starts." & vbNewLine & vbNewLine

s = s & "4.  Now click on the ball, which you want to move." & vbNewLine & vbNewLine

s = s & "5.  Now click on the empty position where you want to place the ball." & vbNewLine & vbNewLine

s = s & "6.  For pause the game click on controls, it show pause and resume." & vbNewLine & vbNewLine

s = s & "7.  Now click on pause then game is paused." & vbNewLine & vbNewLine

s = s & "8.  To resume the game click on resume." & vbNewLine & vbNewLine

txtgamerules.Text = s
ElseIf TabStrip1.SelectedItem.Index = 3 Then
s = " "
txtgamerules.Text = s
s = s & " Game Controls:-" & vbNewLine & vbNewLine



          
s = s & "Shot Cuts:-" & vbNewLine & vbNewLine
    
s = s & "New Game       - Ctrl+N or Alt+F+N" & vbNewLine & vbNewLine

s = s & "Statistics -Ctrl + s Or Alt + F + T" & vbNewLine & vbNewLine

s = s & "Quit -Ctrl + Q Or Alt + F + U" & vbNewLine & vbNewLine

s = s & "Pause -Ctrl + P Or Alt + C + P" & vbNewLine & vbNewLine
     
s = s & "Resume     - Ctrl+R or Alt+C+R" & vbNewLine & vbNewLine

s = s & "Help -Ctrl + F Or F1" & vbNewLine & vbNewLine

s = s & "About -Alt + A" & vbNewLine & vbNewLine
txtgamerules.Text = s

End If
End Sub
