VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Mirror Drawer"
   ClientHeight    =   5550
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture16 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture15 
      BackColor       =   &H00800080&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00800000&
      Height          =   255
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00808000&
      Height          =   255
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   2040
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00008000&
      Height          =   255
      Left            =   2040
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00008080&
      Height          =   255
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00000080&
      Height          =   255
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   600
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000007&
      Height          =   255
      Left            =   600
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox ec 
      BackColor       =   &H80000007&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton Command2 
         Caption         =   "Draw on this side..."
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "... or this side!"
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   0
         Width           =   3310
      End
      Begin VB.Line mirrorline 
         Index           =   0
         Visible         =   0   'False
         X1              =   1080
         X2              =   1440
         Y1              =   3120
         Y2              =   2760
      End
      Begin VB.Line mainline 
         Index           =   0
         Visible         =   0   'False
         X1              =   3960
         X2              =   4320
         Y1              =   3480
         Y2              =   3240
      End
      Begin VB.Line Line1 
         X1              =   3240
         X2              =   3240
         Y1              =   0
         Y2              =   4560
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New                    "
         Shortcut        =   ^N
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public go As Boolean
Public lastX As Integer
Public lastY As Integer
Public go2 As Boolean
Public iIndex As Integer
Public Xchange As Integer
Public Color As String


Private Sub Form_Load()
Color = ec.BackColor

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal


End Sub

Private Sub mnuExit_Click()
End

End Sub

Private Sub mnuNew_Click()
msg = MsgBox("Are you sure?", vbYesNoCancel, "New")
If msg = 6 Then
For i = 1 To (mainline.Count - 1)
Unload mainline(i)
Unload mirrorline(i)
Next i



iIndex = 0
mirrorline(0).Visible = False
mainline(0).Visible = False
go2 = False
go = False
End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And go = False Then
  go = True
  go2 = True
  lastX = X
  lastY = Y
  mainline(iIndex).X1 = X
  mainline(iIndex).Y1 = Y
  mainline(iIndex).BorderColor = Color
  'Get distance from current point to middle line
  dist = X - Line1.X1
  'Set mirror line first point
  mirrorline(iIndex).X1 = X - (dist * 2)
  mirrorline(iIndex).Y1 = Y
  mirrorline(iIndex).BorderColor = Color
  Exit Sub
End If

If Button = 1 Then
  If go2 = True Then
  mainline(iIndex).X2 = X
  mainline(iIndex).Y2 = Y
  mainline(iIndex).BorderColor = Color
  'Get distance from current point to middle line
  dist = X - Line1.X1
  'Draw mirror line second point
  mirrorline(iIndex).X2 = X - (dist * 2)
  mirrorline(iIndex).Y2 = Y
  mirrorline(iIndex).BorderColor = Color
  lastX = X
  lastY = Y
  mainline(iIndex).Visible = True
  go2 = False
  Exit Sub
  End If
  'Increase index and make a new line
  iIndex = iIndex + 1
  Load mainline(iIndex)
  Load mirrorline(iIndex)
  'Get the difference from lastX and current X
  Xchange = X - lastX
  'Start the rest of the draw code
  With mainline(iIndex)
    .X1 = lastX
    .Y1 = lastY
    .X2 = X
    .Y2 = Y
    .Visible = True
    .BorderColor = Color
  End With
  
  'Get Distance from current point from the middle serperator
  dist = lastX - Line1.X1
  'Draw mirror image
  With mirrorline(iIndex)
    .X1 = lastX - (dist * 2)
    .Y1 = lastY
    .X2 = X - (dist * 2) - (Xchange * 2)
    .Y2 = Y
    .Visible = True
    .BorderColor = Color
  End With
  lastX = X
  lastY = Y
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
go = False
go2 = False

End Sub

Private Sub Picture10_Click()
ec.BackColor = Picture10.BackColor
SetColor Picture10.BackColor

End Sub

Private Sub Picture11_Click()
ec.BackColor = Picture11.BackColor
SetColor Picture11.BackColor

End Sub

Private Sub Picture12_Click()
ec.BackColor = Picture12.BackColor
SetColor Picture12.BackColor

End Sub

Private Sub Picture13_Click()
ec.BackColor = Picture13.BackColor
SetColor Picture13.BackColor

End Sub

Private Sub Picture14_Click()
ec.BackColor = Picture14.BackColor
SetColor Picture14.BackColor

End Sub

Private Sub Picture15_Click()
ec.BackColor = Picture15.BackColor
SetColor Picture15.BackColor

End Sub

Private Sub Picture16_Click()
ec.BackColor = Picture16.BackColor
SetColor Picture16.BackColor


End Sub

Private Sub Picture3_Click()
ec.BackColor = Picture3.BackColor
SetColor Picture3.BackColor

End Sub
Private Sub SetColor(mecolor As String)
Color = mecolor

End Sub

Private Sub Picture4_Click()
ec.BackColor = Picture4.BackColor
SetColor Picture4.BackColor

End Sub

Private Sub Picture5_Click()
ec.BackColor = Picture5.BackColor
SetColor Picture5.BackColor

End Sub

Private Sub Picture6_Click()
ec.BackColor = Picture6.BackColor
SetColor Picture6.BackColor

End Sub

Private Sub Picture7_Click()
ec.BackColor = Picture7.BackColor
SetColor Picture7.BackColor

End Sub

Private Sub Picture8_Click()
ec.BackColor = Picture8.BackColor
SetColor Picture8.BackColor

End Sub

Private Sub Picture9_Click()
ec.BackColor = Picture9.BackColor
SetColor Picture9.BackColor

End Sub
