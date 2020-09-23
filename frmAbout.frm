VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About Mirror Drawer"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   2520
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ya, Whatever"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Also Check out Auto-Mouse (Mouser 2.0). It really controls your mouse, and really drives people crazy!"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Please rate me a 5! "
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":0000
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub
