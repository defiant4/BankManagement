VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   5415
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   7080
         Top             =   360
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1920
         Picture         =   "Form5.frx":98FC2
         ScaleHeight     =   375
         ScaleWidth      =   135
         TabIndex        =   1
         Top             =   3240
         Width           =   135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub


Private Sub Timer1_Timer()
Picture3.Width = Picture3.Width + Rnd * 100

If Picture3.Width >= 5175 Then
Unload Me
CreateObject("sapi.spvoice").SPEAK "WELCOME TO ZIMBRA BANK"


'Form3.Show
End If
End Sub
