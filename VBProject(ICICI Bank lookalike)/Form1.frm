VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   3720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      Height          =   855
      Left            =   2160
      TabIndex        =   1
      Top             =   2520
      Width           =   1695
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Timer1.Interval = Rnd * 300 + 20
ProgressBar1.Value = ProgressBar1.Value + 2
Label1.Caption = ProgressBar1.Value & "%"
If Label1.Caption = 100 & "%" Then
MsgBox "COMPLETED"
Unload Me
Form2.Show

End If

End Sub
