VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame famwelcome 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   7575
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   735
         Left            =   840
         TabIndex        =   5
         Top             =   3360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   1296
         _Version        =   327682
         Appearance      =   1
         Max             =   240
      End
      Begin VB.Label labwelcome 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1455
         Index           =   1
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ICICI Bank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   69.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   1680
         Width           =   7935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   4200
         Width           =   6255
      End
      Begin VB.Label labwelcome 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   0
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5760
      Top             =   4200
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   0
      Top             =   840
      Width           =   9735
   End
   Begin VB.Image icicibanklogo 
      Height          =   825
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Top             =   0
      Width           =   11700
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

 

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

End Sub



Private Sub Timer1_Timer()
i = i + 1
ProgressBar1.Value = ProgressBar1.Value + 10
Select Case i
Case 1
Label5.Caption = "Loading Forms..."
Case 5
Label5.Caption = "Connecting Database..."
Case 12
Label5.Caption = "Preparing User Inteface..."
Case 17
Label5.Caption = "Checking Connectivity..."
Case 21
Label5.Caption = "Preparing Accounts Info..."
Case 23
Label5.Caption = "Preparations Complete!!!"
Timer1.Enabled = False

Unload Me
CreateObject("sapi.spvoice").SPEAK "WELCOME TO IC ICI BANK"
frmbackground.Show
frmbackground.Enabled = False
Form3.Show
'frmbackground.Show
'frmLogin.Show
End Select
End Sub
