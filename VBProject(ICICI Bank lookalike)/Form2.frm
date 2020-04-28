VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00000080&
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000080&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9060
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "ED&IT"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   48
      Top             =   8280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&SEARCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   47
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&SEARCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&SEARCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   44
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&SEARCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   480
      TabIndex        =   36
      Top             =   7920
      Width           =   9975
      Begin VB.CommandButton Command11 
         Caption         =   "&CANCEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   49
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Caption         =   "ED&IT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   46
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&DELETE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   43
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "D&ELETE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   38
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&UPDATE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&SEARCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Picture         =   "Form2.frx":0152
      TabIndex        =   35
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3720
      TabIndex        =   34
      Top             =   120
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BANK DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   6120
      TabIndex        =   17
      Top             =   840
      Width           =   4455
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Form2.frx":30194
         Left            =   2400
         List            =   "Form2.frx":3019E
         TabIndex        =   41
         Top             =   4200
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Form2.frx":301B5
         Left            =   2400
         List            =   "Form2.frx":301BF
         TabIndex        =   40
         Text            =   "PENDING"
         Top             =   5640
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Form2.frx":301D8
         Left            =   2400
         List            =   "Form2.frx":301E2
         TabIndex        =   39
         Text            =   "PENDING"
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   28
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   27
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox Text12 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   26
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   25
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   24
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS PROOF :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "ID DOCUMENT PROOF :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   5040
         Width           =   2295
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCESS CODE :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "EXPIRY DATE :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "BALANCE :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT TYPE :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OPENED :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT NUMBER : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PERSONAL DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   5415
      Begin VB.TextBox Text17 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   32
         Top             =   6120
         Width           =   3255
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   16
         Top             =   5400
         Width           =   3255
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   14
         Top             =   4680
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   3960
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   10
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   9
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   31
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF BIRTH :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "GENDER :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "OFFICE PHONE :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "HOME PHONE :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "OFFICE ADDRESS :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "HOME ADDRESS :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "LAST NAME :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "FIRST NAME :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   495
      Left            =   7920
      TabIndex        =   50
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "        ENTER REQUEST ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   33
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture2_Click()
End
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 0) Then
KeyAscii = 0
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 0) Then
KeyAscii = 0
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 0) Then
KeyAscii = 0
End If
End Sub

Private Sub Command1_Click()


'Combo1.Enabled = True
'Combo2.Enabled = True
'Text1.Enabled = True
'Text2.Enabled = True
'Text3.Enabled = True
'Text4.Enabled = True
'Text5.Enabled = True
'Text6.Enabled = True
'Text7.Enabled = True
'Text8.Enabled = True
'Text12.Enabled = True
'Combo3.Enabled = True
'Text17.Enabled = True

If Text18.Text = "" Then
MsgBox ">> PLEASE ENTER THE CUSTOMER REQUEST ID!", vbExclamation, "PLEASE ENTER REQUEST ID!"
Text18.SetFocus
Else


    Module1.opencon
Dim recCustomer1 As ADODB.Recordset
Set recCustomer1 = New Recordset
recCustomer1.Open "select * from CUSTOMERINSERT where REQUESTID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic

If recCustomer1.EOF Then
MsgBox "  >> Record corresponding to this Customer was not found !! ", vbExclamation, "Record absent !!"
Text18.SetFocus

Exit Sub
End If

'Command2.Enabled = True
Command3.Enabled = True


Set Text1.DataSource = recCustomer1
    Text1.DataField = "FIRSTNAME"
    'lbl_name.Caption =  + lbl_name.Caption
Set Text2.DataSource = recCustomer1
    Text2.DataField = "LASTNAME"
Set Text3.DataSource = recCustomer1
    Text3.DataField = "HOMEADDRESS"
    'lbl_col.Caption = lbl_col.Caption + " College"
Set Text4.DataSource = recCustomer1
    Text4.DataField = "OFFICEADDRESS"
    'lbl_strm.Caption = "Stream " + lbl_strm.Caption
Set Text17.DataSource = recCustomer1
    Text17.DataField = "EMAIL"
    Set Text5.DataSource = recCustomer1
    Text5.DataField = "HOMEPHONE"
Set Text6.DataSource = recCustomer1
    Text6.DataField = "OFFICEPHONE"
Set Text7.DataSource = recCustomer1
    Text7.DataField = "GENDER"
Set Text8.DataSource = recCustomer1
    Text8.DataField = "DOB"
Set Text9.DataSource = recCustomer1
    Text9.DataField = "ACCOUNTNUMBER"
Set Text10.DataSource = recCustomer1
    Text10.DataField = "DATEOPENED"
Set Combo3.DataSource = recCustomer1
    Combo3.DataField = "ACCOUNTTYPE"
Set Text12.DataSource = recCustomer1
    Text12.DataField = "BALANCE"
Set Text13.DataSource = recCustomer1
    Text13.DataField = "EXPIRYDATE"
Set Text14.DataSource = recCustomer1
    Text14.DataField = "ACCESSCODE"
Set Combo1.DataSource = recCustomer1
    Combo1.DataField = "IDPROOFSTATUS"
Set Combo2.DataSource = recCustomer1
    Combo2.DataField = "ADDPROOFSTATUS"
    
    
    End If
    

End Sub




Private Sub Command10_Click()
Dim recCustomerB As ADODB.Recordset
 Set recCustomerB = New ADODB.Recordset
 recCustomerB.Open "select * from CUSTOMERAPPROVED where CUSTOMERID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic
 If recCustomerB.EOF Then
 recCustomerB.AddNew
 Else: recCustomerB.Update
 End If
 recCustomerB!FirstName = Text1.Text
 recCustomerB!LastName = Text2.Text
 recCustomerB!HomeAddress = Text3.Text
 recCustomerB!OfficeAddress = Text4.Text
 recCustomerB!HomePhone = Text5.Text
 recCustomerB!OfficePhone = Text6.Text
 recCustomerB!Gender = Text7.Text
 recCustomerB!DOB = Text8.Text
 recCustomerB!Email = Text17.Text
 recCustomerB!AccountNumber = Text9.Text
 recCustomerB!DateOpened = Text10.Text
 recCustomerB!AccountType = Combo3.Text
 recCustomerB!Balance = Text12.Text
 recCustomerB!ExpiryDate = Text13.Text
 recCustomerB!IDPROOFSTATUS = Combo1.Text
 recCustomerB!ADDPROOFSTATUS = Combo2.Text

 Dim res As String
 res = MsgBox("Are You Sure You Want To Update?", vbYesNo, "Update")
 If res = vbYes Then
recCustomerB.Update
 MsgBox "UPDATED SUCCESSFULLY", vbExclamation, "SUCCESSFULL"
CreateObject("SAPI.SPVOICE").SPEAK "UPDATED SUCCESSFULLY"
Unload Me
Else
Unload Me
End If

End Sub

Private Sub Command11_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim recCustomer2 As ADODB.Recordset
 Set recCustomer2 = New ADODB.Recordset
 recCustomer2.Open "select * from CUSTOMERINSERT where REQUESTID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic
 If recCustomer2.EOF Then
 recCustomer2.AddNew
 Else: recCustomer2.Update
 End If
' recCustomer2!FirstName = Text1.Text
 'recCustomer2!LastName = Text2.Text
 'recCustomer2!HomeAddress = Text3.Text
 'recCustomer2!OfficeAddress = Text4.Text
 'recCustomer2!HomePhone = Text5.Text
 'recCustomer2!OfficePhone = Text6.Text
 'recCustomer2!Gender = Text7.Text
 'recCustomer2!DOB = Text8.Text
 'recCustomer2!Email = Text17.Text
 'recCustomer2!AccountNumber = Text9.Text
 'recCustomer2!DateOpened = Text10.Text
 'recCustomer2!AccountType = Combo3.Text
 'recCustomer2!Balance = Text12.Text
 'recCustomer2!ExpiryDate = Text13.Text
 recCustomer2!IDPROOFSTATUS = Combo1.Text
 recCustomer2!ADDPROOFSTATUS = Combo2.Text
 
 
 Dim res As String
 res = MsgBox("Are You Sure You Want To Update?", vbYesNo, "Update")
 If res = vbYes Then
   If Combo1.Text = "PENDING" Or Combo2.Text = "PENDING" Then
    MsgBox "PLEASE APPROVE OR REJECT THE ID AND ADDRESS PROOF  ", vbExclamation, "PLEASE VERIFY!"
    Else
    'If Combo1.Text = "APPROVED" And Combo2.Text = "APPROVED" Then
recCustomer2.Update
 MsgBox "UPDATED SUCCESSFULLY", vbExclamation, "SUCCESSFUL"

Module1.opencon
 Cn.Execute ("INSERT INTO CUSTOMERAPPROVED VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text8.Text & "','" & Text7.Text & "'," & Text5.Text & "," & Text6.Text & "," & Text9.Text & ",'" & Text10.Text & "','" & Combo3.Text & "','" & Text12.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Text18.Text & "','" & Combo1.Text & "','" & Combo2.Text & "','" & Text17.Text & "')")

'Dim A, B As String
'A = Combo1.Text
'B = Combo2.Text

Dim recCustomer4 As ADODB.Recordset
Set recCustomer4 = New ADODB.Recordset
 recCustomer4.Open "select * from CUSTOMERINSERT where REQUESTID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic
 If recCustomer4.EOF Then
 recCustomer4.AddNew
 Else: recCustomer4.DELETE

Unload Me
End If
End If
End If

If Combo1.Text = "REJECTED" Or Combo2.Text = "REJECTED" Then
Dim recCustomer3 As ADODB.Recordset
 Set recCustomer3 = New ADODB.Recordset
 recCustomer3.Open "select * from CUSTOMERINSERT where REQUESTID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic
 
' Dim res1 As String
 'res1 = MsgBox("Are You Sure You Want To Update?", vbYesNo, "Delete")
 'If res1 = vbYes Then
 If recCustomer3.EOF Then
 recCustomer3.AddNew
 Else: recCustomer3.DELETE
 End If
 
 MsgBox "UPDATED SUCCESSFULLY", vbExclamation
 CreateObject("SAPI.SPVOICE").SPEAK "UPDATED SUCCESSFULLY"
Unload Me

End If
' recCustomer5.Open "select * from CUSTOMERINSERT where IDPROOFSTATUS='" & A & "' AND ADDPROOFSTATUS='" & B & "'", Cn, adOpenStatic, adLockOptimistic
 'If recCustomer5.EOF Then
 'recCustomer5.AddNew
 'Else: recCustomer5.DELETE
 
 
'c = "CUSTOMER"
'Cn.Execute ("INSERT INTO login VALUES ('" & Text1.Text & "','" & Text14.Text & "','" & c & "')")
 

 
 'Unload Me
 'End If
 
 
 

 End Sub


Private Sub Command3_Click()
Dim recCustomer3 As ADODB.Recordset
 Set recCustomer3 = New ADODB.Recordset
 recCustomer3.Open "select * from CUSTOMERINSERT where REQUESTID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic
 
 Dim res As String
 res = MsgBox("Are You Sure You Want To Delete?", vbYesNo, "Delete")
 If res = vbYes Then
 If recCustomer3.EOF Then
 recCustomer3.AddNew
 Else: recCustomer3.DELETE
 End If
 
 MsgBox "DELETED SUCCESSFULLY", vbExclamation
 CreateObject("SAPI.SPVOICE").SPEAK "DELETED SUCCESSFULLY"
Unload Me
 End If

End Sub

Private Sub Command4_Click()
'Command2.Enabled = True
'Command3.Enabled = True
'Combo1.Enabled = True
'Combo2.Enabled = True
'Text1.Enabled = True
'Text2.Enabled = True
'Text3.Enabled = True
'Text4.Enabled = True
'Text5.Enabled = True
'Text6.Enabled = True
'Text7.Enabled = True
'Text8.Enabled = True
'Text12.Enabled = True
'Combo3.Enabled = True
'Text17.Enabled = True
If Text18.Text = "" Then
MsgBox ">> PLEASE ENTER THE CUSTOMER ID!", vbExclamation, "PLEASE ENTER REQUEST ID!"
Text18.SetFocus
Else
Module1.opencon
Dim recCustomer3 As ADODB.Recordset
Set recCustomer3 = New Recordset
recCustomer3.Open "select * from CUSTOMERAPPROVED where CUSTOMERID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic

If recCustomer3.EOF Then
MsgBox "  >> Record corresponding to this Customer was not found !! ", vbExclamation, "Record absent !!"
Text18.SetFocus
Exit Sub
End If




Command5.Enabled = True

Set Text1.DataSource = recCustomer3
    Text1.DataField = "FIRSTNAME"
    'lbl_name.Caption =  + lbl_name.Caption
Set Text2.DataSource = recCustomer3
    Text2.DataField = "LASTNAME"
Set Text3.DataSource = recCustomer3
    Text3.DataField = "HOMEADDRESS"
    'lbl_col.Caption = lbl_col.Caption + " College"
Set Text4.DataSource = recCustomer3
    Text4.DataField = "OFFICEADDRESS"
    'lbl_strm.Caption = "Stream " + lbl_strm.Caption
Set Text17.DataSource = recCustomer3
    Text17.DataField = "EMAIL"
    Set Text5.DataSource = recCustomer3
    Text5.DataField = "HOMEPHONE"
Set Text6.DataSource = recCustomer3
    Text6.DataField = "OFFICEPHONE"
Set Text7.DataSource = recCustomer3
    Text7.DataField = "GENDER"
Set Text8.DataSource = recCustomer3
    Text8.DataField = "DOB"
Set Text9.DataSource = recCustomer3
    Text9.DataField = "ACCOUNTNUMBER"
Set Text10.DataSource = recCustomer3
    Text10.DataField = "DATEOPENED"
Set Combo3.DataSource = recCustomer3
    Combo3.DataField = "ACCOUNTTYPE"
Set Text12.DataSource = recCustomer3
    Text12.DataField = "BALANCE"
Set Text13.DataSource = recCustomer3
    Text13.DataField = "EXPIRYDATE"
Set Text14.DataSource = recCustomer3
    Text14.DataField = "ACCESSCODE"
Set Combo1.DataSource = recCustomer3
    Combo1.DataField = "IDPROOFSTATUS"
Set Combo2.DataSource = recCustomer3
    Combo2.DataField = "ADDPROOFSTATUS"
    
    
    
    
End If


End Sub

Private Sub Command5_Click()
Dim recCustomerC As ADODB.Recordset
 Set recCustomerC = New ADODB.Recordset
 recCustomerC.Open "select * from CUSTOMERAPPROVED where CUSTOMERID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic
 
 Dim res As String
 res = MsgBox("Are You Sure You Want To Delete?", vbYesNo, "Delete")
 If res = vbYes Then
If recCustomerC.EOF Then
 recCustomerC.AddNew
 Else: recCustomerC.DELETE
  End If

Dim recCustomer1 As ADODB.Recordset
 Set recCustomer1 = New ADODB.Recordset
 recCustomer1.Open "select * from login where PASSWORD='" & Text14.Text & "'", Cn, adOpenStatic, adLockOptimistic
 If recCustomer1.EOF Then
 recCustomer1.AddNew
 Else: recCustomer1.DELETE
 End If
 
 


 
 
 
 
 MsgBox "DELETED SUCCESSFULLY", vbExclamation, "SUCCESSFULL"
 CreateObject("SAPI.SPVOICE").SPEAK "DELETED SUCCESSFULLY"

 End If
Unload Me

End Sub


Private Sub Command6_Click()

If Text18.Text = "" Then
MsgBox ">> PLEASE ENTER THE CUSTOMER REQUEST ID!", vbExclamation, "PLEASE ENTER REQUEST ID!"
Text18.SetFocus
Else
Module1.opencon
Dim recCustomer4 As ADODB.Recordset
Set recCustomer4 = New Recordset
recCustomer4.Open "select * from CUSTOMERINSERT where REQUESTID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic

If recCustomer4.EOF Then
MsgBox "  >> Record corresponding to this Customer was not found !! ", vbExclamation, "Record absent !!"
Text18.SetFocus
Exit Sub
End If



Command2.Enabled = True
Command3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
'Text1.Enabled = True
'Text2.Enabled = True
'Text3.Enabled = True
'Text4.Enabled = True
'Text5.Enabled = True
'Text6.Enabled = True
'Text7.Enabled = True
'Text8.Enabled = True
'Text12.Enabled = True
'Combo3.Enabled = True
'Text17.Enabled = True



Set Text1.DataSource = recCustomer4
    Text1.DataField = "FIRSTNAME"
    'lbl_name.Caption =  + lbl_name.Caption
Set Text2.DataSource = recCustomer4
    Text2.DataField = "LASTNAME"
Set Text3.DataSource = recCustomer4
    Text3.DataField = "HOMEADDRESS"
    'lbl_col.Caption = lbl_col.Caption + " College"
Set Text4.DataSource = recCustomer4
    Text4.DataField = "OFFICEADDRESS"
    'lbl_strm.Caption = "Stream " + lbl_strm.Caption
Set Text17.DataSource = recCustomer4
    Text17.DataField = "EMAIL"
    Set Text5.DataSource = recCustomer4
    Text5.DataField = "HOMEPHONE"
Set Text6.DataSource = recCustomer4
    Text6.DataField = "OFFICEPHONE"
Set Text7.DataSource = recCustomer4
    Text7.DataField = "GENDER"
Set Text8.DataSource = recCustomer4
    Text8.DataField = "DOB"
Set Text9.DataSource = recCustomer4
    Text9.DataField = "ACCOUNTNUMBER"
Set Text10.DataSource = recCustomer4
    Text10.DataField = "DATEOPENED"
Set Combo3.DataSource = recCustomer4
    Combo3.DataField = "ACCOUNTTYPE"
Set Text12.DataSource = recCustomer4
    Text12.DataField = "BALANCE"
Set Text13.DataSource = recCustomer4
    Text13.DataField = "EXPIRYDATE"
Set Text14.DataSource = recCustomer4
    Text14.DataField = "ACCESSCODE"
Set Combo1.DataSource = recCustomer4
    Combo1.DataField = "IDPROOFSTATUS"
Set Combo2.DataSource = recCustomer4
    Combo2.DataField = "ADDPROOFSTATUS"
    End If
    
    
    




End Sub

Private Sub Command7_Click()
If Text18.Text = "" Then
MsgBox ">> PLEASE ENTER THE CUSTOMER REQUEST ID!", vbExclamation, "PLEASE ENTER REQUEST ID!"
Text18.SetFocus
Else

Module1.opencon
Dim recCustomer2 As ADODB.Recordset
Set recCustomer2 = New Recordset
recCustomer2.Open "select * from CUSTOMERINSERT where REQUESTID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic
If recCustomer2.EOF Then
MsgBox "  >> Record corresponding to this Customer was not found !! ", vbExclamation, "Record absent !!"
Text18.SetFocus
Exit Sub
End If



Command8.Enabled = True
'Command3.Enabled = True
'Combo1.Enabled = True
'Combo2.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text12.Enabled = True
Combo3.Enabled = True
Text17.Enabled = True



Set Text1.DataSource = recCustomer2
    Text1.DataField = "FIRSTNAME"
    'lbl_name.Caption =  + lbl_name.Caption
Set Text2.DataSource = recCustomer2
    Text2.DataField = "LASTNAME"
Set Text3.DataSource = recCustomer2
    Text3.DataField = "HOMEADDRESS"
    'lbl_col.Caption = lbl_col.Caption + " College"
Set Text4.DataSource = recCustomer2
    Text4.DataField = "OFFICEADDRESS"
    'lbl_strm.Caption = "Stream " + lbl_strm.Caption
Set Text17.DataSource = recCustomer2
    Text17.DataField = "EMAIL"
    Set Text5.DataSource = recCustomer2
    Text5.DataField = "HOMEPHONE"
Set Text6.DataSource = recCustomer2
    Text6.DataField = "OFFICEPHONE"
Set Text7.DataSource = recCustomer2
    Text7.DataField = "GENDER"
Set Text8.DataSource = recCustomer2
    Text8.DataField = "DOB"
Set Text9.DataSource = recCustomer2
    Text9.DataField = "ACCOUNTNUMBER"
Set Text10.DataSource = recCustomer2
    Text10.DataField = "DATEOPENED"
Set Combo3.DataSource = recCustomer2
    Combo3.DataField = "ACCOUNTTYPE"
Set Text12.DataSource = recCustomer2
    Text12.DataField = "BALANCE"
Set Text13.DataSource = recCustomer2
    Text13.DataField = "EXPIRYDATE"
Set Text14.DataSource = recCustomer2
    Text14.DataField = "ACCESSCODE"
Set Combo1.DataSource = recCustomer2
    Combo1.DataField = "IDPROOFSTATUS"
Set Combo2.DataSource = recCustomer2
    Combo2.DataField = "ADDPROOFSTATUS"
    
    
    
    
End If

End Sub

Private Sub Command8_Click()
Dim recCustomerA As ADODB.Recordset
 Set recCustomerA = New ADODB.Recordset
 recCustomerA.Open "select * from CUSTOMERINSERT where REQUESTID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic
 If recCustomerA.EOF Then
 recCustomerA.AddNew
 Else: recCustomerA.Update
 End If
 recCustomerA!FirstName = Text1.Text
 recCustomerA!LastName = Text2.Text
 recCustomerA!HomeAddress = Text3.Text
 recCustomerA!OfficeAddress = Text4.Text
 recCustomerA!HomePhone = Text5.Text
 recCustomerA!OfficePhone = Text6.Text
 recCustomerA!Gender = Text7.Text
 recCustomerA!DOB = Text8.Text
 recCustomerA!Email = Text17.Text
 recCustomerA!AccountNumber = Text9.Text
 recCustomerA!DateOpened = Text10.Text
 recCustomerA!AccountType = Combo3.Text
 recCustomerA!Balance = Text12.Text
 recCustomerA!ExpiryDate = Text13.Text
 recCustomerA!IDPROOFSTATUS = Combo1.Text
 recCustomerA!ADDPROOFSTATUS = Combo2.Text
 
 
 Dim res As String
 res = MsgBox("Are You Sure You Want To Update?", vbYesNo, "Update")
 If res = vbYes Then
    recCustomerA.Update
 MsgBox "UPDATED SUCCESSFULLY", vbExclamation, "SUCCESSFULL"
CreateObject("SAPI.SPVOICE").SPEAK "UPDATED SUCCESSFULLY"
Unload Me
Else
Unload Me
End If

End Sub

Private Sub Command9_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text12.Enabled = True
Combo3.Enabled = True
Text17.Enabled = True
Command10.Enabled = True

If Text18.Text = "" Then
MsgBox ">> PLEASE ENTER THE CUSTOMER ID!", vbExclamation, "PLEASE ENTER REQUEST ID!"
Text18.SetFocus
Else
Module1.opencon
Dim recCustomer5 As ADODB.Recordset
Set recCustomer5 = New Recordset
recCustomer5.Open "select * from CUSTOMERAPPROVED where CUSTOMERID=" & Text18.Text & "", Cn, adOpenStatic, adLockOptimistic

If recCustomer5.EOF Then
MsgBox "  >> Record corresponding to this Customer was not found !! ", vbExclamation, "Record absent !!"
Text18.SetFocus
Exit Sub
End If





Set Text1.DataSource = recCustomer5
    Text1.DataField = "FIRSTNAME"
    'lbl_name.Caption =  + lbl_name.Caption
Set Text2.DataSource = recCustomer5
    Text2.DataField = "LASTNAME"
Set Text3.DataSource = recCustomer5
    Text3.DataField = "HOMEADDRESS"
    'lbl_col.Caption = lbl_col.Caption + " College"
Set Text4.DataSource = recCustomer5
    Text4.DataField = "OFFICEADDRESS"
    'lbl_strm.Caption = "Stream " + lbl_strm.Caption
Set Text17.DataSource = recCustomer5
    Text17.DataField = "EMAIL"
    Set Text5.DataSource = recCustomer5
    Text5.DataField = "HOMEPHONE"
Set Text6.DataSource = recCustomer5
    Text6.DataField = "OFFICEPHONE"
Set Text7.DataSource = recCustomer5
    Text7.DataField = "GENDER"
Set Text8.DataSource = recCustomer5
    Text8.DataField = "DOB"
Set Text9.DataSource = recCustomer5
    Text9.DataField = "ACCOUNTNUMBER"
Set Text10.DataSource = recCustomer5
    Text10.DataField = "DATEOPENED"
Set Combo3.DataSource = recCustomer5
    Combo3.DataField = "ACCOUNTTYPE"
Set Text12.DataSource = recCustomer5
    Text12.DataField = "BALANCE"
Set Text13.DataSource = recCustomer5
    Text13.DataField = "EXPIRYDATE"
Set Text14.DataSource = recCustomer5
    Text14.DataField = "ACCESSCODE"
Set Combo1.DataSource = recCustomer5
    Combo1.DataField = "IDPROOFSTATUS"
Set Combo2.DataSource = recCustomer5
    Combo2.DataField = "ADDPROOFSTATUS"
    
End If
End Sub

Private Sub Form_Load()


Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2


End Sub




Private Sub Text17_LostFocus()
Text17.Text = Trim(Text17.Text)
End Sub

Private Sub Text1_LostFocus()
Text1.Text = Trim(Text1.Text)
End Sub

Private Sub Text18_Change()
Text18.Text = Trim(Text18.Text)
End Sub

Private Sub Text18_Click()
Text18.Text = ""
End Sub

Private Sub Text2_LostFocus()
Text2.Text = Trim(Text2.Text)
End Sub
Private Sub Text3_LostFocus()
Text3.Text = Trim(Text3.Text)
End Sub
Private Sub Text4_LostFocus()
Text4.Text = Trim(Text4.Text)
End Sub
Private Sub Text5_LostFocus()
Text5.Text = Trim(Text5.Text)
End Sub
Private Sub Text6_LostFocus()
Text6.Text = Trim(Text6.Text)
End Sub
Private Sub Text7_LostFocus()
Text7.Text = Trim(Text7.Text)
End Sub
Private Sub Text8_LostFocus()
Text8.Text = Trim(Text8.Text)
End Sub
Private Sub Text9_LostFocus()
Text9.Text = Trim(Text9.Text)
End Sub
Private Sub Text10_LostFocus()
Text10.Text = Trim(Text10.Text)
End Sub
Private Sub Text12_LostFocus()
Text12.Text = Trim(Text12.Text)
End Sub

Private Sub Text13_LostFocus()
Text13.Text = Trim(Text13.Text)
End Sub

Private Sub Text14_LostFocus()
Text14.Text = Trim(Text14.Text)
End Sub
