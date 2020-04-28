VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H000040C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CHANGE PASSWORD"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6420
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHANGE PASSWO&RD"
      Height          =   495
      Left            =   2183
      TabIndex        =   4
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox Text2 
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
      IMEMode         =   3  'DISABLE
      Left            =   1463
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox Text1 
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
      IMEMode         =   3  'DISABLE
      Left            =   1463
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "4444"
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Image icicibanklogo 
      Height          =   825
      Index           =   0
      Left            =   0
      Picture         =   "Form14.frx":0152
      Top             =   0
      Width           =   11700
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER NEW PASSWORD "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2085
      TabIndex        =   2
      Top             =   2640
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER OLD PASSWORD "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2010
      TabIndex        =   0
      Top             =   1680
      Width           =   2145
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Enter Old Password"
ElseIf Text2.Text = "" Then
MsgBox "Enter New Password"
ElseIf Text1.Text = Text2.Text Then
MsgBox "New Password Is Same"
Exit Sub
Else
If Label3.Caption = Text1.Text Then
Dim res As String
Module1.opencon
Dim CP As New ADODB.Recordset
CP.Open "select * from LOGIN where PASSWORD='" & Text1.Text & "'", Cn, adOpenStatic, adLockOptimistic
If CP.EOF Then
 CP.AddNew
 Else: CP.Update
 End If
 CP!password = Text2.Text
 res = MsgBox("Are You Sure You Want To CHANGE PASSWORD?", vbYesNo, "Update")
 If res = vbYes Then
    CP.Update
 MsgBox "UPDATED SUCCESSFULLY", vbExclamation, "SUCCESSFULL"
CreateObject("SAPI.SPVOICE").SPEAK "PASSWORD UPDATED SUCCESSFULLY"
MsgBox " Reset system", vbOKOnly, "RECOMMENDED"
End
Else
Unload Me
End If
Else
MsgBox "Invalid Old Password!", vbExclamation, "Error"


End If
 End If
End Sub



Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Label3.Caption = Form3.lbl.Caption
End Sub

