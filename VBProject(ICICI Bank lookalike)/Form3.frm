VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "p"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form3.frx":0000
   Moveable        =   0   'False
   Picture         =   "Form3.frx":0152
   ScaleHeight     =   5985
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1560
      MouseIcon       =   "Form3.frx":1B71
      MousePointer    =   99  'Custom
      Picture         =   "Form3.frx":1CC3
      ScaleHeight     =   735
      ScaleWidth      =   2175
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      MouseIcon       =   "Form3.frx":73ED
      MousePointer    =   99  'Custom
      Picture         =   "Form3.frx":753F
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      ItemData        =   "Form3.frx":80C1
      Left            =   480
      List            =   "Form3.frx":80CB
      TabIndex        =   2
      Text            =   "STATUS"
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox TxtPassword 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox TxtUserName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   5175
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   2640
         MouseIcon       =   "Form3.frx":80F3
         MousePointer    =   99  'Custom
         Picture         =   "Form3.frx":8245
         ScaleHeight     =   735
         ScaleWidth      =   2175
         TabIndex        =   9
         Top             =   3000
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   360
         MouseIcon       =   "Form3.frx":F621
         MousePointer    =   99  'Custom
         Picture         =   "Form3.frx":F773
         ScaleHeight     =   855
         ScaleWidth      =   2175
         TabIndex        =   6
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label lbl 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         Index           =   1
         X1              =   -240
         X2              =   5160
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         Index           =   0
         X1              =   0
         X2              =   5280
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   " Login "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2085
         TabIndex        =   8
         Top             =   0
         Width           =   930
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      Height          =   5055
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   5415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 0) Then
KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
 Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
'Form3.Picture = LoadPicture(App.Path & "/untitle4.bmp")
 End Sub

Private Sub Picture1_Click()

If TxtUserName.Text = "" Then
MsgBox ">> Username is Empty.", vbInformation, "USERNAME EMPTY"
TxtUserName.SetFocus
Exit Sub

ElseIf TxtPassword.Text = "" Then
MsgBox ">> Password is Empty", vbExclamation, "NO PASSWORD"
TxtPassword.SetFocus
Exit Sub

Else
Call login
End If

End Sub
Private Sub login()
Module1.opencon
Dim rs As New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
    rs.Open "select * from login  where username='" & TxtUserName.Text & "' and PASSWORD ='" & TxtPassword.Text & "' and  STATUS='" & Combo1.Text & "'", Cn, adOpenStatic, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "USER UNIDENTIFIED", vbExclamation, "ERROR"
        Exit Sub
    Else
    If rs.Fields(2) = "SALES REPRESENTATIVE" Then
              Form4.Toolbar1.Buttons(3).Visible = False
           Form4.Toolbar1.Buttons(4).Visible = False
Form4.Toolbar1.Buttons(6).Visible = False
Form4.Toolbar1.Buttons(7).Visible = False
Form4.Toolbar1.Buttons(9).Visible = False
Form4.Toolbar1.Buttons(11).Visible = False
Form4.SEARCH1.Visible = False
Form4.EDIT2.Visible = False
Form4.DELETE2.Visible = False
Form14.Label3.Caption = Form3.lbl.Caption
           Form8.Label18.Caption = "REQUEST ID :"
          ' Form8.Command3.Visible = False
        End If
        If rs.Fields(2) = "BANK MANAGER" Then
            Form4.Toolbar1.Buttons(1).Visible = False

           Form4.Toolbar1.Buttons(8).Visible = False
      Form4.Toolbar1.Buttons(10).Visible = False
      'Form4.Toolbar1.Buttons(6).Visible = True
    Form8.Combo1.Enabled = True
    Form8.Combo2.Enabled = True
    Form8.Command1.Visible = False
    Form8.Command1.Enabled = False
Form8.Command1.Visible = False
    Form8.Command1.Enabled = False
'Form8.Command3.Visible = True
Form4.INSERT.Visible = False
Form14.Label3.Caption = Form3.lbl.Caption

'    Form8.Combo2.Enabled = True
 '   Form8.Combo2.Enabled = True
    Form8.Label18.Caption = "CUSTOMER ID :"
        End If
        TxtPassword.Text = ""
        Combo1.Text = "STATUS"
        TxtPassword.SetFocus
        Form4.Show
        
        'Unload Me
    End If

End Sub
Private Sub Picture2_Click()
End
End Sub

Private Sub Picture3_Click()
TxtPassword.Text = ""
TxtPassword.SetFocus
If i = 0 Then
TxtUserName = "Admin"
 Combo1.Text = "ADMIN"
 Combo1.Locked = True
Picture1.Visible = False
Picture3.Visible = False

Picture4.Visible = True
i = 1

ElseIf i = 1 Then
 If validuser("Admin", TxtPassword, Combo1) = True Then
       Label1.Caption = "NEW USER"
       Label3.Caption = "ADD USER"
       TxtUserName = ""
       TxtPassword = ""
       Combo1.Locked = False
       Combo1.Text = "STATUS"
       TxtUserName.SetFocus
       i = 2
 Else
   MsgBox "Access Denied", vbCritical, "Login Failed"
 End If
  
ElseIf i = 2 Then
If CHECKUSER(TxtUserName) = True Then
MsgBox "Username Already Exist! Please select any other Username.", vbInformation, "Username"
Else
If Combo1.Text = "STATUS" Then
MsgBox "Insert Appropriate Status.", vbInformation, "Status"
Else
If AddUser(TxtUserName, TxtPassword, Combo1) = True Then
Picture4.Visible = False
Picture1.Visible = True
Label1.Caption = "USER NAME"
TxtUserName = ""
TxtPassword = ""

Picture3.Visible = True
Label3.Caption = "Log-in"
End If
End If
End If
End If
'Unload Me
End Sub


Private Sub Picture4_Click()
If i = 0 Then
TxtUserName = "Admin"
 Combo1.Locked = True
'Picture1.Visible = False
'Picture3.Visible = False

'Picture4.Visible = True
i = 1

ElseIf i = 1 Then
 If validuser("Admin", TxtPassword, Combo1) = True Then
       Label1.Caption = "NEW USER"
       Label3.Caption = "ADD USER"
       TxtUserName = ""
       TxtPassword = ""
       Combo1.Locked = False
       Combo1.Text = "STATUS"
       TxtUserName.SetFocus
       i = 2
 Else
   MsgBox "Access Denied", vbCritical, "Login Failed"
 End If
  
ElseIf i = 2 Then
If CHECKUSER(TxtUserName) = True Then
MsgBox "Username Already Exist! Please select any other Username.", vbInformation, "Username"
Else
If Combo1.Text = "STATUS" Then
MsgBox "Insert Appropriate Status.", vbInformation, "Status"
Else
If AddUser(TxtUserName, TxtPassword, Combo1) = True Then
Picture4.Visible = False
Label1.Caption = "USER NAME"
TxtUserName = ""
TxtPassword = ""
MsgBox "New Account Created.", vbInformation, "Success"
CreateObject("sapi.spvoice").SPEAK "New Account Created Successfully"
Unload Me

Form3.Show
'Picture3.Visible = True
'Label3.Caption = "Log-in"
End If
End If
End If
End If

End Sub


Private Sub TxtPassword_Change()
lbl.Caption = TxtPassword.Text
End Sub

Private Sub TxtUserName_LostFocus()
TxtUserName.Text = Trim(TxtUserName.Text)
End Sub
