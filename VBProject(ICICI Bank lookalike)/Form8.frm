VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form8 
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "              INSERT CUSTOMER DETAILS"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11925
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&CANCEL"
      Height          =   495
      Left            =   6960
      TabIndex        =   41
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   39
      Top             =   5880
      Width           =   11655
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   495
         Left            =   2880
         TabIndex        =   40
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   5535
      Left            =   9360
      TabIndex        =   32
      Top             =   240
      Width           =   2415
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ADDRESS PROOF  "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   36
         Top             =   2520
         Width           =   2055
         Begin VB.ComboBox Combo2 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Form8.frx":0152
            Left            =   240
            List            =   "Form8.frx":015C
            TabIndex        =   38
            Text            =   "PENDING"
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "STATUS"
            Height          =   375
            Left            =   360
            TabIndex        =   37
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ID PROOF DOCUMENT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   2055
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Form8.frx":0174
            Left            =   240
            List            =   "Form8.frx":017E
            TabIndex        =   35
            Text            =   "PENDING"
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "STATUS"
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
            TabIndex        =   34
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ACCOUNT DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   5040
      TabIndex        =   20
      Top             =   240
      Width           =   4095
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form8.frx":0196
         Left            =   2040
         List            =   "Form8.frx":01A0
         TabIndex        =   44
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text12 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2040
         TabIndex        =   43
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         Enabled         =   0   'False
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   31
         Top             =   3840
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   2040
         TabIndex        =   30
         Top             =   2520
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   84082689
         CurrentDate     =   42083
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   2040
         TabIndex        =   29
         Top             =   3120
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   84082689
         CurrentDate     =   42083
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2040
         TabIndex        =   23
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   42
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCESS CODE"
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
         TabIndex        =   28
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "BALANCE"
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
         TabIndex        =   27
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "EXPIRY DATE"
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
         TabIndex        =   26
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT TYPE"
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
         Left            =   240
         TabIndex        =   25
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OPENED"
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
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT NUMBER"
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
         Left            =   120
         TabIndex        =   21
         Top             =   720
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
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   1800
         TabIndex        =   19
         Top             =   4800
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FEMALE"
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MALE"
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   3360
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   2760
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Format          =   84082689
         CurrentDate     =   42083
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "HOME PHONE"
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
         TabIndex        =   9
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "OFFICE PHONE"
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
         TabIndex        =   8
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "E-MAIL ADDRESS"
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
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF BIRTH"
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
         TabIndex        =   6
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "GENDER"
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
         Left            =   480
         TabIndex        =   5
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "HOME ADDRESS"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "OFFICE ADDRESS"
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
         Left            =   120
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "LAST NAME"
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
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "FIRST NAME"
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public str As String
Dim i As Integer
Dim ct As Integer
'Dim ser As String







Private Sub Combo3_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 0) Then
KeyAscii = 0
End If
End Sub

Private Sub Command2_Click()

    Unload Me
   
End Sub

Private Sub Command1_Click()
    If Text1.Text = "" Then
        MsgBox "Please enter the First Name.", vbExclamation, Title
        Text1.SetFocus
        
        
    ElseIf Text2 = "" Then
        MsgBox "Please enter the Last Name.", vbExclamation, Title
        Text1.SetFocus
        
    
    
    ElseIf Text3.Text = "" Then
        MsgBox "Please enter the Home Address.", vbExclamation, Title
        Text1.SetFocus
        
        
    ElseIf Text4.Text = "" Then
        MsgBox "Please enter the Office Address.", vbExclamation, Title
        Text1.SetFocus
        
    
    
    
    ElseIf DTPicker1.Value = Date Then
        MsgBox "Date of Birth can not be today, Kindly change it", vbExclamation, Title
        Text1.SetFocus
        
    
    
    ElseIf DTPicker1.Value = Date Then
        MsgBox "Date of Birth can not be in future", vbExclamation, Title
        Text1.SetFocus
        
     
    ElseIf Option1.Value = False And Option2.Value = False Then
        MsgBox "Please select the Gender.", vbExclamation, Title
       
        
      
    ElseIf Text5.Text = "" Then
        MsgBox "Please enter the Email Address.", vbExclamation, Title
        Text1.SetFocus
            
    
    
    ElseIf Text6.Text = "" Then
        MsgBox "Please enter the Home Number.", vbExclamation, Title
        Text1.SetFocus
        
        
    
    
       
    ElseIf Combo3.Text = "" Then
    MsgBox "Please Enter Account Type", vbExclamation, Title
    
    
    ElseIf Text10.Text = "" Then
        MsgBox "Please enter the Balance.", vbExclamation, Title
        Text1.SetFocus
        
       
    ElseIf DTPicker3.Value = Date Then
        MsgBox "Expiry Date can not be today, Kindly change it", vbExclamation, Title
        Text1.SetFocus


Exit Sub

'If InStr(Text10.Text, ".") > 1 Then
'MsgBox "enter correct value"
'End If

        Else
        Call INSERT
        
        
    End If
    
    
    
    End Sub
    
    Private Sub INSERT()
 
    Module1.opencon
    Dim i As String
    If Option1.Value = True Then
    i = Option1.Caption
    Else
    i = Option2.Caption
    End If
    
    Cn.Execute ("INSERT INTO CUSTOMERINSERT VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & DTPicker1.Value & "','" & i & "'," & Text6.Text & "," & Text7.Text & "," & Text8.Text & ",'" & DTPicker2.Value & "','" & Combo3.Text & "','" & Text10.Text & "','" & DTPicker3.Value & "','" & Text11.Text & "'," & Text12.Text & ",'" & Combo1.Text & "','" & Combo2.Text & "','" & Text5.Text & "')")
   
    MsgBox "INSERTED SUCCESSFULLY", vbExclamation, "SUCCESSFULL"
    CreateObject("SAPI.SPVOICE").SPEAK "INSERTED SUCCESSFULLY"
Unload Me
  
  End Sub
  
  
Private Sub Form_Load()
Text8.Enabled = False
    Text8.Text = Int(((Rnd * 9) * 1000000) + Rnd)
    Text11.Enabled = False
    Text11.Text = Int(((Rnd * 9) * 100000) + Rnd)
    Text12.Enabled = False
 
    Text12.Text = Int(((Rnd * 9) * 1000000) + Rnd)
    Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

End Sub



Private Sub Text1_LostFocus()
Text1.Text = Trim(Text1.Text)
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
Private Sub Text10_LostFocus()
Text10.Text = Trim(Text10.Text)
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If keyacsii < 48 And KeyAscii > 57 And KeyAscii <> 8 Then
Text10.Locked = True
Else
Text10.Locked = False
End If
Select Case KeyAscii
Case 48 To 57
Case 46
If InStr(Text10.Text, ".") > 0 Then
'Decimal point already there.
KeyAscii = 0
If Left(Text10.Text, 1) = "." Then
Text10 = "0."
End If
End If
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If keyacsii < 48 And KeyAscii > 57 Then
Text6.Locked = True
Else
Text6.Locked = False
End If


End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
If keyacsii < 48 And KeyAscii > 57 Then
Text7.Locked = True
Else
Text7.Locked = False
End If

End Sub

