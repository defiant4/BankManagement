VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form11 
   BackColor       =   &H00000080&
   Caption         =   "Form11"
   ClientHeight    =   9795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10920
   LinkTopic       =   "Form11"
   ScaleHeight     =   9795
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&CANCEL"
      Height          =   615
      Left            =   6600
      TabIndex        =   33
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      TabIndex        =   25
      Top             =   7080
      Width           =   9735
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   615
         Left            =   2760
         TabIndex        =   26
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   6615
      Left            =   6000
      TabIndex        =   18
      Top             =   240
      Width           =   4335
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   32
         Top             =   5640
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   495
         Left            =   2400
         TabIndex        =   31
         Top             =   4680
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92798977
         CurrentDate     =   42102
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form11.frx":0000
         Left            =   2400
         List            =   "Form11.frx":0007
         TabIndex        =   30
         Text            =   "APPROVED"
         Top             =   3720
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   2400
         TabIndex        =   29
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         Format          =   92798977
         CurrentDate     =   42102
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1680
         TabIndex        =   28
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1680
         TabIndex        =   27
         Top             =   480
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4320
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label14 
         Caption         =   "SALARY"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   5760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "JOINING DATE :"
         Height          =   735
         Left            =   240
         TabIndex        =   23
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "QUALIFICATION STATUS :"
         Height          =   615
         Left            =   240
         TabIndex        =   22
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "INTERVIEW DATE :"
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "ACCESS CODE :"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "ID NUMBER :"
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   6615
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   1920
         TabIndex        =   17
         Top             =   5880
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   1920
         TabIndex        =   16
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1920
         TabIndex        =   15
         Top             =   4440
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   1920
         TabIndex        =   14
         Top             =   3600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         _Version        =   393216
         Format          =   92798977
         CurrentDate     =   42102
      End
      Begin VB.OptionButton Option2 
         Caption         =   "FEMALE"
         Height          =   735
         Left            =   3240
         TabIndex        =   13
         Top             =   2760
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MALE"
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1920
         TabIndex        =   11
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1920
         TabIndex        =   10
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1920
         TabIndex        =   9
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label8 
         Caption         =   "EMAIL :"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "OFFICE PHONE :"
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "HOME PHONE :"
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "DATE OF BIRTH :"
         Height          =   735
         Left            =   360
         TabIndex        =   5
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "GENDER :"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "HOME ADDRESS :"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "LAST NAME :"
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "FIRST NAME :"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
        MsgBox "Please enter the Home Phone.", vbExclamation, Title
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
        MsgBox "Please enter the Office Phone.", vbExclamation, Title
        Text1.SetFocus
            
    
    
    ElseIf Text6.Text = "" Then
        MsgBox "Please enter the Email Address.", vbExclamation, Title
        Text1.SetFocus
        
        
   ElseIf Combo2.Text = "" Then
MsgBox "Please Enter Qualification Status", vbExclamation, Title
 
    
    
    
Exit Sub
        Else
        Call INSERT
        
        
    End If
    End Sub

Private Sub INSERT()
 Module1.opencon
    Dim i2 As String
    If Option1.Value = True Then
    i2 = Option1.Caption
    Else
    i2 = Option2.Caption
    End If
    
    Cn.Execute ("INSERT INTO SRINSERT VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & i2 & "','" & DTPicker1.Value & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & DTPicker2.Value & "','" & Combo2.Text & "')")
   'If Combo2.Text = "" Then
   'MsgBox "PLEASE APPROVE OR REJECT THE STATUS!", vbExclamation, "PLAESE VERIFY"
    
    MsgBox "INSERTED SUCCESSFULLY", vbExclamation
    CreateObject("SAPI.SPVOICE").SPEAK "INSERTED SUCCESSFULLY"
    Unload Me
   
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

Module1.opencon
Dim SR1 As ADODB.Recordset
Set SR1 = New Recordset
SR1.Open "select * from SRINSERT where IDNUMBER='" & Text10.Text & "'", Cn, adOpenStatic, adLockOptimistic

If Text10.Text = "" Then
MsgBox ">> PLEASE ENTER THE ID NUMBER!", vbExclamation, "PLEASE ENTER ID!"

ElseIf SR1.EOF Then
MsgBox "  >> Record corresponding to this SALES REPRESENTATIVE was not found !! ", vbExclamation, "Record absent !!"
Exit Sub
End If

Set Text1.DataSource = SR1
    Text1.DataField = "FIRSTNAME"
    'lbl_name.Caption =  + lbl_name.Caption
Set Text2.DataSource = SR1
    Text2.DataField = "LASTNAME"
Set Text3.DataSource = SR1
    Text3.DataField = "HOMEADDRESS"
    'lbl_col.Caption = lbl_col.Caption + " College"
'Set Text4.DataSource = recCustomer
 '   Text4.DataField = "OFFICEADDRESS"
    'lbl_strm.Caption = "Stream " + lbl_strm.Caption
Set Text6.DataSource = recCustomer
    Text6.DataField = "EMAIL"
    Set Text4.DataSource = SR1
    Text4.DataField = "HOMEPHONE"
Set Text5.DataSource = SR1
    Text5.DataField = "OFFICEPHONE"
Set Text11.DataSource = SR1
    Text11.DataField = "GENDER"
Set DTPicker1.DataSource = SR1
    DTPicker1.DataField = "DOB"
Set Text7.DataSource = SR1
    Text7.DataField = "IDNUMBER"
Set DTPicker2.DataSource = SR1
    DTPicker2.DataField = "INTERVIEWDATE"
'Set Combo3.DataSource = recCustomer
 '   Combo3.DataField = "ACCOUNTTYPE"
Set Text8.DataSource = SR1
    Text8.DataField = "ACCESSCODE"
'Set Text13.DataSource = recCustomer
 '   Text13.DataField = "EXPIRYDATE"
'Set Text14.DataSource = recCustomer
 '   Text14.DataField = "ACCESSCODE"
'Set Combo1.DataSource = recCustomer
 '   Combo1.DataField = "IDPROOFSTATUS"
Set Combo2.DataSource = SR1
    Combo2.DataField = "QUALIFICATIONSTATUS"
    


End Sub

Private Sub Command4_Click()
'DTPicker1.Visible = False
'DTPicker2.Visible = False

End Sub
If Combo1.Text = "APPROVED" Then
Text9.Enabled = True
DTPicker3.Enabled = True
MsgBox "PLEASE ENTER THE REMAINING DETAILS!", vbExclamation, "PLEASE VERIFY"
    If DTPicker3.Value = Date Then
        MsgBox "Joining date can not be today, Kindly change it", vbExclamation, Title
        Text1.SetFocus

    ElseIf Text9.Text = "" Then
        MsgBox "Please enter the Home Number.", vbExclamation, Title
        Text1.SetFocus
        
Exit Sub
Else
Call update

End If

ElseIf Combo1.Text = "REJECTED" Then

Private Sub update()
Dim SR As ADODB.Recordset
 Set SR = New ADODB.Recordset
 SR.Open "select * from SRINSERT where IDNUMBER='" & Text10.Text & "'", Cn, adOpenStatic, adLockOptimistic
 If SR.EOF Then
SR.AddNew
 Else: SR.update
 End If
 SR!FirstName = Text1.Text
 SR!LastName = Text2.Text
 SR!HomeAddress = Text3.Text
 SR!Gender = Text11.Text
 SR!DOB = DTPicker1.Value
 SR!HomePhone = Text4.Text
SR!OfficePhone = Text5.Text
 SR!Email = Text6.Text
 SR!IDNUMBERNumber = Text7.Text
 SR!AccessCode = Text8.Text
 SR!INTERVIEWDATE = DTPicker2.Value
 SR!QUALIFICATIONSTATUS = Combo2.Text
 SR!JOININGDATE = DTPicker3.Value
 SR!SALARY = Text9.Text
 SR!APPROVALSTATUS = Combo1.Text
 
 If Combo1.Text = "" Then
 MsgBox "PLEASE APPROVE OR REJECT THE APPROVAL!", vbExclamation, "PLEASE VERIFY"
 Else
 SR.update
 MsgBox "UPDATED SUCCESSFULLY", vbExclamation, "SUCCESS"
  End
 
 Dim res As String
 res = MsgBox("Are You Sure You Want To Update?", vbYesNo, "Update")

Module1.opencon
Cn.Execute ("INSERT INTO SRINSERT VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & i2 & "','" & DTPicker1.Value & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & DTPicker2.Value & "','" & Combo2.Text & "')")

End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Text7.Text = Int(((Rnd * 9) * 1000000) + 1)
Text8.Text = Int(((Rnd * 9) * 100000) + 1)
End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
If keyacsii < 48 And KeyAscii > 57 Then
Text6.Locked = True
Else
Text6.Locked = False
End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If keyacsii < 48 And KeyAscii > 57 Then
Text6.Locked = True
Else
Text6.Locked = False
End If

End Sub

