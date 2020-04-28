VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13125
   LinkTopic       =   "Form12"
   ScaleHeight     =   10935
   ScaleWidth      =   13125
   StartUpPosition =   3  'Windows Default
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
      Height          =   375
      Left            =   12000
      TabIndex        =   38
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command3 
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
      Left            =   10200
      TabIndex        =   35
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Height          =   615
      Left            =   4800
      TabIndex        =   34
      Top             =   240
      Width           =   5175
   End
   Begin VB.Frame Frame3 
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3840
      TabIndex        =   29
      Top             =   8760
      Width           =   5655
      Begin VB.CommandButton Command6 
         Caption         =   "DEL&ETE"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3120
         TabIndex        =   40
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&DELETE"
         Enabled         =   0   'False
         Height          =   495
         Left            =   840
         TabIndex        =   39
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&CANCEL"
         Height          =   495
         Left            =   3120
         TabIndex        =   33
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   495
         Left            =   840
         TabIndex        =   32
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form12.frx":0000
         Left            =   1320
         List            =   "Form12.frx":000A
         TabIndex        =   31
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label15 
         Caption         =   "APPROVAL:"
         Height          =   495
         Left            =   2400
         TabIndex        =   30
         Top             =   360
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
      Height          =   7575
      Left            =   7200
      TabIndex        =   16
      Top             =   1080
      Width           =   5055
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2880
         TabIndex        =   41
         Top             =   4560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2880
         TabIndex        =   28
         Top             =   5400
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form12.frx":0022
         Left            =   2880
         List            =   "Form12.frx":002C
         TabIndex        =   27
         Top             =   3840
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2880
         TabIndex        =   26
         Top             =   4560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   107151361
         CurrentDate     =   42103
      End
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         Height          =   615
         Left            =   2880
         TabIndex        =   23
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   615
         Left            =   2160
         TabIndex        =   20
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   615
         Left            =   2160
         TabIndex        =   18
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label14 
         Caption         =   "SALARY :"
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
         Left            =   720
         TabIndex        =   25
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "JOINING DATE :"
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
         Left            =   600
         TabIndex        =   24
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "QUALIFICATION STATUS :"
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
         TabIndex        =   22
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Label Label11 
         Caption         =   "INTERVIEW DATE :"
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
         Left            =   480
         TabIndex        =   21
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5040
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label10 
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
         Height          =   735
         Left            =   480
         TabIndex        =   19
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "ID NUMBER :"
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
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   1695
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
      Height          =   7575
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   6255
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1920
         TabIndex        =   37
         Top             =   3960
         Width           =   3495
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1920
         TabIndex        =   15
         Top             =   6720
         Width           =   3495
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1920
         TabIndex        =   13
         Top             =   5760
         Width           =   3495
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1920
         TabIndex        =   11
         Top             =   4800
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1920
         TabIndex        =   9
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1920
         TabIndex        =   7
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1920
         TabIndex        =   5
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label16 
         Caption         =   "DOB:"
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
         Left            =   360
         TabIndex        =   36
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label8 
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
         Height          =   495
         Left            =   480
         TabIndex        =   14
         Top             =   6840
         Width           =   1455
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   12
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   10
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label Label5 
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
         Left            =   360
         TabIndex        =   8
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         TabIndex        =   6
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "ENTER SALES REPRESENTATIVE ID :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    If Combo2.Text = "" Then
    MsgBox "Please Enter Approval Status", vbExclamation, Title
  
  ElseIf Text11.Text = "" Then
    MsgBox "Please Enter Salary", vbExclamation, Title
  
  
  ElseIf DTPicker1.Value = Date Then
        MsgBox "Joining Date can not be today, Kindly change it", vbExclamation, Title
     
Exit Sub
        Else
        Call update
End If
End Sub
Private Sub update()
If Combo2.Text = "APPROVED" Then
Module1.opencon
    Cn.Execute ("INSERT INTO SRAPPROVED VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text13.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & Combo1.Text & "','" & DTPicker1.Value & "','" & Text11.Text & "','" & Combo2.Text & "')")
MsgBox "UPDATED SUCCESSFULLY", vbExclamation, "DONE"
CreateObject("SAPI.SPVOICE").SPEAK "UPDATED SUCCESSFULLY"
Unload Me
ElseIf Combo2.Text = "REJECTED" Then
Dim SR As ADODB.Recordset
 Set SR = New ADODB.Recordset
 SR.Open "select * from SRINSERT where IDNUMBER='" & Text12.Text & "'", Cn, adOpenStatic, adLockOptimistic
 
 If SR.EOF Then
 SR.AddNew
 Else: SR.DELETE
MsgBox "UPDATED SUCCESSFULLY", vbExclamation, "DONE"
CreateObject("SAPI.SPVOICE").SPEAK "UPDATED SUCCESSFULLY"
Unload Me
 
 End If
 
 End If


End Sub

Private Sub Command3_Click()
Module1.opencon
Dim SR1 As ADODB.Recordset
Set SR1 = New Recordset
SR1.Open "select * from SRINSERT where IDNUMBER='" & Text12.Text & "'", Cn, adOpenStatic, adLockOptimistic

If Text12.Text = "" Then
MsgBox ">> PLEASE ENTER THE ID NUMBER!", vbExclamation, "PLEASE ENTER ID!"

ElseIf SR1.EOF Then
MsgBox "  >> Record corresponding to this SALES REPRESENTATIVE was not found !! ", vbExclamation, "Record absent !!"
Exit Sub
End If
Command1.Enabled = True
Command5.Enabled = True

DTPicker1.Enabled = True
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
Set Text7.DataSource = SR1
    Text7.DataField = "EMAIL"
    Set Text5.DataSource = SR1
    Text5.DataField = "HOMEPHONE"
Set Text6.DataSource = SR1
    Text6.DataField = "OFFICEPHONE"
Set Text4.DataSource = SR1
    Text4.DataField = "GENDER"
Set Text13.DataSource = SR1
    Text13.DataField = "DOB"
Set Text8.DataSource = SR1
    Text8.DataField = "IDNUMBER"
Set Text10.DataSource = SR1
    Text10.DataField = "INTERVIEWDATE"
'Set Combo3.DataSource = recCustomer
 '   Combo3.DataField = "ACCOUNTTYPE"
Set Text9.DataSource = SR1
    Text9.DataField = "ACCESSCODE"
'Set Text13.DataSource = recCustomer
 '   Text13.DataField = "EXPIRYDATE"
'Set Text14.DataSource = recCustomer
 '   Text14.DataField = "ACCESSCODE"
'Set Combo1.DataSource = recCustomer
 '   Combo1.DataField = "IDPROOFSTATUS"
Set Combo1.DataSource = SR1
    Combo1.DataField = "QUALIFICATIONSTATUS"
    



End Sub

Private Sub Command4_Click()
Dim sr2 As ADODB.Recordset
Set sr2 = New Recordset
sr2.Open "select * from SRAPPROVED where IDNUMBER='" & Text12.Text & "'", Cn, adOpenStatic, adLockOptimistic

If Text12.Text = "" Then
MsgBox ">> PLEASE ENTER THE ID NUMBER!", vbExclamation, "PLEASE ENTER ID!"

ElseIf sr2.EOF Then
MsgBox "  >> Record corresponding to this SALES REPRESENTATIVE was not found !! ", vbExclamation, "Record absent !!"
Exit Sub
End If
Command6.Enabled = True
Set Text1.DataSource = sr2
    Text1.DataField = "FIRSTNAME"
    'lbl_name.Caption =  + lbl_name.Caption
Set Text2.DataSource = sr2
    Text2.DataField = "LASTNAME"
Set Text3.DataSource = sr2
    Text3.DataField = "HOMEADDRESS"
    'lbl_col.Caption = lbl_col.Caption + " College"
'Set Text4.DataSource = recCustomer
 '   Text4.DataField = "OFFICEADDRESS"
    'lbl_strm.Caption = "Stream " + lbl_strm.Caption
Set Text7.DataSource = sr2
    Text7.DataField = "EMAIL"
    Set Text5.DataSource = sr2
    Text5.DataField = "HOMEPHONE"
Set Text6.DataSource = sr2
    Text6.DataField = "OFFICEPHONE"
Set Text4.DataSource = sr2
    Text4.DataField = "GENDER"
Set Text13.DataSource = sr2
    Text13.DataField = "DOB"
Set Text8.DataSource = sr2
    Text8.DataField = "IDNUMBER"
Set Text10.DataSource = sr2
    Text10.DataField = "INTERVIEWDATE"
'Set Combo3.DataSource = recCustomer
 '   Combo3.DataField = "ACCOUNTTYPE"
Set Text9.DataSource = sr2
    Text9.DataField = "ACCESSCODE"
Set Text14.DataSource = sr2
    Text14.DataField = "JOININGDATE"
Set Text11.DataSource = sr2
    Text11.DataField = "SALARY"
Set Combo2.DataSource = sr2
    Combo2.DataField = "APPROVALSTATUS"
Set Combo1.DataSource = sr2
    Combo1.DataField = "QUALIFICATIONSTATUS"
    

End Sub

Private Sub Command5_Click()
Dim SRD As ADODB.Recordset
 Set SRD = New ADODB.Recordset
 SRD.Open "select * from SRINSERT where IDNUMBER='" & Text12.Text & "'", Cn, adOpenStatic, adLockOptimistic
 
 Dim res As String
 res = MsgBox("Are You Sure You Want To Delete?", vbYesNo, "Delete")
 If res = vbYes Then
 If SRD.EOF Then
 SRD.AddNew
 Else: SRD.DELETE
 End If
 
 MsgBox "DELETED SUCCESSFULLY", vbExclamation
 CreateObject("SAPI.SPVOICE").SPEAK "DELETED SUCCESSFULLY"
 End If

End Sub

Private Sub Command6_Click()
Dim srd1 As ADODB.Recordset
 Set srd1 = New ADODB.Recordset
 srd1.Open "select * from SRAPPROVED where IDNUMBER='" & Text12.Text & "'", Cn, adOpenStatic, adLockOptimistic
 
 Dim res As String
 res = MsgBox("Are You Sure You Want To Delete?", vbYesNo, "Delete")
 If res = vbYes Then
 If srd1.EOF Then
 srd1.AddNew
 Else: srd1.DELETE
 End If
 
 MsgBox "DELETED SUCCESSFULLY", vbExclamation
 CreateObject("SAPI.SPVOICE").SPEAK "DELETED SUCCESSFULLY"
 End If


End Sub

