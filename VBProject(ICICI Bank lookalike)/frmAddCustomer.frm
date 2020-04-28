VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddCustomer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD CUSTOMER"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10830
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FEMALE"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2760
      TabIndex        =   37
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   6120
      TabIndex        =   35
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10815
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bank Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   8400
         TabIndex        =   33
         Top             =   240
         Width           =   2295
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frmAddCustomer.frx":0000
            Left            =   240
            List            =   "frmAddCustomer.frx":0007
            TabIndex        =   41
            Top             =   2880
            Width           =   1815
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmAddCustomer.frx":0014
            Left            =   240
            List            =   "frmAddCustomer.frx":001B
            TabIndex        =   38
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   " Address Proof Document                                STATUS"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   240
            TabIndex        =   40
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ID Proof Document          STATUS                               "
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   360
            TabIndex        =   39
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Personal Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   4095
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "MALE"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1320
            TabIndex        =   36
            Top             =   2880
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtpDateOfBirth 
            Height          =   375
            Left            =   1320
            TabIndex        =   4
            Top             =   2280
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   107085825
            CurrentDate     =   38954
         End
         Begin VB.TextBox txtFirstName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   0
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtLastName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   1
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtOfficeAddress 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   2
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtHomeAddress 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   3
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   5
            Top             =   3240
            Width           =   2655
         End
         Begin VB.TextBox txtOfficePhone 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   6
            Top             =   3720
            Width           =   2655
         End
         Begin VB.TextBox txtHomePhone 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   7
            Top             =   4200
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Office Address"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Home Address"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Birth"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail Address"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Office Phone"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   3720
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Home Phone"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   4200
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Account Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   4320
         TabIndex        =   16
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtAccountNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   8
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtATM 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   9
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtBalance 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   11
            Top             =   2280
            Width           =   2655
         End
         Begin VB.TextBox txtAccessCode 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   12
            Top             =   2760
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker dtpExpiryDate 
            Height          =   375
            Left            =   1200
            TabIndex        =   10
            Top             =   1800
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   107085825
            CurrentDate     =   38954
         End
         Begin MSComCtl2.DTPicker dtpOpenDate 
            Height          =   375
            Left            =   1200
            TabIndex        =   13
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   107085825
            CurrentDate     =   38954
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Account No"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "ATM Card No"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Access Code"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Opened"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         TabIndex        =   15
         Top             =   5160
         Width           =   10575
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Caption         =   "&OK"
            Height          =   495
            Left            =   1920
            TabIndex        =   34
            Top             =   240
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmAddCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public strPath As String



Private Sub cmdClose_Click()

    Unload Me
   
End Sub

Private Sub Command1_Click()
    If txtFirstName.Text = "" Then
        MsgBox "Please enter the First Name.", vbExclamation, Title
        txtFirstName.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If txtLastName.Text = "" Then
        MsgBox "Please enter the Last Name.", vbExclamation, Title
        txtLastName.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If dtpDateOfBirth.Value = Date Then
        MsgBox "Date of Birth can not be today, Kindly change it", vbExclamation, Title
        dtpDateOfBirth.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If dtpDateOfBirth.Value = Date Then
        MsgBox "Date of Birth can not be in future", vbExclamation, Title
        dtpDateOfBirth.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If Option1.Value = False And Option2.Value = False Then
        MsgBox "Please select the Gender.", vbExclamation, Title
        cboGender.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If txtAccountNo.Text = "" Then
        MsgBox "Please enter the Account Number.", vbExclamation, Title
        txtAccountNo.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    
    If txtBalance.Text = "" Then
        MsgBox "Please enter the Balance.", vbExclamation, Title
        txtBalance.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If blAddCustomer = True Then
        recCustomers.AddNew
        recCustomers!FirstName = txtFirstName.Text & ""
        recCustomers!LastName = txtLastName.Text & ""
        recCustomers!OfficeAddress = txtOfficeAddress.Text & ""
        recCustomers!HomeAddress = txtHomeAddress.Text & ""
        recCustomers!DateOfBirth = dtpDateOfBirth.Value
        If Option1.Caption = "Male" Then
            recCustomers!Gender = "M"
            Else
                recCustomers!Gender = "F"
        End If
        recCustomers!Email = txtEmail.Text & ""
        recCustomers!OfficePhone = txtOfficePhone.Text & ""
        recCustomers!HomePhone = txtHomePhone.Text & ""
        recCustomers!AccountNumber = txtAccountNo.Text & ""
        recCustomers!AccountType = cboAccountType.Text & ""
        recCustomers!ATMCardNumber = txtATM.Text & ""
        recCustomers!ExpiryDate = dtpExpiryDate.Value & ""
        recCustomers!Balance = txtBalance.Text & ""
        recCustomers!AccessCode = txtAccessCode.Text & ""
        recCustomers!DateOpened = dtpDateOpened
        recCustomers!CustomerPicture = strPath
        recCustomers.Update
        
        Unload Me
        frmCustomers.Show
    End If
    
    If blUpdateCustomer = True Then
    
        recCustomers.MoveFirst
        Do While Not recCustomers.EOF
            If recCustomers!AccountNumber = Trim(txtAccountNo.Text) Then
                recCustomers!FirstName = txtFirstName.Text & ""
                recCustomers!LastName = txtLastName.Text & ""
                recCustomers!OfficeAddress = txtOfficeAddress.Text & ""
                recCustomers!HomeAddress = txtHomeAddress.Text & ""
                recCustomers!DateOfBirth = dtpDateOfBirth.Value
                If cboGender.Text = "Male" Then
                    recCustomers!Gender = "M"
                    Else
                        recCustomers!Gender = "F"
                End If
                recCustomers!Email = txtEmail.Text & ""
                recCustomers!OfficePhone = txtOfficePhone.Text & ""
                recCustomers!HomePhone = txtHomePhone.Text & ""
                recCustomers!AccountNumber = txtAccountNo.Text & ""
                recCustomers!AccountType = cboAccountType.Text & ""
                recCustomers!ATMCardNumber = txtATM.Text & ""
                recCustomers!ExpiryDate = dtpExpiryDate.Value & ""
                recCustomers!Balance = txtBalance.Text & ""
                recCustomers!AccessCode = txtAccessCode.Text & ""
                recCustomers!DateOpened = dtpDateOpened
                recCustomers!CustomerPicture = strPath
                recCustomers.UpdateBatch adAffectCurrent
                
            End If
            recCustomers.MoveNext
        Loop
        
        Unload Me
        frmCustomers.Show
    
    End If
        
EXITPROCEDURE:
    Exit Sub

End Sub







Private Sub Form_Load()
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3
    Call ConnectMe
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blAddCustomer = False
    blUpdateCustomer = False
End Sub



Private Sub txtBalance_KeyPress(KeyAscii As Integer)
    Call onlyNumbers(KeyAscii)
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        KeyAscii = 0
    End Select
End Sub
