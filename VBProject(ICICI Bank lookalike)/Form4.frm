VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "ICICI BANK(CUSTOMER DETAILS)"
   ClientHeight    =   6600
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   7125
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   6390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   11271
      ButtonWidth     =   6218
      ButtonHeight    =   1852
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "INSERT CUSTOMER DETAILS"
            Key             =   "INSERT1"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "SEARCH INSERTED CUSTOMERS"
            Key             =   "SEARCH1"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "SEARCH APPROVED CUSTOMERS"
            Key             =   "SEARCH2"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "APPROVE"
            Key             =   "UPDATE"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "VIEW DATAGRID (INSERTED CUSTOMERS)"
            Key             =   "DATAGRID1"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "VIEW DATAGRID (APPROVED CUSTOMERS)"
            Key             =   "DATAGRID2"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Caption         =   "VIEW LOGIN ACCESS"
            Key             =   "VIEW"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "DELETE INSERTED CUSTOMERS"
            Key             =   "DELETE"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "DELETE APPROVED CUSTOMERS"
            Key             =   "DELETE2"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "EDIT INSERTED CUSTOMERS"
            Key             =   "EDIT1"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "EDIT APPROVED CUSTOMERS"
            Key             =   "EDIT2"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "EXIT"
            Key             =   "exit"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form4.frx":0152
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form4.frx":7C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form4.frx":37CAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form4.frx":44CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form4.frx":69F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form4.frx":71A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form4.frx":7D922
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu UTILITY 
      Caption         =   "&FILE"
      Begin VB.Menu INSERT 
         Caption         =   "INSERT CUSTOMER"
         Shortcut        =   ^N
      End
      Begin VB.Menu SEARCH 
         Caption         =   "SEARCH CUSTOMER"
         Shortcut        =   ^S
      End
      Begin VB.Menu DELETE 
         Caption         =   "DELETE CUSTOMER INFO"
         Shortcut        =   ^D
      End
      Begin VB.Menu EDIT 
         Caption         =   "EDIT CUSTOMER INFO"
         Shortcut        =   ^E
      End
      Begin VB.Menu SEARCH1 
         Caption         =   "SEARCH APPROVED CUSTOMERS"
         Shortcut        =   ^T
      End
      Begin VB.Menu DELETE2 
         Caption         =   "DELETE APPROVED  CUSRTOMERS"
         Shortcut        =   ^A
      End
      Begin VB.Menu EDIT2 
         Caption         =   "EDIT APPROVED CUSTOMERS"
         Shortcut        =   ^M
      End
      Begin VB.Menu CP 
         Caption         =   "CHANGE PASSWORD"
         Shortcut        =   ^P
      End
      Begin VB.Menu quit 
         Caption         =   "QUIT"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu VIEW 
      Caption         =   "&VIEW"
      Begin VB.Menu DATAGRID 
         Caption         =   "DATAAGRID"
         Shortcut        =   ^G
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CP_Click()
Form14.Show

End Sub

Private Sub DATAGRID_Click()
Form9.Show
End Sub

Private Sub DELETE2_Click()
Form2.Caption = "DELETE CUSTOMER DETAILS"
     Form2.Label18.Caption = "ENTER CUSTOMER ID:"
     Form2.Command1.Visible = False
     Form2.Command4.Visible = True
     Form2.Command5.Visible = True
     Form2.Show
     
End Sub

Private Sub EDIT_Click()

Form2.Show
    Form2.Command1.Visible = False
    Form2.Command7.Visible = True
    Form2.Command8.Visible = True

End Sub

Private Sub EDIT2_Click()
 Form2.Show
    Form2.Caption = "EDIT CUSTOMER DETAILS"
    Form2.Label18.Caption = "EDIT CUSTOMER DETAILS"
    Form2.Command1.Visible = False
    Form2.Command9.Visible = True
    Form2.Command10.Visible = True

End Sub

Private Sub Form_Load()

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
'Form4.Picture = LoadPicture(App.Path & "/A.jpg")

End Sub

Private Sub INSERT_Click()
 Form8.Show
End Sub


Private Sub quit_Click()
Unload Me
End Sub

Private Sub SEARCH_Click()
Form2.Label18.Caption = "ENTER REQUEST ID:"
     Form2.Frame3.Visible = False

Form2.Height = 8610
     
Form2.Show
End Sub
Private Sub DELETE_Click()
Form2.Command3.Visible = True
Form2.Command1.Visible = True
'Form2.Command3.Visible = False
 '    Form2.Show

'  Form2.Command1.Visible = True
 '    Form2.Command3.Visible = True
  '   Form2.Command4.Enabled = False
   '  Form2.Command2.Enabled = False
     Form2.Show
     
End Sub

'Private Sub SRD_Click()
'Form10.Show

'End Sub

Private Sub SEARCH1_Click()
Form2.Caption = "SEARCH AND DISPLAY CUSTOMER DETAILS"
'Form2.Label18.Visible = False
'Form2.Label20.Visible = True
Form2.Command1.Enabled = False
     Form2.Command1.Visible = False
     Form2.Command4.Enabled = True
     Form2.Command4.Visible = True
     Form2.Frame3.Visible = False
Form2.Label18.Caption = "ENTER CUSTOMER ID"
Form2.Height = 8610

     Form2.Show 1
          

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   If Button.Key = "exit" Then
    Dim res As String
    res = MsgBox("Are You Sure You Want To Exit?", vbYesNo, "Exit")
    If res = vbYes Then
    End
    'Form3.Show
    Else
    Me.Show
    End If
    ElseIf Button.Key = "EDIT1" Then
    Form2.Show
    Form2.Caption = "EDIT CUSTOMER DETAILS"
    Form2.Label18.Caption = "ENTER REQUEST ID"
    Form2.Command1.Visible = False
    Form2.Command7.Visible = True
    Form2.Command8.Visible = True
  
    ElseIf Button.Key = "EDIT2" Then
 Form2.Show
    Form2.Caption = "EDIT CUSTOMER DETAILS"
    Form2.Label18.Caption = "EDIT CUSTOMER ID"
    Form2.Command1.Visible = False
    Form2.Command9.Visible = True
    Form2.Command10.Visible = True
  
    ElseIf Button.Key = "UPDATE" Then
    Form2.Caption = "APPROVE CUSTOMER DETAILS"
     Form2.Command1.Visible = False
      Form2.Command1.Enabled = False
       Form2.Command6.Visible = True
     Form2.Command2.Visible = True
     'Form2.Command4.Enabled = False
     'Form2.Command4.Visible = False
          'Form2.Command5.Visible = False
     Form2.Show 1
          ElseIf Button.Key = "SEARCH1" Then
          Form2.Caption = "SEARCH AND DISPLAY CUSTOMER DETAILS"
Form2.Label18.Caption = "ENTER REQUEST ID"
     Form2.Frame3.Visible = False

Form2.Height = 8610
     Form2.Show 1
     
     ElseIf Button.Key = "SEARCH2" Then
Form2.Caption = "SEARCH AND DISPLAY CUSTOMER DETAILS"
'Form2.Label18.Visible = False
'Form2.Label20.Visible = True
Form2.Command1.Enabled = False
     Form2.Command1.Visible = False
     Form2.Command4.Enabled = True
     Form2.Command4.Visible = True
     Form2.Frame3.Visible = False
Form2.Label18.Caption = "ENTER CUSTOMER ID"
Form2.Height = 8610

     Form2.Show 1
          
     ElseIf Button.Key = "DELETE2" Then
     Form2.Caption = "DELETE CUSTOMER DETAILS"
     Form2.Label18.Caption = "ENTER CUSTOMER ID:"
     Form2.Command1.Visible = False
     Form2.Command4.Visible = True
     Form2.Command5.Visible = True
     Form2.Show
     
     ElseIf Button.Key = "DELETE" Then
     Form2.Caption = "DELETE CUSTOMER DETAILS"
Form2.Command3.Visible = True
Form2.Command1.Visible = True
'Form2.Command3.Visible = False
     Form2.Show
     'Form2.Command1.Enabled = False
     'Form2.Command3.Enabled = False
     'Form2.Command2.Enabled = False
'Form2.Command3.Visible = False
 '    Form2.Command2.Visible = False
'Form2.Command1.Visible = False
'Form2.Label18.Visible = False
'Form2.Text18.Visible = False


ElseIf Button.Key = "INSERT1" Then
    
    Form8.Show 1
'ElseIf Button.Key = "DELETE2" Then
 '    Form2.Command1.Visible = True
  '   Form2.Command3.Visible = True
   '  Form2.Command4.Enabled = False
   '  Form2.Command2.Enabled = False
    ' Form2.Show 1
     
     
ElseIf Button.Key = "DATAGRID1" Then
     Form9.Show 1
          
ElseIf Button.Key = "DATAGRID2" Then
     Form6.Show 1
'ElseIf Button.Key = "VIEW" Then
'     Form5.Show 1
        
      End If
      
End Sub
