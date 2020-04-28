VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form10 
   BackColor       =   &H8000000D&
   Caption         =   "ICICI BANK(SALES REPRESENTATIVE VIEW)"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12855
   LinkTopic       =   "Form10"
   ScaleHeight     =   3225
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   3240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   5715
      ButtonWidth     =   7488
      ButtonHeight    =   1852
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&INSERT SALES REPRESNTATIVE"
            Key             =   "INSERT"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&UPDATE"
            Key             =   "UPDATE"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&SEARCH INSERTED SALES REPRESENTATIVE"
            Key             =   "SEARCH1"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "S&EARCH APPROVED SALES REPRESENTATIVES"
            Key             =   "SEARCH2"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "DELETE INSERTED SALES REPRESE&NTATIVE"
            Key             =   "DELETE1"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "DELETE &APPROVED SALES REPRESENTATIVE"
            Key             =   "DELETE2"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "VIEW INSERTED SALES REPRESENTATI&VE DETAILS"
            Key             =   "DATAGRID1"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "VIEW APPROVED SALES REPRESENTATIVE DETAI&LS"
            Key             =   "DATAGRID2"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "E&XIT"
            Description     =   "EXIT"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":7B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":37B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":3F666
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":648DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":6C3E6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Form10.Picture = LoadPicture(App.Path & "/a.jpg")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
If Button.Key = "INSERT" Then
'Form11.Frame3.Caption = ""
'Form11.Label15.Visible = False
'Form11.Command4.Visible = False
Form11.Show
ElseIf Button.Key = "SEARCH1" Then
Form12.Label13.Visible = False
Form12.Label14.Visible = False
Form12.Text14.Visible = False
Form12.Text11.Visible = False
Form12.Height = 9630
Form12.Frame3.Visible = False
'Form11.Option1.Visible = False
Form12.DTPicker1.Visible = False
'Form12.Text14.Visible = True

'Form11.Option1.Visible = False
Form12.Command2.Visible = False
Form12.Show
ElseIf Button.Key = "SEARCH2" Then
Form12.Frame3.Visible = False
Form12.DTPicker1.Visible = False
Form12.Text14.Visible = True

Form12.Command4.Visible = True
Form12.Combo1.Enabled = False
Form12.Combo2.Enabled = False
Form12.DTPicker1.Enabled = False
Form12.Text11.Enabled = False
Form12.Command3.Visible = False
Form12.Show
ElseIf Button.Key = "UPDATE" Then
Form12.Show
ElseIf Button.Key = "DATAGRID1" Then
Form7.Show
ElseIf Button.Key = "DATAGRID2" Then
Form13.Show
ElseIf Button.Key = "DELETE1" Then
Form12.Label15.Visible = False
Form12.Label13.Visible = False
Form12.Label14.Visible = False
Form12.Text14.Visible = False
Form12.Text11.Visible = False
Form12.DTPicker1.Visible = False
Form12.Text9.Visible = False
Form12.Combo2.Visible = False
Form12.Frame3.Caption = ""
Form12.Command1.Visible = False
Form12.Command6.Visible = False
Form12.Command5.Visible = True
Form12.Show
ElseIf Button.Key = "DELETE2" Then
Form12.Frame3.Caption = ""
Form12.Command4.Visible = True
Form12.Command3.Visible = False
Form12.Command1.Visible = False
Form12.Command5.Visible = False
Form12.Command6.Visible = True
Form12.Label15.Visible = False
Form12.Combo2.Visible = False
Form12.Show

End If
    
End Sub
