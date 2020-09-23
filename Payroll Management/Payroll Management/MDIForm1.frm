VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00A07247&
   Caption         =   "Salary Automation System"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2820
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Picture         =   "MDIForm1.frx":0E42
            Text            =   "User Name:"
            TextSave        =   "User Name:"
            Object.ToolTipText     =   "login user name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "waiting..."
            TextSave        =   "waiting..."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2963
            Picture         =   "MDIForm1.frx":1C96
            Text            =   "Time Log-in:"
            TextSave        =   "Time Log-in:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "waiting..."
            TextSave        =   "waiting..."
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "11/10/06"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Text            =   "Caps Lock"
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Text            =   "Num Lock"
            TextSave        =   "Num Lock"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Text            =   "Insert"
            TextSave        =   "Insert"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   4080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":23F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":30CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3820
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":44FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":51D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7862
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":853C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9216
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9EF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":ABCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B8A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C57E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D258
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":DF32
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":EC0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F8E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":105C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1129A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":116EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":12042
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":12186
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":122CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":123F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":12546
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   8400
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   42
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1264A
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13324
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15AD8
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1692A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1777C
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":18056
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":18930
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1920A
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":19BD4
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A4AE
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A7C8
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B0A2
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B97C
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C256
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C570
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CE4A
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D724
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1DFFE
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E8D8
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F1B2
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1FA8C
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":20366
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":20C40
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2151A
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":21DF4
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":226CE
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":22FA8
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":23882
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2415C
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":24A36
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":25310
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":25BC6
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":264A0
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":268F2
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26D44
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":294F6
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":296A6
            Key             =   "IMG36"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":297AA
            Key             =   "IMG37"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":29896
            Key             =   "IMG38"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":299BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":29B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":29C6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   3  'Align Left
      Height          =   2820
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   4974
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "i32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Employee Details"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salary Sheet"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Arrear"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Back Up"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New User"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Change Password"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Remove User"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Log-Off"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "Enter Text to Search"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A07247&
         Height          =   360
         Left            =   0
         TabIndex        =   2
         Top             =   2205
         Width           =   5000
      End
   End
   Begin VB.Menu mnumain 
      Caption         =   "Main"
      Begin VB.Menu mnulogin 
         Caption         =   "Log-in"
         Begin VB.Menu mnuloginadmin 
            Caption         =   "Admin"
            Shortcut        =   {F12}
         End
         Begin VB.Menu mnuloginusers 
            Caption         =   "Users"
            Shortcut        =   {F11}
         End
      End
      Begin VB.Menu mnulogoff 
         Caption         =   "Lof Off"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuadmin 
      Caption         =   "Admin"
      Begin VB.Menu mnuemployeedetails 
         Caption         =   "Employee Details"
         Shortcut        =   ^E
      End
      Begin VB.Menu tdsmaster 
         Caption         =   "TDS Conditions Master"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuemployeetransfer 
         Caption         =   "Transfer"
      End
      Begin VB.Menu mnuseptds 
         Caption         =   "-"
      End
      Begin VB.Menu mnusalarysheet 
         Caption         =   "Salary Sheet"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuarrear 
         Caption         =   "Arrear"
      End
      Begin VB.Menu mnumedical 
         Caption         =   "Medical"
      End
      Begin VB.Menu mnubackup 
         Caption         =   "Back UP"
      End
      Begin VB.Menu mu 
         Caption         =   "-"
      End
      Begin VB.Menu mnunewuser 
         Caption         =   "New User"
      End
      Begin VB.Menu mnuchangepass 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnudeleteuser 
         Caption         =   "Delete User"
      End
      Begin VB.Menu mnuBranchPriority 
         Caption         =   "Branch Priority"
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "Reports"
      Begin VB.Menu MnusalarySheetMonthly 
         Caption         =   "Monthly Salary Sheet"
      End
      Begin VB.Menu mnusalarysheetd 
         Caption         =   "Salary Sheet With Deductions"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnusalaryslip 
         Caption         =   "Salary Slip for the month"
      End
      Begin VB.Menu mnudetailedsalaryslip 
         Caption         =   "Detailed Salary Slip"
      End
      Begin VB.Menu mnustafflist 
         Caption         =   "Staff List-Details"
      End
      Begin VB.Menu mnuleavedstaff 
         Caption         =   "Leaved Staff Details"
      End
      Begin VB.Menu mnutaxableincome 
         Caption         =   "Taxable Income Staff"
      End
      Begin VB.Menu mnumonthlypay 
         Caption         =   "Monthly Salary Payable"
      End
      Begin VB.Menu mwsps 
         Caption         =   "Month Wise Salary Paid Summary"
      End
      Begin VB.Menu tdspaid 
         Caption         =   "TDS PAID"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuarrear_Click()
    FrmArrear.Show
End Sub

Private Sub mnubackup_Click()
    FrmBackUp.Show
End Sub

Private Sub mnuBranchPriority_Click()
    FrmBranchPriority.Show
End Sub

Private Sub mnuchangepass_Click()
    FrmChangePass.Show
End Sub

Private Sub mnudeleteuser_Click()
    FrmDeleteUser.Show
End Sub
Private Sub mnudetailedsalaryslip_Click()
    Dialog1.Caption = "Detailed Salary Slip"
    Dialog1.Show
End Sub
Private Sub mnuemployeedetails_Click()
    FrmStaff.Show
End Sub
Private Sub mnuemployeetransfer_Click()
    FrmTransfer.Show
End Sub
Private Sub mnuexit_Click()
    End
End Sub
Private Sub mnuleavedstaff_Click()
    Set DataEnvironment1 = Nothing
    DataEnvironment1.EmployeeDetails_Grouping "1"
    DtrStaffDetails.Show
End Sub
Private Sub mnuloginadmin_Click()
    UType = 1
    frmLogin.Show
End Sub
Private Sub mnuloginusers_Click()
    UType = 0
    frmLogin.Show
End Sub
Private Sub mnulogoff_Click()
    LogOff
End Sub
Private Sub mnumedical_Click()
    FrmMedical.Show
End Sub
Private Sub mnumonthlypay_Click()
    Unload DataEnvironment1
    DataEnvironment1.EmployeeDetails_Grouping "0"
    DTRMonthlyPay.Show
End Sub
Private Sub mnunewuser_Click()
    FrmNewUser.Show
End Sub
Private Sub mnusalarysheet_Click()
    FrmSalary.Show
End Sub
Private Sub mnusalarysheetd_Click()
    Dialog.Caption = "Salary Sheet With Deductions"
    Dialog.Show
End Sub
Private Sub MnusalarySheetMonthly_Click()
    Dialog.Caption = "Monthly Salary Sheet"
    Dialog.Show
End Sub
Private Sub mnusalaryslip_Click()
    Dialog1.Show
End Sub
Private Sub mnustafflist_Click()
    DataEnvironment1.EmployeeDetails_Grouping "0"
    DtrStaffDetails.Show
End Sub
Private Sub mnutaxableincome_Click()
    Set Recset = Nothing
    Set Recset = Con.Execute("Select staffid,branch,basicsalary+pf+ps+hra as Salary,tds from staffdetails where basicsalary+ps+pf+hra>12500 and leaved='0' order by branch")
    DTRTMP.Sections("Detail").Controls("text1").DataField = "Staffid"
    DTRTMP.Sections("Detail").Controls("text2").DataField = "branch"
    DTRTMP.Sections("Detail").Controls("text3").DataField = "Salary"
    DTRTMP.Sections("Detail").Controls("text4").DataField = "TDS"
    DTRTMP.Sections("PAgeHeader").Controls("Label1").Caption = "Staff ID"
    DTRTMP.Sections("PAgeHeader").Controls("Label2").Caption = "Branch Name"
    DTRTMP.Sections("PAgeHeader").Controls("Label4").Caption = "Salary"
    DTRTMP.Sections("PAgeHeader").Controls("Label6").Caption = "TDS"
    DTRTMP.Sections("ReportFooter").Controls("Function1").DataField = "Salary"
    DTRTMP.Sections("ReportFooter").Controls("Function2").DataField = "Salary"
    DTRTMP.Sections("ReportFooter").Controls("Function3").DataField = "tds"
    DTRTMP.Title = "Taxable Income Staff List"
    Set DTRTMP.DataSource = Recset
    DTRTMP.Show
End Sub
Private Sub mwsps_Click()
    FrmMWSS.Show
End Sub
Private Sub tdsmaster_Click()
    FrmTDSRange.Show
End Sub
Private Sub tdspaid_Click()
    a = InputBox("Enter Month Name")
    If Trim(a) = "" Then
        Exit Sub
    End If
        
    Set Recset = Nothing
    Set Recset = Con.Execute("Select staffid,branch,total as Salary,tds from salarysheet where tds>0 and salarymonth='" & Trim(a) & "'")
    If Recset.EOF = True Then
        MsgBox "Nothing Found!"
        Exit Sub
    End If
    DTRTMP.Sections("Detail").Controls("text1").DataField = "Staffid"
    DTRTMP.Sections("Detail").Controls("text2").DataField = "branch"
    DTRTMP.Sections("Detail").Controls("text3").DataField = "Salary"
    DTRTMP.Sections("Detail").Controls("text4").DataField = "TDS"
    DTRTMP.Sections("PAgeHeader").Controls("Label1").Caption = "Staff ID"
    DTRTMP.Sections("PAgeHeader").Controls("Label2").Caption = "Branch Name"
    DTRTMP.Sections("PAgeHeader").Controls("Label4").Caption = "Salary"
    DTRTMP.Sections("PAgeHeader").Controls("Label6").Caption = "TDS"
    DTRTMP.Sections("ReportFooter").Controls("Function1").DataField = "Salary"
    DTRTMP.Sections("ReportFooter").Controls("Function2").DataField = "Salary"
    DTRTMP.Sections("ReportFooter").Controls("Function3").DataField = "tds"
    DTRTMP.Title = "Tax Paid For The Month of " & a
    Set DTRTMP.DataSource = Recset
    DTRTMP.Show
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
                FrmStaff.Show
        Case 3
                FrmSalary.Show
        Case 5
                FrmArrear.Show
        Case 7
                FrmBackUp.Show
        Case 9
                FrmNewUser.Show
        Case 11
                FrmChangePass.Show
        Case 13
                FrmDeleteUser.Show
        Case 15
                LogOff
    End Select
End Sub
