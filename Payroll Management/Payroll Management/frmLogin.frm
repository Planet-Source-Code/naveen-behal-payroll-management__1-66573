VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00EFEADE&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1335
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00EFEADE&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1020
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean
Dim LogCounter As Integer

Private Sub CmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub CmdOK_Click()
LogCounter = LogCounter + 1
    If Trim(txtUserName.Text) = "" Then
        MsgBox "Please enter user name", vbInformation
        txtUserName.SetFocus
        Exit Sub
    End If
    If Trim(txtPassword.Text) = "" Then
        MsgBox "Please enter password", vbInformation
        txtPassword.SetFocus
        Exit Sub
    End If
    OpenCon
    Set Recset = Con.Execute("select * from users where userids='" & Trim(txtUserName.Text) & "' and userpasswords='" & Trim(txtPassword.Text) & "' and usertypes='" & UType & "'")
    If Recset.EOF = True Then
        MsgBox "Invalid User ID or Password, try again!", , "Login"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
        If LogCounter = 3 Then Unload Me
        Exit Sub
    End If
    UserName = Trim(txtUserName.Text)
    MDIForm1.StatusBar1.Panels(2).Text = UserName
    MDIForm1.StatusBar1.Panels(4).Text = Time
    MDIForm1.mnureports.Visible = True
    If UType = "1" Then
        MDIForm1.mnuadmin.Visible = True
        MDIForm1.Toolbar.Visible = True
    End If
    MDIForm1.mnulogin.Visible = False
    MDIForm1.mnulogoff.Visible = True
    Unload Me
End Sub

Private Sub Form_Load()
    CenterMe Me
End Sub
