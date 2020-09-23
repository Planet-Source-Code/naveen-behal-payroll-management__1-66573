VERSION 5.00
Begin VB.Form FrmNewUser 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New User"
   ClientHeight    =   1935
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OkButton 
      BackColor       =   &H00EFEADE&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   210
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00EFEADE&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EFEADE&
         Caption         =   "Normal"
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EFEADE&
         Caption         =   "Admin"
         Height          =   195
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox CmbUserName 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox TxtNewPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox TxtConfirmPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EFEADE&
         Caption         =   "User Name"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00EFEADE&
         Caption         =   "New Password"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00EFEADE&
         Caption         =   "Confirm Pass"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterMe Me
    Set Recset = Nothing
    Set Recset = Con.Execute("select userids from users")
    While Recset.EOF = False
        CmbUserName.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
End Sub

Private Sub OKButton_Click()
    If Trim(CmbUserName.Text) = "" Then
        MsgBox "Please Enter username", vbInformation
        CmbUserName.SetFocus
        Exit Sub
    End If
    If Trim(TxtNewPass.Text) = "" Then
        MsgBox "Please enter new password", vbInformation
        TxtNewPass.SetFocus
        Exit Sub
    End If
    If Trim(TxtNewPass.Text) <> Trim(TxtConfirmPass.Text) Then
        MsgBox "New password doesn't match with confirmation pass", vbInformation
        TxtConfirmPass.SetFocus
        Exit Sub
    End If
    If Option1.Value = False And Option2.Value = False Then
        MsgBox "Please select a user type", vbInformation
        Exit Sub
    End If
    If Option1.Value = True Then
        ut = "1"
    Else
        ut = "0"
    End If
    Set Recset = Nothing
    Set Recset = Con.Execute("select * from users where userids='" & Trim(CmbUserName.Text) & "'")
    If Recset.EOF = True Then
        Con.Execute "insert into users values('" & Trim(CmbUserName.Text) & "','" & Trim(TxtNewPass.Text) & "','" & ut & "')"
        MsgBox "User successfully created", vbInformation
        Exit Sub
    Else
        MsgBox "User id already exist", vbInformation
        Exit Sub
    End If
End Sub
