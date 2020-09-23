VERSION 5.00
Begin VB.Form FrmChangePass 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   1950
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4455
      Begin VB.TextBox TxtConfirmPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox TxtNewPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox TxtCurrentPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox CmbUserName 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00EFEADE&
         Caption         =   "Confirm Pass"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00EFEADE&
         Caption         =   "New Password"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00EFEADE&
         Caption         =   "Current Pass"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EFEADE&
         Caption         =   "User Name"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00EFEADE&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00EFEADE&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

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
        MsgBox "Please select username", vbInformation
        CmbUserName.SetFocus
        Exit Sub
    End If
    If Trim(TxtCurrentPass.Text) = "" Then
        MsgBox "Please enter current password", vbInformation
        TxtCurrentPass.SetFocus
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
    Set Recset = Nothing
    Set Recset = Con.Execute("select * from users where userids='" & Trim(CmbUserName.Text) & "' and userpasswords='" & Trim(TxtCurrentPass.Text) & "'")
    If Recset.EOF = False Then
        Con.Execute "update users set userpasswords='" & Trim(TxtNewPass.Text) & "' where userids='" & Trim(CmbUserName.Text) & "'"
        MsgBox "User Password successfully changed", vbInformation
        Exit Sub
    Else
        MsgBox "Current Password not matched", vbInformation
        Exit Sub
    End If
End Sub
