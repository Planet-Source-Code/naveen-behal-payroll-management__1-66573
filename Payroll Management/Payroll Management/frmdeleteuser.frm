VERSION 5.00
Begin VB.Form FrmDeleteUser 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   2145
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00EFEADE&
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5775
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton OKButton 
         BackColor       =   &H00EFEADE&
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CancelButton 
         BackColor       =   &H00EFEADE&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmDeleteUser"
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
    Set Recset = Con.Execute("select * from users")
    List1.Clear
    While Recset.EOF = False
        List1.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend

End Sub

Private Sub OKButton_Click()

    If Trim(List1.Text) <> "" Then
        If Trim(List1.Text) = Trim(UserName) Then
            MsgBox "U have logged in with this user id. U can not delete it", vbInformation
            Exit Sub
        End If
        Con.Execute "delete from users where userids='" & Trim(List1.Text) & "'"
        MsgBox "User successfully deleted", vbInformation
        Form_Load
        Exit Sub
    Else
        MsgBox "Please select userid"
        Exit Sub
    End If
End Sub
