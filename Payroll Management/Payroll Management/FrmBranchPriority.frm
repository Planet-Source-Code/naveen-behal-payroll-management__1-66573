VERSION 5.00
Begin VB.Form FrmBranchPriority 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Branch Priority Section"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdDown 
      BackColor       =   &H00EFEADE&
      Caption         =   "Down"
      Height          =   315
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton CmdUp 
      BackColor       =   &H00EFEADE&
      Caption         =   "Up"
      Height          =   315
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00EFEADE&
      Caption         =   "O.K."
      Height          =   315
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00EFEADE&
      Caption         =   "Exit"
      Height          =   315
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EFEADE&
      Caption         =   "Left click for ""UP"" Right Click For ""DOWN"""
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.ListBox List1 
         Height          =   5715
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "FrmBranchPriority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDown_Click()
    curs = List1.ListIndex
    If curs = List1.ListCount - 1 Then
        MsgBox "This is the Last item in the list"
        Exit Sub
    End If
    Downitem = List1.List(curs + 1)
    List1.RemoveItem (curs + 1)
    List1.AddItem Downitem, curs
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdOk_Click()
On Error GoTo er
    List1.ListIndex = 0
    Con.Execute "delete from branchpriority"
    While List1.ListIndex <= List1.ListCount - 1
        Con.Execute "insert into branchpriority values('" & List1.List(List1.ListIndex) & "'," & List1.ListIndex + 1 & ")"
        List1.ListIndex = List1.ListIndex + 1
    Wend
    MsgBox "Priority Set"
er:
End Sub
Private Sub CmdUp_Click()
    curs = List1.ListIndex
    If curs = 0 Then
        MsgBox "This is the Top item in the list"
        Exit Sub
    End If
    upitem = List1.List(curs - 1)
    List1.RemoveItem (curs - 1)
    List1.AddItem upitem, curs
End Sub
Private Sub Form_Load()
    Set Recset = Nothing
    Set Recset = Con.Execute("Select distinct(branch) from staffdetails")
    While Recset.EOF = False
        List1.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
    CenterMe Me
End Sub
