VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTransfer 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transfer Section"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00EFEADE&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton CmdTransfer 
      BackColor       =   &H00EFEADE&
      Caption         =   "&Transfer"
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EFEADE&
      Height          =   1755
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5895
      Begin MSComCtl2.DTPicker DTPTransferDate 
         Height          =   285
         Left            =   3360
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   39002
      End
      Begin VB.ComboBox CmbNewBranchName 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   3255
      End
      Begin VB.ComboBox CmbEmployee 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3255
      End
      Begin VB.ComboBox CmbBranchName 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Date"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Branch/Division Name"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1980
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Existing Branch/Division Name"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2190
      End
   End
End
Attribute VB_Name = "FrmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbBranchName_Change()
    Set TmpRecset = Nothing
    Set TmpRecset = Con.Execute("select staffid from staffdetails where branch='" & Trim(CmbBranchName.Text) & "' and leaved='0'")
    CmbEmployee.Clear
    While TmpRecset.EOF = False
        CmbEmployee.AddItem TmpRecset.Fields(0).Value
        TmpRecset.MoveNext
    Wend
End Sub

Private Sub CmbBranchName_Click()
    Set TmpRecset = Nothing
    Set TmpRecset = Con.Execute("select staffid from staffdetails where branch='" & Trim(CmbBranchName.Text) & "' and leaved='0'")
    CmbEmployee.Clear
    While TmpRecset.EOF = False
        CmbEmployee.AddItem TmpRecset.Fields(0).Value
        TmpRecset.MoveNext
    Wend
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdTransfer_Click()
    If CmbBranchName.Text = "" Then
        MsgBox "Please select branch/division name first", vbInformation
        CmbBranchName.SetFocus
        Exit Sub
    End If
    If CmbEmployee.Text = "" Then
        MsgBox "Please select employee name", vbInformation
        CmbEmployee.SetFocus
        Exit Sub
    End If
    If CmbNewBranchName.Text = "" Then
        MsgBox "Please select new branch/division name", vbInformation
        CmbNewBranchName.SetFocus
        Exit Sub
    End If
    Set Recset = Nothing
    Set Recset = Con.Execute("select * from staffdetails where staffid='" & CmbEmployee.Text & "' and branch='" & CmbBranchName.Text & "'")
    Con.Execute "insert into staffdetails values('" & Recset.Fields(0).Value & "'," & Recset.Fields(1).Value & ",'" & CmbNewBranchName.Text & "','" & DTPTransferDate.Value & "','" & Recset.Fields(4).Value & "'," & Recset.Fields(5).Value & "," & Recset.Fields(6).Value & "," & Recset.Fields(7).Value & "," & Recset.Fields(8).Value & "," & Recset.Fields(9).Value & "," & Recset.Fields(10).Value & ",'" & Recset.Fields(11).Value & "'," & Recset.Fields(12).Value & ",'" & Recset.Fields(13).Value & "'," & Recset.Fields(14).Value & "," & Recset.Fields(15).Value & ",'" & Recset.Fields(16).Value & "','" & Recset.Fields(17).Value & "')"
    Con.Execute "insert into transferdata values('" & CmbBranchName.Text & "','" & CmbEmployee.Text & "','" & CmbNewBranchName.Text & "','" & DTPTransferDate.Value & "')"
    MsgBox "Transfered"
End Sub

Private Sub Form_Load()
    CenterMe Me
    Set Recset = Nothing
    Set Recset = Con.Execute("select distinct(branch) from staffdetails")
    CmbBranchName.Clear
    CmbNewBranchName.Clear
    While Recset.EOF = False
        CmbBranchName.AddItem Recset.Fields(0).Value
        CmbNewBranchName.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
End Sub
