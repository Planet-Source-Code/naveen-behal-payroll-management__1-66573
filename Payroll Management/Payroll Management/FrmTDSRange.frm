VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmTDSRange 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EFEADE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TDS Range"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00EFEADE&
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   315
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5760
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2415
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4260
      _Version        =   393216
      BackColorFixed  =   13086108
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      Caption         =   "Rebate Section"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   2160
      Width           =   9855
      Begin VB.CommandButton CmdRemoveRebateContions 
         BackColor       =   &H00EFEADE&
         Caption         =   "Remove"
         Height          =   315
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton CmdAddRebateConditions 
         BackColor       =   &H00EFEADE&
         Caption         =   "Add"
         Height          =   315
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox CmbRebateType 
         Height          =   315
         ItemData        =   "FrmTDSRange.frx":0000
         Left            =   8400
         List            =   "FrmTDSRange.frx":000A
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtRebateHead 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox TxtRebateValue 
         Height          =   285
         Left            =   6480
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "In"
         Height          =   195
         Left            =   7680
         TabIndex        =   30
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rebate For"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Then Rebate is"
         Height          =   195
         Left            =   5040
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      Caption         =   "Rebate Section"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   9855
      Begin VB.CheckBox Check2 
         BackColor       =   &H00EFEADE&
         Caption         =   "Use This"
         Height          =   255
         Left            =   6600
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EFEADE&
         Caption         =   "Use This"
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtAgeRebate 
         Height          =   285
         Left            =   5040
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtAge 
         Height          =   285
         Left            =   2880
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtFemale 
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Then Rebate is"
         Height          =   195
         Left            =   3720
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If Age is greater than or equal to"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   2265
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If Sex is Fe-Male then rebate is Rs."
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   2460
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      Caption         =   "Salary Range"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton CmdRemoveSalRange 
         BackColor       =   &H00EFEADE&
         Caption         =   "Remove"
         Height          =   315
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton CmdAddSalRange 
         BackColor       =   &H00EFEADE&
         Caption         =   "Add"
         Height          =   315
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtValue 
         Height          =   285
         Left            =   9000
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtSecondParameter 
         Height          =   285
         Left            =   6600
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtFirstParameter 
         Height          =   285
         Left            =   5040
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox CmbCondition 
         Height          =   315
         ItemData        =   "FrmTDSRange.frx":0015
         Left            =   2880
         List            =   "FrmTDSRange.frx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   9480
         TabIndex        =   22
         Top             =   240
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Than TDS will be"
         Height          =   195
         Left            =   7680
         TabIndex        =   21
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "And"
         Height          =   195
         Left            =   6120
         TabIndex        =   20
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If gross salary of an employee is"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2235
      End
   End
End
Attribute VB_Name = "FrmTDSRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        If Val(TxtFemale.Text) <= 0 Then
            MsgBox "Please enter valid amount", vbInformation
            TxtFemale.SetFocus
            Exit Sub
        End If
        Set Recset = Con.Execute("select * from tdsrebatefemale")
        If Recset.EOF = True Then
            Con.Execute ("insert into tdsrebatefemale values(" & Val(TxtFemale.Text) & ")")
        Else
            MBox = MsgBox("Entry Already exist! Want to update it?", vbYesNo)
            If MBox = vbYes Then
                Con.Execute "update tdsrebatefemale set rebate=" & Val(TxtFemale.Text)
                MsgBox "Updated!", vbInformation
            End If
        End If
    Else
        MBox = MsgBox("Want to delete this entry?", vbYesNo)
        If MBox = vbYes Then
            Con.Execute "delete from tdsrebatefemale"
            MsgBox "Deleted.", vbInformation
        End If
    End If
    ShowRebates
End Sub
Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        If Val(TxtAge.Text) <= 0 Then
            MsgBox "Please enter valid age", vbInformation
            TxtAge.SetFocus
            Exit Sub
        End If
        If Val(TxtAgeRebate.Text) <= 0 Then
            MsgBox "Please enter valid age-rebate", vbInformation
            TxtAgeRebate.SetFocus
            Exit Sub
        End If
        Set Recset = Con.Execute("select * from tdsrebateseniorcitizen")
        If Recset.EOF = True Then
            Con.Execute ("insert into tdsrebateseniorcitizen values(" & Val(TxtAge.Text) & "," & Val(TxtAgeRebate.Text) & ")")
        Else
            MBox = MsgBox("Entry Already exist! Want to update it?", vbYesNo)
            If MBox = vbYes Then
                Con.Execute "update tdsrebateseniorcitizen set age=" & Val(TxtAge.Text) & ", rebate=" & Val(TxtAgeRebate.Text)
                MsgBox "Updated!", vbInformation
            End If
        End If
    Else
        MBox = MsgBox("Want to delete this entry?", vbYesNo)
        If MBox = vbYes Then
            Con.Execute "delete from tdsrebateseniorcitizen"
            MsgBox "Deleted.", vbInformation
        End If
    End If
    ShowRebates
End Sub
Private Sub CmdAddRebateConditions_Click()
    If Trim(TxtRebateHead.Text) = "" Then
        MsgBox "Please enter rebate head", vbInformation
        TxtRebateHead.SetFocus
        Exit Sub
    ElseIf Trim(TxtRebateValue.Text) <= 0 Then
        MsgBox "Please enter rebate value", vbInformation
        TxtRebateValue.SetFocus
        Exit Sub
    ElseIf Trim(CmbRebateType.Text) = "" Then
        MsgBox "Please select rebate type", vbInformation
        CmbRebateType.SetFocus
        Exit Sub
    End If
    Set Recset = Nothing
    Set Recset = Con.Execute("select * from tdsrebatevarious where rebatehead='" & Trim(TxtRebateHead.Text) & "'")
    If Recset.EOF = False Then
        MBox = MsgBox("Entry Already Exist! Want to continue?", vbYesNo)
        If MBox = vbYes Then
            Con.Execute "update tdsrebatevarios set rebatevalue=" & Val(TxtRebateValue.Text) & ",'" & Trim(CmbRebateType.Text) & "' where rebatehead='" & Trim(TxtRebateHead.Text) & "'"
            MsgBox "Updated!", vbInformation
        End If
    Else
        Con.Execute "insert into tdsrebatevarious values('" & Trim(TxtRebateHead.Text) & "'," & Trim(TxtRebateValue.Text) & ",'" & Trim(CmbRebateType.Text) & "')"
    End If
    ShowRebates
End Sub
Private Sub CmdAddSalRange_Click()
    If Trim(CmbCondition.Text) = "" Then
        MsgBox "Please select a condition first!", vbInformation
        CmbCondition.SetFocus
        Exit Sub
    End If
    If Val(TxtFirstParameter.Text) <= 0 Then
        MsgBox "Please enter valid value!", vbInformation
        TxtFirstParameter.SetFocus
        Exit Sub
    End If
    If Trim(CmbCondition.Text) = "Between" And Val(TxtSecondParameter.Text) <= 0 Then
        MsgBox "Please enter valid value", vbInformation
        TxtSecondParameter.SetFocus
        Exit Sub
    End If
    If Val(TxtValue.Text) < 0 Then
        MsgBox "Please enter valid value!", vbInformation
        TxtValue.SetFocus
        Exit Sub
    End If
    Set Recset = Nothing
    Set Recset = Con.Execute("select * from TDSOnSalary where condition='" & Trim(CmbCondition.Text) & "' and firstamount=" & Val(TxtFirstParameter.Text))
    If Recset.EOF = False Then
        MsgBox ("Entry already exist! For this condition.")
        Exit Sub
    End If
    Con.Execute "insert into tdsonsalary values('" & Trim(CmbCondition.Text) & "'," & Val(TxtFirstParameter.Text) & "," & Val(TxtSecondParameter.Text) & "," & Val(TxtValue.Text) & ")"
    MsgBox "Update!", vbExclamation
    ShowRebates
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    CenterMe Me
    ShowRebates
End Sub
Sub ShowRebates()
    Set Recset = Nothing
    Set Recset = Con.Execute("select * from tdsonsalary")
    Grid1.Rows = 1
    Grid1.Rows = 2
    Grid1.FormatString = "Condition|From|To|Value"
    Grid1.Cols = 4
    While Recset.EOF = False
        For i = 0 To 3
            Grid1.TextMatrix(Grid1.Rows - 1, i) = Recset.Fields(i).Value
        Next
        Recset.MoveNext
        Grid1.Rows = Grid1.Rows + 1
    Wend
    Set Recset = Nothing
    Set Recset = Con.Execute("select * from tdsrebatefemale")
    If Recset.EOF = False Then
        Grid1.TextMatrix(Grid1.Rows - 1, 0) = "If Sex=='Female'"
        Grid1.TextMatrix(Grid1.Rows - 1, 1) = Recset.Fields(0).Value
        Grid1.Rows = Grid1.Rows + 1
    End If
    Set Recset = Nothing
    Set Recset = Con.Execute("select * from tdsrebateseniorcitizen")
    If Recset.EOF = False Then
        Grid1.TextMatrix(Grid1.Rows - 1, 0) = "If Age is >="
        For i = 1 To 2
            Grid1.TextMatrix(Grid1.Rows - 1, i) = Recset.Fields(i).Value
        Next
        Grid1.Rows = Grid1.Rows + 1
    End If
    Set Recset = Nothing
    Set Recset = Con.Execute("select * from tdsrebatevarious")
    While Recset.EOF = False
        For i = 0 To 2
            Grid1.TextMatrix(Grid1.Rows - 1, i) = Recset.Fields(i).Value
        Next
        Recset.MoveNext
        Grid1.Rows = Grid1.Rows + 1
    Wend
    Grid1.Rows = Grid1.Rows - 1
    For i = 0 To 3
        Grid1.ColWidth(i) = Grid1.Width / 4.05
    Next
End Sub
