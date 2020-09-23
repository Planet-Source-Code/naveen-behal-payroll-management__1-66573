VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmStaff 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Staff Details"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2775
      Left            =   120
      TabIndex        =   26
      Top             =   4320
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4895
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   4194304
      ForeColorFixed  =   16777215
      WordWrap        =   -1  'True
      MergeCells      =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   3
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   9495
      Begin MSComCtl2.DTPicker DTPLeaved 
         Height          =   315
         Left            =   3360
         TabIndex        =   16
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   24444931
         CurrentDate     =   38985
      End
      Begin VB.CheckBox ChkLeaved 
         BackColor       =   &H00EFEADE&
         Caption         =   "Leaved"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox TxtTDS 
         Height          =   315
         Left            =   6480
         TabIndex        =   25
         Top             =   3240
         Width           =   2775
      End
      Begin VB.ComboBox CmbPriority 
         Height          =   315
         ItemData        =   "FrmStaff.frx":0000
         Left            =   1920
         List            =   "FrmStaff.frx":004F
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox TxtTotal 
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox TxtMedicalReim 
         Height          =   315
         Left            =   6480
         TabIndex        =   24
         Top             =   2880
         Width           =   2775
      End
      Begin VB.ComboBox CmbDesignation 
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   2520
         Width           =   2775
      End
      Begin VB.ComboBox CmBBranchName 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   624
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DTPLASTID 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   25362435
         CurrentDate     =   38939
      End
      Begin MSComCtl2.DTPicker DTPDOJ 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   1395
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   25362435
         CurrentDate     =   38939
      End
      Begin MSComCtl2.DTPicker DTPDOB 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   1005
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   25362435
         CurrentDate     =   38939
      End
      Begin VB.TextBox TxtLastIncrement 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   1776
         Width           =   2775
      End
      Begin VB.TextBox TxtMedical 
         Height          =   315
         Left            =   6480
         TabIndex        =   22
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox TxtPS 
         Height          =   315
         Left            =   6480
         TabIndex        =   21
         Top             =   1776
         Width           =   2775
      End
      Begin VB.TextBox TxtConveyance 
         Height          =   315
         Left            =   6480
         TabIndex        =   20
         Top             =   1392
         Width           =   2775
      End
      Begin VB.TextBox TxtHRA 
         Height          =   315
         Left            =   6480
         TabIndex        =   19
         Top             =   1008
         Width           =   2775
      End
      Begin VB.TextBox TXTPF 
         Height          =   315
         Left            =   6480
         TabIndex        =   18
         Top             =   624
         Width           =   2775
      End
      Begin VB.TextBox TxtBasicSal 
         Height          =   315
         Left            =   6480
         TabIndex        =   17
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox TxtEmpName 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leaving Date"
         Height          =   195
         Left            =   2280
         TabIndex        =   49
         Top             =   3240
         Width           =   960
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TDS"
         Height          =   195
         Left            =   5085
         TabIndex        =   48
         Top             =   3240
         Width           =   330
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Priority in reports"
         Height          =   195
         Left            =   240
         TabIndex        =   47
         Top             =   2880
         Width           =   1155
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   5080
         TabIndex        =   46
         Top             =   2520
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medical Reim."
         Height          =   195
         Left            =   5085
         TabIndex        =   45
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   1008
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joining Date"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   1392
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Increment"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   1776
         Width           =   1050
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Increment Date"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   2160
         Width           =   1440
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Salary"
         Height          =   195
         Left            =   5080
         TabIndex        =   35
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "In lieu of PF"
         Height          =   195
         Left            =   5080
         TabIndex        =   34
         Top             =   630
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HRA"
         Height          =   195
         Left            =   5080
         TabIndex        =   33
         Top             =   1005
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conveyance"
         Height          =   195
         Left            =   5080
         TabIndex        =   32
         Top             =   1395
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PS"
         Height          =   195
         Left            =   5080
         TabIndex        =   31
         Top             =   1770
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medical Allowance"
         Height          =   195
         Left            =   5080
         TabIndex        =   30
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   624
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee name"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EFEADE&
      Height          =   495
      Left            =   120
      TabIndex        =   40
      Top             =   3600
      Width           =   9495
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00EFEADE&
         Caption         =   "E&xit"
         Height          =   315
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   140
         Width           =   1150
      End
      Begin VB.CommandButton CmdFind 
         BackColor       =   &H00EFEADE&
         Caption         =   "&Find"
         Height          =   315
         Left            =   6825
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   140
         Width           =   1150
      End
      Begin VB.CommandButton CmdDelete 
         BackColor       =   &H00EFEADE&
         Caption         =   "&Delete"
         Height          =   315
         Left            =   5490
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   140
         Width           =   1150
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00EFEADE&
         Caption         =   "&Save"
         Height          =   315
         Left            =   4155
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1150
      End
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00EFEADE&
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   140
         Width           =   1150
      End
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H00EFEADE&
         Caption         =   "&Edit"
         Height          =   315
         Left            =   1485
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   140
         Width           =   1150
      End
      Begin VB.CommandButton CmdAdd 
         BackColor       =   &H00EFEADE&
         Caption         =   "&New Record"
         Height          =   315
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   140
         Width           =   1150
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00EFEADE&
      Caption         =   "Show All Branches"
      Height          =   195
      Left            =   7960
      TabIndex        =   42
      Top             =   4110
      Value           =   1  'Checked
      Width           =   1630
   End
   Begin VB.ComboBox CmbBranches 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5160
      TabIndex        =   41
      Text            =   "Branch Names"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Branch Name"
      Height          =   195
      Left            =   4080
      TabIndex        =   43
      Top             =   4110
      Width           =   975
   End
End
Attribute VB_Name = "FrmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MBox As VbMsgBoxResult

Private Sub Check1_Click()
    ShowGrid
End Sub

Private Sub CmbBranches_Change()
    ShowGrid
End Sub

Private Sub CmbBranches_Click()
    ShowGrid
End Sub

Private Sub CmBBranchName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub CmbDesignation_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub CmbPriority_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub CmdAdd_Click()
    BlankMeHere
    OpenCon
    Set Recset = Con.Execute("select distinct(branch) from staffdetails")
    CmBBranchName.Clear
    While Recset.EOF = False
        CmBBranchName.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
    Set Recset = Nothing
    CmbDesignation.Clear
    Set Recset = Con.Execute("select distinct(designation) from staffdetails")
    While Recset.EOF = False
        If IsNull(Recset.Fields("Designation")) = False Then
            CmbDesignation.AddItem Recset.Fields(0).Value
        End If
        Recset.MoveNext
    Wend
    DisableCmdMe Me
    Frame1.Enabled = True
    AddEdit = True
    TxtEmpName.SetFocus
End Sub
Private Sub CmdCancel_Click()
    MBox = MsgBox("Sure?", vbYesNo)
    If MBox = vbYes Then
        BlankMeHere
        EnableCmdMe Me
        CmdAdd.SetFocus
        If Grid.Rows > 1 Then
            If Grid.TextMatrix(1, 1) <> "" Then
                CmdFind.Enabled = True
            Else
                CmdFind.Enabled = False
            End If
        End If
    End If
    Frame1.Enabled = False
End Sub
Private Sub CmdDelete_Click()
    MBox = MsgBox("Sure?", vbYesNo)
    If MBox = vbYes Then
        Con.Execute ("delete from staffdetails where staffid='" & Trim(TxtEmpName.Text) & "' and branch='" & Trim(CmBBranchName.Text) & "'")
        BlankMeHere
        DisableCmdMe Me
        EnableCmdMe Me
        CmdAdd.SetFocus
        ShowGrid
    End If
End Sub

Private Sub CmdEdit_Click()
    If Trim(TxtEmpName.Text) = "" Then
        MsgBox "First select a record to modify", vbInformation
        Exit Sub
    End If
    If Trim(CmBBranchName.Text) = "" Then
        MsgBox "First select a record to modify", vbInformation
        Exit Sub
    End If
    TxtEmpName.Tag = Trim(TxtEmpName.Text)
    CmBBranchName.Tag = Trim(CmBBranchName.Text)
    AddEdit = False
    DisableCmdMe Me
    Frame1.Enabled = True
    TxtEmpName.SetFocus
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdFind_Click()
    MsgBox "Please select from grid"
    Grid.SetFocus
End Sub

Private Sub CmdSave_Click()
    If Trim(TxtEmpName.Text) = "" Then
        MsgBox "Please enter employee id", vbInformation
        TxtEmpName.SetFocus
        Exit Sub
    End If
    If Trim(CmBBranchName.Text) = "" Then
        MsgBox "Please select or enter branch id", vbInformation
        CmBBranchName.SetFocus
        Exit Sub
    End If
    If Val(TxtLastIncrement.Text) < 0 Then
        MsgBox "Please enter valid increment value"
        TxtLastIncrement.SetFocus
        Exit Sub
    End If
    If DateDiff("D", DTPDOJ.Value, DTPLASTID.Value) < 0 And Val(TxtLastIncrement.Text) > 0 Then
        MsgBox "Please check Joining Date against Last increment date", vbInformation
        DTPDOJ.SetFocus
        Exit Sub
    End If
    If Trim(CmbDesignation.Text) = "" Then
        MsgBox "Please choose or enter designation", vbInformation
        CmbDesignation.SetFocus
        Exit Sub
    End If
    If Trim(CmbPriority.Text) = "" Then
        MsgBox "Please select a priority level for the employee." & vbCrLf & "It will sort the employee names in reports"
        CmbPriority.SetFocus
        Exit Sub
    End If
    If ChkLeaved.Value = vbChecked Then
        l = 1
    Else
        l = 0
    End If
    If Val(TxtBasicSal.Text) <= 0 Then
        MsgBox "Please enter valid basic salary", vbInformation
        TxtBasicSal.SetFocus
        Exit Sub
    End If
    OpenCon
    If AddEdit = True Then
        Set Recset = Con.Execute("select * from staffdetails where staffid='" & Trim(TxtEmpName.Text) & "' and branch='" & Trim(CmBBranchName.Text) & "'")
        If Recset.EOF = False Then
            MsgBox "Sorry Sir! But Staff Information already exist for this ID", vbInformation
            TxtEmpName.SetFocus
            Exit Sub
        End If
        Con.Execute "insert into staffdetails values('" & Trim(TxtEmpName.Text) & "'," & Val(TxtBasicSal.Text) & ",'" & Trim(CmBBranchName.Text) & "','" & DTPDOJ.Value & "','" & DTPDOB.Value & "'," & Val(TXTPF.Text) & "," & Val(TxtHRA.Text) & "," & Val(TxtConveyance.Text) & "," & Val(TxtPS.Text) & "," & Val(TxtMedical.Text) & "," & Val(TxtLastIncrement.Text) & ",'" & DTPLASTID.Value & "'," & Val(TxtMedicalReim.Text) & ",'" & Trim(CmbDesignation.Text) & "'," & Val(CmbPriority.Text) & "," & Val(TxtTDS.Text) & ",'" & l & "','" & DTPLeaved.Value & "')"
    Else
        Con.Execute "update staffdetails set staffid='" & Trim(TxtEmpName.Text) & "', basicsalary=" & Val(TxtBasicSal.Text) & ", branch='" & Trim(CmBBranchName.Text) & "', doj='" & DTPDOJ.Value & "',  dob='" & DTPDOB.Value & "', pf=" & Val(TXTPF.Text) & ", hra=" & Val(TxtHRA.Text) & ", conveyance=" & Val(TxtConveyance.Text) & ", ps=" & Val(TxtPS.Text) & ", medical=" & Val(TxtMedical.Text) & ", lastincrement='" & Trim(TxtLastIncrement.Text) & "', doli='" & DTPLASTID.Value & "', medicalreim=" & Val(TxtMedicalReim.Text) & ",designation='" & Trim(CmbDesignation.Text) & "',reportpriority=" & Val(CmbPriority.Text) & ",tds=" & Val(TxtTDS.Text) & ",leaved='" & l & "',leavingdate='" & DTPLeaved.Value & "'  where staffid='" & TxtEmpName.Tag & "' and branch='" & CmBBranchName.Tag & "'"
        Con.Execute "update salarysheet set reportpriority='" & Trim(CmbPriority.Text) & "' where staffid='" & TxtEmpName.Tag & "' and branch='" & CmBBranchName.Tag & "'"
    End If
    EnableCmdMe Me
    CmdEdit.Enabled = False
    CmdFind.Enabled = True
    CmdDelete.Enabled = True
    ShowGrid
    Frame1.Enabled = False
End Sub
Private Sub DTPDOB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub DTPDOB_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbEnter Then
        SendKeys ("{Tab}")
    End If
End Sub

Private Sub DTPDOJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub DTPLASTID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub Form_Resize()
    CenterMe Me
    DisableCmdMe Me
    EnableCmdMe Me
    ShowGrid
    Frame1.Enabled = False
    CmdAdd.SetFocus
    OpenCon
    Set Recset = Con.Execute("select distinct(branch) from staffdetails")
    While Recset.EOF = False
        CmbBranches.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
    Set Recset = Nothing
    Set Recset = Con.Execute("select distinct(designation) from staffdetails")
    While Recset.EOF = False
        If IsNull(Recset.Fields(0).Value) = False Then
            CmbDesignation.AddItem Recset.Fields(0).Value
        End If
        Recset.MoveNext
    Wend
End Sub
Sub BlankMeHere()
    TxtEmpName.Text = ""
    TxtLastIncrement.Text = ""
    TxtBasicSal.Text = ""
    TXTPF.Text = ""
    TxtHRA.Text = ""
    TxtConveyance.Text = ""
    TxtPS.Text = ""
    TxtMedical.Text = ""
    TxtMedicalReim.Text = ""
    CmbDesignation.Text = ""
    TxtTotal.Text = ""
    TxtTDS.Text = ""
End Sub
Sub ShowFields()
    TxtEmpName.Text = Recset.Fields("staffid").Value
    TxtEmpName.Tag = TxtEmpName.Text
    TxtBasicSal.Text = Recset.Fields("basicsalary").Value
    CmBBranchName.Text = Recset.Fields("Branch").Value
    CmBBranchName.Tag = CmBBranchName.Tag
    DTPDOB.Value = Recset.Fields("dob").Value
    DTPDOJ = Recset.Fields("doj").Value
    TxtLastIncrement.Text = Recset.Fields("lastincrement").Value
    DTPLASTID.Value = Recset.Fields("DOLI").Value
    TXTPF.Text = Recset.Fields("PF").Value
    TxtHRA.Text = Recset.Fields("HRA").Value
    TxtConveyance.Text = Recset.Fields("conveyance").Value
    TxtPS.Text = Recset.Fields("ps").Value
    TxtMedical.Text = Recset.Fields("medical").Value
    CmbDesignation.Text = Recset.Fields("designation").Value
    CmbPriority.ListIndex = Recset.Fields("reportPriority").Value - 1
    TxtMedicalReim.Text = Recset.Fields("medicalreim").Value
    TxtTDS.Text = Recset.Fields("TDS").Value
    If Recset.Fields("Leaved").Value = "1" Then
        ChkLeaved.Value = vbChecked
    Else
        ChkLeaved.Value = vbUnchecked
    End If
    DTPLeaved.Value = Recset.Fields("leavingDate").Value
End Sub
Sub ShowGrid()
    OpenCon
    If Check1.Value = vbChecked Then
        Set Recset = Con.Execute("Select branch,staffid,Designation,basicsalary,pf,hra,conveyance,ps,medical,basicsalary + pf + hra + conveyance + ps + medical,medicalreim,tds,lastincrement,doli,dob,doj from staffdetails order by branch, reportpriority")
    Else
        Set Recset = Con.Execute("Select branch,staffid,Designation,basicsalary,pf,hra,conveyance,ps,medical,basicsalary + pf + hra + conveyance + ps + medical,medicalreim,tds,lastincrement,doli,dob,doj from staffdetails where branch='" & Trim(CmbBranches.Text) & "' order by reportpriority")
    End If
    If Recset.EOF = False Then
        CmdFind.Enabled = True
    End If
    i = 1
    Grid.Cols = 16
    Grid.Rows = 2
    Grid.FormatString = "Branch|Employee|Designation|Basic|PF|HRA|Coveyance|PS|Medical|Total|Medical Reim.|TDS|Last Incr.|Last Incr. Date|Birth Date|Joining Date"
    While Recset.EOF = False
        For i = 0 To 15
            Grid.TextMatrix(Grid.Rows - 1, i) = Recset.Fields(i).Value
        Next
        Recset.MoveNext
        Grid.Rows = Grid.Rows + 1
    Wend
    Grid.Rows = Grid.Rows - 1
    Grid.MergeCol(0) = True
    For i = 0 To 15
        Grid.ColWidth(i) = Grid.Width / 10
    Next
    Grid.ColWidth(1) = Grid.Width / 4
    Grid.ColWidth(3) = Grid.Width / 16
    Grid.ColWidth(4) = Grid.Width / 16
    Grid.ColWidth(5) = Grid.Width / 16
    Grid.ColWidth(7) = Grid.Width / 16
    Grid.ColWidth(8) = Grid.Width / 14
End Sub

Private Sub Grid_Click()
    If Grid.TextMatrix(Grid.Row, 0) <> "" Then
        OpenCon
        Set Recset = Con.Execute("select * from staffdetails where staffid='" & Grid.TextMatrix(Grid.Row, 1) & "' and branch='" & Grid.TextMatrix(Grid.Row, 0) & "'")
        If Recset.EOF = True Then
            Exit Sub
        End If
        ShowFields
        CmdEdit.Enabled = True
        CmdDelete.Enabled = True
    End If
End Sub
Private Sub Grid_EnterCell()
    If Grid.TextMatrix(Grid.Row, 0) <> "" Then
        OpenCon
        Set Recset = Con.Execute("select * from staffdetails where staffid='" & Grid.TextMatrix(Grid.Row, 1) & "' and branch='" & Grid.TextMatrix(Grid.Row, 0) & "'")
        If Recset.EOF = True Then
            Exit Sub
        End If
        ShowFields
        CmdEdit.Enabled = True
        CmdDelete.Enabled = True
    End If
End Sub

Private Sub TxtBasicSal_Change()
    CalculateTotal
End Sub

Private Sub TxtBasicSal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub TxtConveyance_Change()
    CalculateTotal
End Sub

Private Sub TxtConveyance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub TxtEmpName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub TxtHRA_Change()
    CalculateTotal
End Sub

Private Sub TxtHRA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub TxtLastIncrement_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub TxtMedical_Change()
    CalculateTotal
End Sub

Private Sub TxtMedical_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
        
End Sub

Private Sub TxtMedicalReim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 14 Then SendKeys ("{Tab}")
End Sub

Private Sub TXTPF_Change()
    CalculateTotal
End Sub

Private Sub TXTPF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub TxtPS_Change()
    CalculateTotal
End Sub

Private Sub TxtPS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub
Sub CalculateTotal()
    TxtTotal.Text = Val(TxtBasicSal.Text) + Val(TXTPF.Text) + Val(TxtMedical.Text) + Val(TxtPS.Text) + Val(TxtHRA.Text) + Val(TxtConveyance.Text)
End Sub
Private Sub TxtTDS_GotFocus()
    If CmdSave.Enabled = True Then CmdSave.Default = True
End Sub
