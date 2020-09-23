VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmSalary 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Salary Sheet"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   34
      Top             =   5
      Width           =   9375
      Begin VB.TextBox TxtMM 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8760
         TabIndex        =   14
         Top             =   650
         Width           =   495
      End
      Begin VB.TextBox TxtPaidLeave 
         Height          =   285
         Left            =   7920
         TabIndex        =   13
         Top             =   650
         Width           =   510
      End
      Begin VB.TextBox TxtWorkingDays 
         Height          =   285
         Left            =   6600
         TabIndex        =   12
         Top             =   650
         Width           =   495
      End
      Begin MSComCtl2.DTPicker DTPPayDate 
         Height          =   315
         Left            =   8160
         TabIndex        =   10
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   38939
      End
      Begin VB.ComboBox CmbMonth 
         Height          =   315
         ItemData        =   "FrmSalarySheet.frx":0000
         Left            =   6600
         List            =   "FrmSalarySheet.frx":0028
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox CmbYear 
         Height          =   315
         ItemData        =   "FrmSalarySheet.frx":008E
         Left            =   4680
         List            =   "FrmSalarySheet.frx":00B6
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox CmbEmpName 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   650
         Width           =   3975
      End
      Begin VB.ComboBox CmBBranchName 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MM"
         Height          =   195
         Left            =   8430
         TabIndex        =   54
         Top             =   650
         Width           =   270
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paid Leave"
         Height          =   195
         Left            =   7080
         TabIndex        =   53
         Top             =   650
         Width           =   810
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Working Days"
         Height          =   195
         Left            =   5400
         TabIndex        =   52
         Top             =   650
         Width           =   1005
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   7800
         TabIndex        =   51
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   50
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         Height          =   315
         Left            =   6045
         TabIndex        =   49
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee name"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   650
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   9375
      Begin VB.TextBox TxtLastIncrement 
         Height          =   315
         Left            =   1680
         TabIndex        =   56
         Top             =   575
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DTPLASTID 
         Height          =   315
         Left            =   6480
         TabIndex        =   57
         Top             =   575
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   24510467
         CurrentDate     =   38939
      End
      Begin MSComCtl2.DTPicker DTPDOJ 
         Height          =   315
         Left            =   6480
         TabIndex        =   58
         Top             =   200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   24510467
         CurrentDate     =   38939
      End
      Begin MSComCtl2.DTPicker DTPDOB 
         Height          =   315
         Left            =   1680
         TabIndex        =   59
         Top             =   200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   24510467
         CurrentDate     =   38939
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Increment Date"
         Height          =   195
         Left            =   4920
         TabIndex        =   32
         Top             =   570
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Increment"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   570
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joining Date"
         Height          =   255
         Left            =   4920
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2775
      Left            =   120
      TabIndex        =   27
      Top             =   4320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4895
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   65535
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
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   37
      Top             =   2040
      Width           =   9375
      Begin VB.TextBox TxtTotal 
         Height          =   285
         Left            =   7920
         TabIndex        =   61
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtSalary 
         Height          =   285
         Left            =   3120
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtMedical 
         Height          =   285
         Left            =   5280
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtPS 
         Height          =   285
         Left            =   3120
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtConveyance 
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtHRA 
         Height          =   285
         Left            =   7920
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TXTPF 
         Height          =   285
         Left            =   5280
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtBasicSal 
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   6840
         TabIndex        =   62
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
         Height          =   195
         Left            =   2640
         TabIndex        =   55
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Salary"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "In Lieu Of PF"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   4560
         TabIndex        =   42
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HRA"
         Height          =   195
         Left            =   6840
         TabIndex        =   41
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conveyance"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PS"
         Height          =   195
         Left            =   2640
         TabIndex        =   39
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medical"
         Height          =   195
         Left            =   4560
         TabIndex        =   38
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   44
      Top             =   3120
      Width           =   9375
      Begin VB.TextBox TxtOther 
         Height          =   285
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtTDS 
         Height          =   285
         Left            =   1200
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtPF1 
         Height          =   285
         Left            =   2400
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtDeductions 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtGrandTotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7920
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Others"
         Height          =   195
         Left            =   3240
         TabIndex        =   60
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TDS"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Ded."
         Height          =   195
         Left            =   4560
         TabIndex        =   47
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PF"
         Height          =   195
         Left            =   2040
         TabIndex        =   46
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
         Height          =   195
         Left            =   6840
         TabIndex        =   45
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   33
      Top             =   3720
      Width           =   9375
      Begin VB.CommandButton CmdAdd 
         BackColor       =   &H00EFEADE&
         Caption         =   "&New Record"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   140
         Width           =   1150
      End
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H00EFEADE&
         Caption         =   "&Edit"
         Height          =   315
         Left            =   1450
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   140
         Width           =   1150
      End
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00EFEADE&
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   2780
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   140
         Width           =   1150
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00EFEADE&
         Caption         =   "&Save"
         Height          =   315
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   140
         Width           =   1150
      End
      Begin VB.CommandButton CmdDelete 
         BackColor       =   &H00EFEADE&
         Caption         =   "&Delete"
         Height          =   315
         Left            =   5440
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   140
         Width           =   1150
      End
      Begin VB.CommandButton CmdFind 
         BackColor       =   &H00EFEADE&
         Caption         =   "&Find"
         Height          =   315
         Left            =   6770
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   140
         Width           =   1150
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00EFEADE&
         Caption         =   "E&xit"
         Height          =   315
         Left            =   8105
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   140
         Width           =   1150
      End
   End
End
Attribute VB_Name = "FrmSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MBox As VbMsgBoxResult, mno As Integer
Private Sub CmbBranchName_Change()
    CmbEmpName.Clear
    Set Recset = Nothing
    If Trim(CmbBranchName.Text) = "" Then Exit Sub
    Set Recset = Con.Execute("select staffid from staffdetails where branch='" & Trim(CmbBranchName.Text) & "' and leaved='0'")
    While Recset.EOF = False
        CmbEmpName.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
End Sub
Private Sub CmbBranchName_Click()
    CmbEmpName.Clear
    Set Recset = Nothing
    If Trim(CmbBranchName.Text) = "" Then Exit Sub
    Set Recset = Con.Execute("select staffid from staffdetails where branch='" & Trim(CmbBranchName.Text) & "' and leaved='0'")
    While Recset.EOF = False
        CmbEmpName.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
End Sub

Private Sub CmBBranchName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub CmbEmpName_Change()
    If Trim(CmbEmpName.Text) = "" Then Exit Sub
    Set Recset = Con.Execute("select * from staffdetails where staffid='" & Trim(CmbEmpName.Text) & "' and branch='" & Trim(CmbBranchName.Text) & "'")
    If Recset.EOF = False Then
        ShowDetails
    Else
        BlankMeHere
    End If
End Sub
Private Sub CmbEmpName_Click()
    Set Recset = Nothing
    If Trim(CmbEmpName.Text) = "" Then Exit Sub
    Set Recset = Con.Execute("select * from staffdetails where staffid='" & Trim(CmbEmpName.Text) & "' and branch='" & Trim(CmbBranchName.Text) & "'")
    If Recset.EOF = False Then
        ShowDetails
    End If
End Sub

Private Sub CmbEmpName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub CmbMonth_Change()
    ShowGrid
End Sub

Private Sub CmbMonth_Click()
    ShowGrid
End Sub

Private Sub CmbMonth_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub CmbYear_Change()
    ShowGrid
End Sub

Private Sub CmbYear_Click()
    ShowGrid
End Sub

Private Sub CmbYear_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub CmdAdd_Click()
    BlankMeHere
    DisableCmdMe Me
    Frame1.Enabled = True
    CmbEmpName.Text = ""
    CmbEmpName.SetFocus
End Sub
Private Sub CmdCancel_Click()
    BlankMeHere
    EnableCmdMe Me
End Sub

Private Sub CmdDelete_Click()
    MBox = MsgBox("Sure?", vbYesNo)
    If MBox = vbYes Then
        Con.Execute "delete from salarysheet where staffid='" & Trim(CmbEmpName.Text) & "' and salaryyear='" & Trim(CmbYear.Text) & "' and salarymonth='" & Trim(CmbMonth.Text) & "' and branch='" & Trim(CmbBranchName.Text) & "'"
        BlankMeHere
        CmdDelete.Enabled = False
        ShowGrid
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdFind_Click()
    MsgBox "Please select from the grid", vbInformation
End Sub

Private Sub CmdSave_Click()
    If Trim(CmbEmpName.Text) = "" Then
        MsgBox "Please select valid employee name.", vbInformation
        CmbEmpName.SetFocus
        Exit Sub
    End If
    If Trim(CmbMonth.Text) = "January" Or Trim(CmbMonth.Text) = "February" Or Trim(CmbMonth.Text) = "March" Or Trim(CmbMonth.Text) = "April" Or Trim(CmbMonth.Text) = "May" Or Trim(CmbMonth.Text) = "June" Or Trim(CmbMonth.Text) = "July" Or Trim(CmbMonth.Text) = "August" Or Trim(CmbMonth.Text) = "September" Or Trim(CmbMonth.Text) = "October" Or Trim(CmbMonth.Text) = "November" Or Trim(CmbMonth.Text) = "December" Then
    Else
        MsgBox "Please select valid Month", vbInformation
        CmbMonth.SetFocus
        Exit Sub
    End If
    If Val(TxtMM.Text) > 1 Then
        MsgBox "Please enter valid attendance", vbInformation
        TxtWorkingDays.SetFocus
        Exit Sub
    End If
    Select Case Trim(CmbMonth.Text)
        Case "April"
            mno = 1
        Case "May"
            mno = 2
        Case "June"
            mno = 3
        Case "July"
            mno = 5
        Case "August"
            mno = 6
        Case "September"
            mno = 7
        Case "October"
            mno = 8
        Case "November"
            mno = 9
        Case "December"
            mno = 10
        Case "January"
            mno = 11
        Case "February"
            mno = 12
        Case "March"
            mno = 13
    End Select
    If Val(TxtMM.Text) > 1 Then
        MsgBox "Please enter valid attendance", vbInformation
        TxtWorkingDays.SetFocus
        Exit Sub
    End If
    If Val(TxtBasicSal.Text) <= 0 Then
        MsgBox "Please select valid Branch or Employee Name", vbInformation
        CmbBranchName.SetFocus
        Exit Sub
    End If
    If Val(TxtTDS.Text) < 0 Or Val(TxtSalary.Text) < Val(TxtTDS.Text) Then
        MsgBox "Please Enter valid TDS", vbInformation
        TxtTDS.SetFocus
    End If
    Set Recset = Nothing
    Set Recset = Con.Execute("select * from salarysheet where staffid='" & Trim(CmbEmpName.Text) & "' and salarymonth='" & Trim(CmbMonth.Text) & "' and salaryyear='" & Trim(CmbYear.Text) & "' and branch='" & Trim(CmbBranchName.Text) & "'")
    If Recset.EOF = False Then
        MsgBox "Sorry Sir! But entery already exist for this staff member of this month", vbInformation
        Exit Sub
    End If
    Set TmpRecset = Con.Execute("select reportpriority from staffdetails where staffid='" & Trim(CmbEmpName.Text) & "' and branch='" & Trim(CmbBranchName.Text) & "'")
    If TmpRecset.EOF = True Then
        stp = 1
    Else
        stp = TmpRecset.Fields(0).Value
    End If
    Con.Execute "insert into salarysheet values('" & Trim(CmbEmpName.Text) & "','" & Trim(CmbMonth.Text) & "','" & Trim(CmbYear.Text) & "','" & DTPPayDate.Value & "'," & Val(TxtWorkingDays.Text) & "," & Val(TxtPaidLeave.Text) & "," & Round(Val(TxtMM.Text), 6) & "," & Val(TXTPF.Text) & "," & Val(TxtHRA.Text) & "," & Val(TxtConveyance.Text) & "," & Val(TxtPS.Text) & "," & Val(TxtMedical.Text) & "," & Val(TxtTotal.Text) & "," & Val(TxtBasicSal.Text) & "," & Val(TxtSalary.Text) & ",'" & Trim(CmbBranchName.Text) & "'," & Val(TxtTDS.Text) & "," & Val(TxtPF1.Text) & "," & mno & "," & Val(TxtLastIncrement.Text) & ",'" & DTPLASTID.Value & "'," & Val(TxtOther.Text) & "," & stp & ")"
    EnableCmdMe Me
    CmdEdit.Enabled = True
    CmdDelete.Enabled = True
    CmdFind.Enabled = True
    ShowGrid
End Sub

Private Sub DTPPayDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub Form_Load()
    CenterMe Me
    OpenCon
    Set Recset = Nothing
    Set Recset = Con.Execute("select distinct(branch) from staffdetails")
    CmbBranchName.Clear
    While Recset.EOF = False
        CmbBranchName.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
    ShowGrid
    EnableCmdMe Me
End Sub
Sub ShowDetails()
    TxtBasicSal.Text = Recset.Fields("BasicSalary").Value
    TxtBasicSal.Tag = Recset.Fields("BasicSalary").Value
    DTPDOJ.Value = Recset.Fields("doj").Value
    DTPDOB.Value = Recset.Fields("dob").Value
    TXTPF.Text = Recset.Fields("Pf").Value
    TXTPF.Tag = Recset.Fields("Pf").Value
    TxtHRA.Text = Recset.Fields("hra").Value
    TxtHRA.Tag = Recset.Fields("hra").Value
    TxtConveyance.Text = Recset.Fields("conveyance").Value
    TxtConveyance.Tag = Recset.Fields("conveyance").Value
    TxtPS.Text = Recset.Fields("PS").Value
    TxtPS.Tag = Recset.Fields("PS").Value
    TxtMedical.Text = Recset.Fields("medical").Value
    TxtMedical.Tag = Recset.Fields("medical").Value
    TxtLastIncrement.Text = Recset.Fields("LastIncrement").Value
    DTPLASTID.Value = Recset.Fields("Doli").Value
    TxtTDS.Text = Recset.Fields("Tds").Value
    CmdEdit.Enabled = True
    CmdDelete.Enabled = True
End Sub
Sub Calculate()
    TxtTotal.Text = Val(TxtBasicSal.Text) + Val(TXTPF.Text) + Val(TxtHRA.Text) + Val(TxtConveyance.Text) + Val(TxtPS.Text) + Val(TxtMedical.Text)
    If Val(TxtBasicSal.Text) <= 6500 Then
        If Val(TXTPF.Text) <= 0 Then
            TxtFinalPF.Text = Round(Val(TxtBasicSal.Text) * 0.12, 0)
        Else
            TxtFinalPF.Text = 0
        End If
    Else
        TxtFinalPF.Text = 0
    End If
End Sub

Private Sub Grid_Click()
    If Grid.TextMatrix(Grid.Row, 0) <> "" Then
        Set TmpRecset = Nothing
        Set TmpRecset = Con.Execute("select * from salarysheet where branch='" & Trim(CmbBranchName.Text) & "' and staffid='" & Grid.TextMatrix(Grid.Row, 0) & "' and salaryyear='" & Trim(CmbYear.Text) & "' and salarymonth='" & Trim(CmbMonth.Text) & "'")
        If TmpRecset.EOF = True Then Exit Sub
        ShowSalary
        CmdDelete.Enabled = True
        CmdEdit.Enabled = True
    End If
End Sub

Private Sub Grid_EnterCell()
   If Grid.TextMatrix(Grid.Row, 0) <> "" Then
        Set TmpRecset = Nothing
        Set TmpRecset = Con.Execute("select * from salarysheet where branch='" & Trim(CmbBranchName.Text) & "' and staffid='" & Grid.TextMatrix(Grid.Row, 0) & "' and salaryyear='" & Trim(CmbYear.Text) & "' and salarymonth='" & Trim(CmbMonth.Text) & "'")
        If TmpRecset.EOF = True Then Exit Sub
        ShowSalary
        CmdDelete.Enabled = True
        CmdEdit.Enabled = True
    End If
End Sub

Private Sub TxtMM_Change()
    CalculateSal
End Sub

Private Sub TxtOther_Change()
    TxtDeductions.Text = Round(Val(TxtTDS.Text) + Val(TxtPF1.Text) + Val(TxtOther.Text), 0)
    TxtGrandTotal.Text = Round(Val(TxtTotal.Text) - Val(TxtDeductions.Text), 0)
End Sub

Private Sub TxtOther_GotFocus()
    If CmdSave.Enabled = True Then CmdSave.Default = True
End Sub
Private Sub TxtPaidLeave_Change()
    If Trim(CmbMonth.Text) = "February" Then
        If Val(CmbYear.Text) Mod 4 = 0 Then
            dayn = 29
        Else
            dayn = 28
        End If
    ElseIf Trim(CmbMonth.Text) = "January" Or Trim(CmbMonth.Text) = "March" Or Trim(CmbMonth.Text) = "May" Or Trim(CmbMonth.Text) = "July" Or Trim(CmbMonth.Text) = "August" Or Trim(CmbMonth.Text) = "October" Or Trim(CmbMonth.Text) = "December" Then
        dayn = 31
    Else
        dayn = 30
    End If
    TxtMM.Text = Round((Val(TxtWorkingDays.Text) + Val(TxtPaidLeave.Text)) / dayn, 6)
End Sub

Private Sub TxtPaidLeave_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub TxtPF1_Change()
        TxtDeductions.Text = Round(Val(TxtTDS.Text) + Val(TxtPF1.Text) + Val(TxtOther.Text), 0)
        TxtGrandTotal.Text = Round(Val(TxtTotal.Text) - Val(TxtDeductions.Text), 0)
End Sub

Private Sub TxtPF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub TxtTDS_Change()
    TxtDeductions.Text = Round(Val(TxtTDS.Text) + Val(TxtPF1.Text) + Val(TxtOther.Text), 0)
    TxtGrandTotal.Text = Round(Val(TxtTotal.Text) - Val(TxtDeductions.Text), 0)
End Sub
Private Sub TxtTDS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub
Private Sub TxtTotal_Change()
    TxtTDS_Change
End Sub

Private Sub TxtWorkingDays_Change()
    If Trim(CmbMonth.Text) = "February" Then
        If Val(CmbYear.Text) Mod 4 = 0 Then
            dayn = 29
        Else
            dayn = 28
        End If
    ElseIf Trim(CmbMonth.Text) = "January" Or Trim(CmbMonth.Text) = "March" Or Trim(CmbMonth.Text) = "May" Or Trim(CmbMonth.Text) = "July" Or Trim(CmbMonth.Text) = "August" Or Trim(CmbMonth.Text) = "October" Or Trim(CmbMonth.Text) = "December" Then
        dayn = 31
    Else
        dayn = 30
    End If
    TxtMM.Text = Round((Val(TxtWorkingDays.Text) + Val(TxtPaidLeave.Text)) / dayn, 6)
End Sub
Sub CalculateSal()
    TxtSalary.Text = Round(Val(TxtBasicSal.Text) * Val(TxtMM.Text), 0)
    TXTPF.Text = Round(Val(TXTPF.Tag) * Val(TxtMM.Text), 0)
    TxtHRA.Text = Round(Val(TxtHRA.Tag) * Val(TxtMM.Text), 0)
    TxtConveyance.Text = Round(Val(TxtConveyance.Tag) * Val(TxtMM.Text), 0)
    TxtPS.Text = Round(Val(TxtPS.Tag) * Val(TxtMM.Text), 0)
    TxtMedical.Text = Round(Val(TxtMedical.Tag) * Val(TxtMM.Text), 0)
    TxtTotal.Text = Round(Val(TxtSalary.Text) + Val(TXTPF.Text) + Val(TxtHRA.Text) + Val(TxtConveyance.Text) + Val(TxtPS.Text) + Val(TxtMedical.Text), 0)
    If Val(TxtBasicSal.Text) <= 6500 Then
        If Val(TXTPF.Text) <= 0 Then
            TxtPF1.Text = Round(Val(TxtSalary.Text) * 0.12, 0)
        Else
            TxtPF1.Text = 0
        End If
    Else
        TxtPF1.Text = 0
    End If
End Sub
Sub BlankMeHere()
'    CmbYear.Text = ""
 '   CmbMonth.Text = ""
    TxtWorkingDays.Text = ""
    TxtPaidLeave.Text = ""
    TxtMM.Text = ""
    TxtLastIncrement.Text = ""
    TxtBasicSal.Text = ""
    TxtBasicSal.Tag = ""
    TxtSalary.Text = ""
    TxtSalary.Tag = ""
    TXTPF.Text = ""
    TXTPF.Tag = ""
    TxtHRA.Text = ""
    TxtHRA.Tag = ""
    TxtConveyance.Text = ""
    TxtConveyance.Tag = ""
    TxtPS.Text = ""
    TxtPS.Tag = ""
    TxtMedical.Text = ""
    TxtMedical.Tag = ""
    TxtTotal.Text = ""
    TxtPF1.Text = ""
    TxtTDS.Text = ""
    TxtOther.Text = ""
End Sub
Sub ShowGrid()
    Set Recset = Nothing
    Set Recset = Con.Execute("Select staffid,salarydate,salary,workingdays,paidleaves,mm,inlieuofpf,hra,conveyance,projectallowance,medical,total,tds,pf,otherded,tds+pf+otherded,total-tds-pf-otherded from salarysheet where salaryyear='" & Trim(CmbYear.Text) & "' and salarymonth='" & Trim(CmbMonth.Text) & "' and branch='" & Trim(CmbBranchName.Text) & "'")
    Grid.Cols = 17
    Grid.Rows = 2
    Grid.FormatString = "Employee Name|Salary Date|B.Salary|W.Days|P. Leave|MM|In Lieu OF PF|HRA|Conveyance|PS|Medical|Total|TDS|PF|Total Deductions|Grand Total"
    i = 2
    While Recset.EOF = False
        For j = 0 To 16
            Grid.TextMatrix(i - 1, j) = Recset.Fields(j).Value
        Next
        i = i + 1
        Grid.Rows = i
        Recset.MoveNext
    Wend
    Grid.Rows = Grid.Rows - 1
    Grid.ColWidth(0) = Grid.Width / 4
    For i = 1 To 16
        Grid.ColWidth(i) = Grid.Width / 6
    Next
End Sub
Sub ShowSalary()
    CmbBranchName.Text = TmpRecset.Fields("branch").Value
    CmbEmpName.Text = TmpRecset.Fields("staffid").Value
    DTPPayDate.Value = TmpRecset.Fields("salarydate").Value
    TxtWorkingDays = TmpRecset.Fields("workingdays").Value
    TxtPaidLeave.Text = TmpRecset.Fields("paidleaves").Value
    TxtBasicSal.Text = TmpRecset.Fields("basicsal").Value
    TxtTDS.Text = TmpRecset.Fields("tds").Value
    TxtOther.Text = TmpRecset.Fields("otherded").Value
End Sub

Private Sub TxtWorkingDays_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub
