VERSION 5.00
Begin VB.Form Dialog1 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salary Slip"
   ClientHeight    =   2595
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   4215
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EFEADE&
         Caption         =   "Hide Bottum Labels"
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   2100
         Width           =   1815
      End
      Begin VB.ComboBox CmbBranchName 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   3255
      End
      Begin VB.ComboBox CmbMonth 
         Height          =   315
         ItemData        =   "Salary Slip.frx":0000
         Left            =   840
         List            =   "Salary Slip.frx":0028
         TabIndex        =   3
         Top             =   1725
         Width           =   3255
      End
      Begin VB.ComboBox CmbYear 
         Height          =   315
         ItemData        =   "Salary Slip.frx":008E
         Left            =   840
         List            =   "Salary Slip.frx":00B6
         TabIndex        =   2
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox CmbEmployee 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   450
      End
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub CmBBranchName_Change()
    OpenCon
    Set Recset = Con.Execute("select distinct(Staffid) from salarysheet where branch='" & Trim(CmBBranchName.Text) & "'")
    CmbEmployee.Clear
    While Recset.EOF = False
        CmbEmployee.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
End Sub

Private Sub CmBBranchName_Click()
    OpenCon
    Set Recset = Con.Execute("select distinct(Staffid) from salarysheet where branch='" & Trim(CmBBranchName.Text) & "'")
    CmbEmployee.Clear
    While Recset.EOF = False
        CmbEmployee.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend

End Sub

Private Sub Form_Load()
    CenterMe Me
    OpenCon
    Set Recset = Con.Execute("select distinct(Branch) from salarysheet")
    While Recset.EOF = False
        CmBBranchName.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
End Sub

Private Sub Form_Paint()
    If Me.Caption <> "Salary Slip" Then
        CmbMonth.Visible = False
        Label3.Visible = False
    End If
End Sub

Private Sub OKButton_Click()
Dim mno As Integer
Set DataEnvironment1 = Nothing
    mno = 0
    mno = MonthNo(Trim(CmbMonth.Text))
    If Me.Caption = "Salary Slip" Then
        If mno = 0 Then
            MsgBox "Please select valid month name", vbInformation
            CmbMonth.SetFocus
            Exit Sub
        End If
    End If
    Set Recset = Nothing
    If Trim(CmbEmployee.Text) = "" Then
        MsgBox "Please choose an employee name", vbInformation
        CmbEmployee.SetFocus
        Exit Sub
    End If
    If Trim(CmbYear.Text) = "" Then
        MsgBox "Please select a year", vbInformation
        CmbYear.SetFocus
        Exit Sub
    End If
    If Me.Caption = "Salary Slip" Then
        If Trim(CmbMonth.Text) = "" Then
            MsgBox "Please select a month name", vbInformation
            CmbMonth.SetFocus
            Exit Sub
        End If
        Set Recset = Con.Execute("select * from salarysheet where staffid='" & Trim(CmbEmployee.Text) & "' and salaryyear='" & Trim(CmbYear.Text) & "' and salarymonth='" & Trim(CmbMonth.Text) & "' and branch='" & Trim(CmBBranchName.Text) & "'")
        If Recset.EOF = True Then
            MsgBox "No record found!", vbInformation
            Exit Sub
        End If
        DataEnvironment1.SalarySlip Trim(CmbEmployee.Text), mno, Trim(CmbYear.Text), Trim(CmbMonth.Text), Trim(CmBBranchName.Text)
        DtrSalarySlip.Title = "Salary slip for the month of " & Trim(CmbMonth.Text) & ", " & Trim(CmbYear.Text)
        DtrSalarySlip.Show
    Else
        Set Recset = Con.Execute("select * from salarysheet where staffid='" & Trim(CmbEmployee.Text) & "' and salaryyear='" & Trim(CmbYear.Text) & "' and branch='" & Trim(CmBBranchName.Text) & "'")
        If Recset.EOF = True Then
            MsgBox "No record found!", vbInformation
            Exit Sub
        End If
        Set DataEnvironment1 = Nothing
        DataEnvironment1.DetailedSalary_Grouping Trim(CmbEmployee.Text), Trim(CmbYear.Text), Trim(CmBBranchName.Text)
        DTRDetailedSalarySlip.Show
        If Check1.Value = vbChecked Then
            DTRDetailedSalarySlip.Sections("ReportFooter").Controls("LblPreparedBy").Visible = False
            DTRDetailedSalarySlip.Sections("ReportFooter").Controls("LblApprovedBy").Visible = False
            DTRDetailedSalarySlip.Sections("ReportFooter").Controls("LblCheckedBy").Visible = False
        End If
    End If
    Unload Me
End Sub
