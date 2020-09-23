VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   2010
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox CmBBranchName 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   3255
      End
      Begin VB.ComboBox CmbYear 
         Height          =   315
         ItemData        =   "Dialog.frx":0000
         Left            =   840
         List            =   "Dialog.frx":0028
         TabIndex        =   1
         Top             =   765
         Width           =   3255
      End
      Begin VB.ComboBox CmbMonth 
         Height          =   315
         ItemData        =   "Dialog.frx":0062
         Left            =   840
         List            =   "Dialog.frx":008A
         TabIndex        =   2
         Top             =   1245
         Width           =   3255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
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
    OpenCon
    Set Recset = Con.Execute("select distinct(branch) from salarysheet")
    While Recset.EOF = False
        CmBBranchName.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
End Sub

Private Sub OKButton_Click()
    Set Recset = Nothing
    Set Recset = Con.Execute("select * from salarysheet where branch='" & Trim(CmBBranchName.Text) & "' and salaryyear='" & Trim(CmbYear.Text) & "' and salarymonth='" & Trim(CmbMonth.Text) & "'")
    If Recset.EOF = True Then
        MsgBox "No record found for the criteria", vbInformation
        Exit Sub
    End If
    Set DataEnvironment1 = Nothing
    DataEnvironment1.SalarySheet Trim(CmBBranchName.Text), Trim(CmbYear.Text), Trim(CmbMonth.Text)
    If Me.Caption = "Salary Sheet With Deductions" Then
        DtrSalarySheetD.Sections("section2").Controls("Lblbranch").Caption = CmBBranchName.Text
        DtrSalarySheetD.Sections("section2").Controls("LblMonth").Caption = "Salary Sheet for the month of " & CmbMonth.Text & " " & Mid(CmbYear.Text, 1, 4)
        DtrSalarySheetD.Show
    Else
        DtrSalarySheet.Sections("section2").Controls("Lblbranch").Caption = CmBBranchName.Text
        DtrSalarySheet.Sections("section2").Controls("LblMonth").Caption = "Salary Sheet for the month of " & CmbMonth.Text & " " & Mid(CmbYear.Text, 1, 4)
        DtrSalarySheet.Show
    End If
    Unload Me
End Sub
