VERSION 5.00
Begin VB.Form FrmMWSS 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Month Wise Salary Paid Summary"
   ClientHeight    =   840
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4455
      Begin VB.ComboBox CmbMonth 
         Height          =   315
         ItemData        =   "FrmMWSS.frx":0000
         Left            =   1560
         List            =   "FrmMWSS.frx":0028
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00EFEADE&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   285
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   400
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00EFEADE&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   285
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   85
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMWSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CancelButton_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Set Recset = Nothing
    Set Recset = Con.Execute("select distinct(salarymonth),mno from salarysheet order by mno")
    CmbMonth.Clear
    While Recset.EOF = False
        CmbMonth.AddItem Recset.Fields(0).Value
        Recset.MoveNext
    Wend
    CenterMe Me
End Sub
Private Sub OKButton_Click()
    If Trim(CmbMonth.Text) = "" Then
        MsgBox "Please select valid month", vbInformation
        CmbMonth.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    DataEnvironment1.MWSS Trim(CmbMonth.Text)
    DTRMWSS.Sections("ReportHeader").Controls("Label16").Caption = "Salary Paid Summary for the month of " & CmbMonth.Text
    DTRMWSS.Show
    Unload Me
    Screen.MousePointer = vbNormal
End Sub
