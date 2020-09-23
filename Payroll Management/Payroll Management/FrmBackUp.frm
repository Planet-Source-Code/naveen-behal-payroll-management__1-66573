VERSION 5.00
Begin VB.Form FrmBackUp 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Data Back Up Section"
   ClientHeight    =   4680
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   2640
      TabIndex        =   2
      Top             =   45
      Width           =   4335
      Begin VB.DirListBox Dir1 
         Height          =   3465
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   3855
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00EFEADE&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H00EFEADE&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   4830
      Left            =   0
      Picture         =   "FrmBackUp.frx":0000
      Top             =   -120
      Width           =   2490
   End
End
Attribute VB_Name = "FrmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim a As New Scripting.FileSystemObject

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Screen.MousePointer = vbHourglass
    a.CopyFile SysPath & "\sowildata.mdb", Dir1.Path & "Salary" & Day(Date) & "_" & Month(Date) & "_" & Year(Date) & ".mdb", True
    MsgBox "Backup successfully taken", vbInformation
    Screen.MousePointer = vbNormal
End Sub
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    CenterMe Me
End Sub
