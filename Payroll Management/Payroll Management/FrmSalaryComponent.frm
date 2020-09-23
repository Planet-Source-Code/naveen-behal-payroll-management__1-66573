VERSION 5.00
Begin VB.Form FrmSalaryComponante 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Salary Componante"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtComponanteName 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.ComboBox CmbRateType 
         Height          =   315
         ItemData        =   "FrmSalaryComponent.frx":0000
         Left            =   6240
         List            =   "FrmSalaryComponent.frx":000A
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox CmbAddDed 
         Height          =   315
         ItemData        =   "FrmSalaryComponent.frx":0021
         Left            =   2640
         List            =   "FrmSalaryComponent.frx":002B
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox TxtValueRate 
         Height          =   285
         Left            =   6240
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Type"
         Height          =   195
         Left            =   5160
         TabIndex        =   4
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Addition/Dedcuction"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value/Rate"
         Height          =   195
         Left            =   5160
         TabIndex        =   2
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Componante Name"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FrmSalaryComponante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
