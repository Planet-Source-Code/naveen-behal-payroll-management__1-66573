VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmMedical 
   BackColor       =   &H00EFEADE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Medical Reimbursesment"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FRMMEDICAL.frx":0000
      Height          =   5895
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "mmonths"
         Caption         =   "For The Months"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "branch"
         Caption         =   "Branch Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "employee"
         Caption         =   "Employee"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "medical"
         Caption         =   "Medical"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "mfirst"
         Caption         =   "First"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "msecond"
         Caption         =   "Second"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "mthird"
         Caption         =   "Third"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Expr1007"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column01 
            WrapText        =   -1  'True
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2025.071
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   720
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00EFEADE&
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   780
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEADE&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.ComboBox CmbMonths 
         Height          =   315
         ItemData        =   "FRMMEDICAL.frx":0015
         Left            =   1920
         List            =   "FRMMEDICAL.frx":0025
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Months"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00EFEADE&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   780
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2040
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\SowilData.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\SowilData.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select mmonths, branch, employee, medical, mfirst, msecond, mthird, mfirst + msecond + mthird from medical"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmMedical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbMonths_Click()
    Screen.MousePointer = vbHourglass
    DoEvents
    DoEvents
    ShowGrid
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Unload DataEnvironment1
    If CmbMonths.ListIndex = 0 Then
        DtrMedical.Sections("ReportHeader").Controls("LblHead").Caption = "Detail of the MEDICAL, payable for the period from April to June"
        DtrMedical.Sections("PageHeader").Controls("Lblfirst").Caption = "April"
        DtrMedical.Sections("PageHeader").Controls("LblSecond").Caption = "May"
        DtrMedical.Sections("PageHeader").Controls("LblThird").Caption = "June"
    ElseIf CmbMonths.ListIndex = 1 Then
        DtrMedical.Sections("ReportHeader").Controls("LblHead").Caption = "Detail of the MEDICAL, payable for the period from July to September"
        DtrMedical.Sections("PageHeader").Controls("Lblfirst").Caption = "July"
        DtrMedical.Sections("PageHeader").Controls("LblSecond").Caption = "August"
        DtrMedical.Sections("PageHeader").Controls("LblThird").Caption = "September"
    ElseIf CmbMonths.ListIndex = 2 Then
        DtrMedical.Sections("ReportHeader").Controls("LblHead").Caption = "Detail of the MEDICAL, payable for the period from October to December"
        DtrMedical.Sections("PageHeader").Controls("Lblfirst").Caption = "October"
        DtrMedical.Sections("PageHeader").Controls("LblSecond").Caption = "November"
        DtrMedical.Sections("PageHeader").Controls("LblThird").Caption = "December"
    ElseIf CmbMonths.ListIndex = 3 Then
        DtrMedical.Sections("ReportHeader").Controls("LblHead").Caption = "Detail of the MEDICAL, payable for the period from January to March"
        DtrMedical.Sections("PageHeader").Controls("Lblfirst").Caption = "January"
        DtrMedical.Sections("PageHeader").Controls("LblSecond").Caption = "February"
        DtrMedical.Sections("PageHeader").Controls("LblThird").Caption = "March"
    Else
        MsgBox "Please select months first", vbInformation
        CmbMonths.SetFocus
        Exit Sub
    End If
    DataEnvironment1.Medical_Grouping CmbMonths.Text
    DtrMedical.Show
End Sub
Private Sub Form_Load()
    CenterMe Me
    OpenCon
End Sub
Sub ShowGrid()
Screen.MousePointer = vbHourglass
DoEvents
Dim TTT As Integer
    If CmbMonths = "" Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    Set Recset = Nothing
    If CmbMonths.ListIndex = 0 Then
        Set Recset = Con.Execute("select * from salarysheet where mno=3")
    ElseIf CmbMonths.ListIndex = 1 Then
        Set Recset = Con.Execute("select * from salarysheet where mno=7")
    ElseIf CmbMonths.ListIndex = 2 Then
        Set Recset = Con.Execute("select * from salarysheet where mno=10")
    Else
        Set Recset = Con.Execute("select * from salarysheet where mno=13")
    End If
    If CmbMonths.ListIndex = 0 Then TTT = 1 Else TTT = 2
    If Recset.EOF = True Then
        Adodc1.RecordSource = "select * from medical where 1>1"
        Adodc1.Refresh
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    Set TmpRecset = Nothing
    Set Recset = Nothing
    Set Recset = Con.Execute("select mmonths, branch, employee, medical, mfirst, msecond, mthird, mfirst+msecond+mthird from medical where mmonths='" & CmbMonths.Text & "'")
    If Recset.EOF = False Then
        Adodc1.RecordSource = Recset.Source
        Adodc1.Refresh
        DataGrid1.Refresh
        Screen.MousePointer = vbNormal
        Exit Sub
    Else
        Set Recset = Nothing
        Set Recset = Con.Execute("select branch,staffid,medicalreim from staffdetails where medicalreim>0")
        If Recset.EOF = True Then
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        While Recset.EOF = False
            Con.Execute "insert into medical values('" & Trim(Recset.Fields(0).Value) & "','" & Trim(Recset.Fields(1).Value) & "'," & Val(Recset.Fields(2).Value) & ",0,0,0,'" & CmbMonths.Text & "')"
            Set TmpRecset = Con.Execute("select mno, mm * " & Recset.Fields(2).Value & " from salarysheet where branch='" & Recset.Fields(0).Value & "' and staffid='" & Recset.Fields(1).Value & "' and mno=" & TTT + 3 * Val(CmbMonths.ListIndex))
            While TmpRecset.EOF = False
                Con.Execute "update medical set mfirst=" & Round(TmpRecset.Fields(1).Value, 0) & " where branch='" & Recset.Fields(0).Value & "' and employee='" & Recset.Fields(1).Value & "' and mmonths='" & CmbMonths.Text & "'"
                TmpRecset.MoveNext
            Wend
            Set TmpRecset = Nothing
            Set TmpRecset = Con.Execute("select mno, mm * " & Recset.Fields(2).Value & " from salarysheet where branch='" & Recset.Fields(0).Value & "' and staffid='" & Recset.Fields(1).Value & "' and mno=" & 1 + TTT + 3 * Val(CmbMonths.ListIndex))
            While TmpRecset.EOF = False
                Con.Execute "update medical set msecond=" & Round(TmpRecset.Fields(1).Value, 0) & " where branch='" & Recset.Fields(0).Value & "' and employee='" & Recset.Fields(1).Value & "' and mmonths='" & CmbMonths.Text & "'"
                TmpRecset.MoveNext
            Wend
            Set TmpRecset = Nothing
            Set TmpRecset = Con.Execute("select mno, mm * " & Recset.Fields(2).Value & " from salarysheet where branch='" & Recset.Fields(0).Value & "' and staffid='" & Recset.Fields(1).Value & "' and mno=" & 2 + TTT + 3 * Val(CmbMonths.ListIndex))
            While TmpRecset.EOF = False
                Con.Execute "update medical set mthird=" & Round(TmpRecset.Fields(1).Value, 0) & " where branch='" & Recset.Fields(0).Value & "' and employee='" & Recset.Fields(1).Value & "' and mmonths='" & CmbMonths.Text & "'"
                TmpRecset.MoveNext
            Wend
            Recset.MoveNext
        Wend
    End If
    Set TmpRecset = Nothing
    Set Recset = Nothing
    Set Recset = Con.Execute("select mmonths, branch, employee, medical, mfirst, msecond, mthird, mfirst+msecond+mthird from medical where mmonths='" & CmbMonths.Text & "'")
    If Recset.EOF = False Then
        Adodc1.RecordSource = Recset.Source
        Adodc1.Refresh
        DataGrid1.Refresh
    End If
    Screen.MousePointer = vbNormal
End Sub
