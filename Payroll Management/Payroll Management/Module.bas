Attribute VB_Name = "Module1"
Public AddEdit As Boolean, UserName As String, SysPath As String, UType As String
Public Con As New ADODB.Connection, Recset As New ADODB.Recordset, TmpRecset As New ADODB.Recordset
Sub CenterMe(Frm As Form)
    Frm.Left = (MDIForm1.ScaleWidth - Frm.Width) / 2
    Frm.Top = (MDIForm1.ScaleHeight - Frm.Height) / 2
End Sub
Sub EnableCmdMe(Frm As Form)
    With Frm
        .CmdAdd.Enabled = True
        .CmdCancel.Enabled = False
        .CmdSave.Enabled = False
        .CmdExit.Enabled = True
        .CmdExit.Cancel = True
        .CmdAdd.Default = True
    End With
End Sub
Sub DisableCmdMe(Frm As Form)
    With Frm
        .CmdAdd.Enabled = False
        .CmdEdit.Enabled = False
        .CmdDelete.Enabled = False
        .CmdCancel.Enabled = True
        .CmdSave.Enabled = True
        .CmdExit.Enabled = False
        .CmdFind.Enabled = False
        .CmdCancel.Cancel = True
        '.CmdSave.Default = True
    End With
End Sub
Sub OpenCon()
    Set Recset = Nothing
    Set TmpRecset = Nothing
    Set Con = Nothing
    Con.ConnectionString = "provider='Microsoft.jet.oledb.4.0'; data source='SowilData.mdb'"
    Con.Open
End Sub
Sub Closecon()
    Set Recset = Nothing
    Set TmpRecset = Nothing
    Set Con = Nothing
End Sub
Sub LogOff()
    MDIForm1.mnuadmin.Visible = False
    MDIForm1.mnureports.Visible = False
    MDIForm1.mnulogin.Visible = True
    MDIForm1.mnulogoff.Visible = False
    MDIForm1.StatusBar1.Panels(2).Text = "waiting..."
    MDIForm1.StatusBar1.Panels(4).Text = "waiting..."
    MDIForm1.Toolbar.Visible = False
    Unload Dialog
    Unload Dialog1
    Unload FrmArrear
    Unload FrmBackUp
    Unload FrmChangePass
    Unload FrmDeleteUser
    Unload frmLogin
    Unload FrmMedical
    Unload FrmMWSS
    Unload FrmNewUser
    Unload FrmSalary
    Unload FrmStaff
    Unload FrmTDSRange
End Sub
Function MonthNo(MonNAme As String) As Integer
    Select Case Trim(MonNAme)
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
    MonthNo = mno
End Function
Sub main()
    SysPath = App.Path
    frmSplash.Show
    LogOff
End Sub
Function OnlyAlpha(a As Integer) As Integer
    If (a >= 65 And a <= 90) Or (a >= 97 And a <= 122) Or a = 32 Then
    End If
End Function
