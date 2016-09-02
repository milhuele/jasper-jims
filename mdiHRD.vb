Imports System.Windows.Forms
Imports System.IO

Public Class mdiHRD

    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticleToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileVerticalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileHorizontalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub mdiHRD_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F1
                NaviBar1.ActiveBand = NaviBand1
            Case Keys.F2
                NaviBar1.ActiveBand = NaviBand2
            Case Keys.F4
                NaviBar1.ActiveBand = NaviBand4
            Case Keys.F9
                If navUserMgt.Enabled = True Then
                    navUserMgt.PerformClick()
                Else
                    MsgBox("Sorry, you have insufficient privilege to perform this operation.", MsgBoxStyle.Exclamation, "Manage Users")
                End If
            Case Keys.F8
                navConDef.PerformClick()
            Case Keys.F7
                navChangeLogin.PerformClick()
            Case Keys.F12
                navModAccount.PerformClick()
            Case Keys.Escape
                navShutdown.PerformClick()
        End Select
    End Sub

    Private Sub MDIParent1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim counter As Integer
        Dim connectionString As String
        Dim oneKey As String
        Dim oneValue As String

        connectionString = My.Settings.MySQLDBConnection & ""
        For counter = 1 To CountSubStr(connectionString, ";") + 1
            oneKey = GetSubStr(connectionString, ";", counter)
            oneValue = Trim(GetSubStr(oneKey, "=", 2))
            oneKey = Replace(UCase(Trim(GetSubStr(oneKey, "=", 1))), " ", "")

            Select Case oneKey
                Case "SERVER"
                    tstHostName.Text = "Server/Host: " & oneValue
                Case "DATABASE"
                    tstDBName.Text = "Database: " & oneValue
            End Select
        Next counter
        LoadUserAccessPriv()
        LoadAppPathSettings()
        tstUserLogName.Text = "Currently Logged User: " & UserLogFullName
        tstLogTime.Text = "Log Date/Time: " & Date.Now.ToString("MMMM dd, yyyy hh:mm:ss tt")
    End Sub

    Private Sub LoadUserAccessPriv()
        Select Case UserDept
            Case "ACT"
                NaviBar1.ActiveBand = NaviBand2
            Case "HRD"
                NaviBar1.ActiveBand = NaviBand1
            Case "WAR"
                NaviBar1.ActiveBand = NaviBand5
            Case "ITD"
                NaviBar1.ActiveBand = NaviBand4
            Case Else
                NaviBar1.ActiveBand = Nothing
        End Select
        If UserUSRMgt <> "" Then
            If UserUSRMgt.Contains("Mgt-US") Then
                navUserMgt.Enabled = True
            Else
                navUserMgt.Enabled = False
            End If
            If UserUSRMgt.Contains("Mgt-Con") Then
                navConDef.Enabled = True
            Else
                navConDef.Enabled = False
            End If
            If UserUSRMgt.Contains("Mgt-Items") Then
                navItemMnt.Enabled = True
            Else
                navItemMnt.Enabled = False
            End If
            If UserUSRMgt.Contains("Mgt-Logins") Then
                navChangeLogin.Enabled = True
            Else
                navChangeLogin.Enabled = False
            End If
        Else
            navUserMgt.Enabled = False
            navItemMnt.Enabled = True
        End If

        If UserHRDMgt <> "" Then
            If UserHRDMgt.Contains("ER-Add") Then
                navNewEmployee.Enabled = True
            Else
                navNewEmployee.Enabled = False
            End If
            If UserHRDMgt.Contains("ER-VPrint") Then
                navViewEmployee.Enabled = True
            Else
                navViewEmployee.Enabled = False
            End If

            If UserHRDMgt.Contains("VR-Add") Then
                navNewViolations.Enabled = True
            Else
                navNewViolations.Enabled = False
            End If
            If UserHRDMgt.Contains("VR-VPrint") Then
                navViewViolations.Enabled = True
            Else
                navViewViolations.Enabled = False
            End If
        Else
            navNewEmployee.Enabled = False
            navNewViolations.Enabled = False
        End If

        If UserACTMgt <> "" Then
            If UserACTMgt.Contains("PY-Add") Then
                navNewPayroll.Enabled = True
            Else
                navNewPayroll.Enabled = False
            End If
            If UserACTMgt.Contains("PY-Edit") Then
                navViewPayroll.Enabled = True
            Else
                navViewPayroll.Enabled = False
            End If
            If UserACTMgt.Contains("PY-VPrint") Then
                navPrintPayslip.Enabled = True
            Else
                navPrintPayslip.Enabled = False
            End If

            If UserACTMgt.Contains("CM-Add") Then
                navNewCA.Enabled = True
            Else
                navNewCA.Enabled = False
            End If
            If UserACTMgt.Contains("CM-Payment") Then
                navNewPayments.Enabled = True
            Else
                navNewPayments.Enabled = False
            End If
            If UserACTMgt.Contains("CM-VPrint") Then
                navViewCAPayments.Enabled = True
            Else
                navViewCAPayments.Enabled = False
            End If

            If UserACTMgt.Contains("VO-Add") Then
                navNewVoucher.Enabled = True
            Else
                navNewVoucher.Enabled = False
            End If
            If UserACTMgt.Contains("VO-VPrint") Then
                navViewVoucher.Enabled = True
            Else
                navViewVoucher.Enabled = False
            End If
        Else
            navNewPayroll.Enabled = True
            navNewCA.Enabled = False
            navNewPayments.Enabled = False
            navNewVoucher.Enabled = False
        End If

        If UserDTRMgt <> "" Then
            If UserDTRMgt.Contains("PY-Add") Then
                navNewPayroll.Enabled = True
            Else
                navNewPayroll.Enabled = False
            End If
        End If
    End Sub

    Private Sub LoadAppPathSettings()
        Dim appPath As String = Path.GetDirectoryName(Application.ExecutablePath)
        My.Settings.ReportPath = appPath & "\Reports"
        My.Settings.EmpPhotoPath = appPath & "\Images\Photos"
        My.Settings.XMLPath = "C:\MyTemp"
    End Sub

    Private Sub tsbExitProgram_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim response As MsgBoxResult

        response = MsgBox("Are you sure you want to exit the application? ", MsgBoxStyle.Information Or MsgBoxStyle.YesNo, "Exit JIMS-HRD")

        If response = MsgBoxResult.Yes Then
            For Each ChildForm As Form In Me.MdiChildren
                ChildForm.Close()
            Next
            LogOut()
            CleanUpProgram()
        End If
    End Sub

    Private Sub tsbChangeLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "ChangeLogin" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfChangeLogin As New JIMS_HRD.frmChangeLogin
        cfChangeLogin.MdiParent = Me
        cfChangeLogin.Name = "ChangeLogin"
        cfChangeLogin.Show()
        'CheckLoadedForm("ChangeLogin")
    End Sub

    Private Sub tsbSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "AppSettings" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfAppSettings As New JIMS_HRD.frmAppSettings
        cfAppSettings.MdiParent = Me
        cfAppSettings.Name = "AppSettings"
        cfAppSettings.Show()
        'CheckLoadedForm("LocateDBase")
    End Sub

    Private Sub tsbUserMgt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "UserMgt" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfUserMgt As New JIMS_HRD.frmUserMgt
        cfUserMgt.MdiParent = Me
        cfUserMgt.Name = "UserMgt"
        cfUserMgt.Show()
        'CheckLoadedForm("UserMgt")
    End Sub

    Private Sub CheckLoadedForm(ByVal FormName As String)
        Dim LoadedForm As Form
        Try
            For Each LoadedForm In Me.MdiChildren
                If LoadedForm.Name = FormName Then
                    LoadedForm.Activate()
                    LoadedForm.Focus()
                    Exit For
                Else
                    LoadedForm.Show()
                End If
            Next
        Catch ex As Exception

        End Try
        Application.DoEvents()
    End Sub

    Private Sub MyForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        ' Do cleanup stuff here or cancel the form close
        Dim response As MsgBoxResult

        response = MsgBox("Are you sure you want to exit the application? ", MsgBoxStyle.Information Or MsgBoxStyle.YesNo, "Exit JIMS-HRD")

        If response = MsgBoxResult.Yes Then
            For Each ChildForm As Form In Me.MdiChildren
                ChildForm.Close()
            Next
            LogOut()
            Application.Exit()
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        JIMS_HRD.Splash.Show()
    End Sub

    Private Sub tsiViewEmployeeEntries_Click(sender As System.Object, e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "ViewEmployee" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfViewEmployeeEntries As New JIMS_HRD.frmGenReports
        cfViewEmployeeEntries.MdiParent = Me
        cfViewEmployeeEntries.Name = "ViewEmployee"
        cfViewEmployeeEntries.Show()
    End Sub

    Private Sub tsiNewViolationEntry_Click(sender As System.Object, e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "NewViolation" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfNewViolationEntry As New JIMS_HRD.frmNewViolation
        cfNewViolationEntry.MdiParent = Me
        cfNewViolationEntry.Name = "NewViolation"
        cfNewViolationEntry.Show()
    End Sub

    Private Sub tsiViewViolationEntries_Click(sender As System.Object, e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "ViewViolation" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfViewViolationEntries As New JIMS_HRD.frmViewViolations
        cfViewViolationEntries.MdiParent = Me
        cfViewViolationEntries.Name = "ViewViolation"
        cfViewViolationEntries.Show()
    End Sub

    Private Sub tsiNewPayrollEntry_Click(sender As System.Object, e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "NewPayroll" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfNewPayrollEntry As New JIMS_HRD.frmNewPayroll
        cfNewPayrollEntry.MdiParent = Me
        cfNewPayrollEntry.Name = "NewPayroll"
        cfNewPayrollEntry.Show()
    End Sub

    Private Sub tsiViewPayrollEntries_Click(sender As System.Object, e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "ViewPayroll" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfViewPayrollEntries As New JIMS_HRD.frmViewPayroll
        cfViewPayrollEntries.MdiParent = Me
        cfViewPayrollEntries.Name = "ViewPayroll"
        cfViewPayrollEntries.Show()
    End Sub

    Private Sub tsiPayslip_Click(sender As System.Object, e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "ViewPayslip" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfViewPayslipEntries As New JIMS_HRD.frmPrintPayroll
        cfViewPayslipEntries.MdiParent = Me
        cfViewPayslipEntries.Name = "ViewPayslip"
        cfViewPayslipEntries.Show()
    End Sub

    Private Sub tsiNewCashAdvancesEntry_Click(sender As System.Object, e As System.EventArgs)
        'For Each mdiChild As Form In Me.MdiChildren
        'If mdiChild.Name = "NewCashAdvance" Then
        'mdiChild.Activate()
        'Exit Sub
        'End If
        'Next
        'Dim cfNewCashAdvance As New JIMS_HRD.frmNewCashAdvance
        'cfNewCashAdvance.MdiParent = Me
        'cfNewCashAdvance.Name = "NewCashAdvance"
        'cfNewCashAdvance.Show()
        If JIMS_HRD.frmNewCashAdvance Is Nothing Then
            JIMS_HRD.frmNewCashAdvance.Show()
        Else
            JIMS_HRD.frmNewCashAdvance.Close()
            JIMS_HRD.frmNewCashAdvance.Show()
        End If
    End Sub

    Private Sub tsiNewCAPayment_Click(sender As System.Object, e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "NewCashAdvancePayment" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfNewCashAdvancePayment As New JIMS_HRD.frmNewCAPayment
        cfNewCashAdvancePayment.MdiParent = Me
        cfNewCashAdvancePayment.Name = "NewCashAdvancePayment"
        cfNewCashAdvancePayment.Show()
    End Sub

    Private Sub tsiViewCashAdvancesEntries_Click(sender As System.Object, e As System.EventArgs)
        Dim cfViewCashAdvanceRecords As New JIMS_HRD.frmViewCARecords
        cfViewCashAdvanceRecords.MdiParent = Me
        cfViewCashAdvanceRecords.Name = "ViewCashAdvanceRecords"
        cfViewCashAdvanceRecords.Show()

        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "ViewCashAdvanceRecords" Then
                mdiChild.Activate()
            End If
        Next
    End Sub

    Private Sub tsiNewVoucherEntry_Click(sender As System.Object, e As System.EventArgs)
        'For Each mdiChild As Form In Me.MdiChildren
        'If mdiChild.Name = "NewVoucherRecords" Then
        'mdiChild.Activate()
        'Exit Sub
        'End If
        'Next
        'Dim cfNewVoucherRecords As New JIMS_HRD.frmNewVoucher
        'cfNewVoucherRecords.MdiParent = Me
        'cfNewVoucherRecords.Name = "NewVoucherRecords"
        'cfNewVoucherRecords.Show()

        If JIMS_HRD.frmNewVoucher Is Nothing Then
            JIMS_HRD.frmNewVoucher.Show()
        Else
            JIMS_HRD.frmNewVoucher.Close()
            JIMS_HRD.frmNewVoucher.Show()
        End If
    End Sub

    Private Sub tsiViewVoucherEntries_Click(sender As System.Object, e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "ViewVoucherRecords" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfViewVoucherRecords As New JIMS_HRD.frmSearchVoucherQry
        cfViewVoucherRecords.MdiParent = Me
        cfViewVoucherRecords.Name = "ViewVoucherRecords"
        cfViewVoucherRecords.Show()
    End Sub

    Private Sub tsiNewDTREntry_Click(sender As System.Object, e As System.EventArgs)
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "NewDTRRecords" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfNewDTRRecords As New JIMS_HRD.frmNewDTREntry
        cfNewDTRRecords.MdiParent = Me
        cfNewDTRRecords.Name = "NewDTRRecords"
        cfNewDTRRecords.Show()
    End Sub

    Private Sub mdiHRD_SizeChanged(sender As Object, e As System.EventArgs) Handles Me.SizeChanged
        NaviBar1.Size = New Size(176, Me.Height - 90)
    End Sub

    Private Sub navNewEmployee_Click(sender As System.Object, e As System.EventArgs) Handles navNewEmployee.Click
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmEmployeeEntry Is Nothing Then
            JIMS_HRD.frmEmployeeEntry.Show()
        Else
            JIMS_HRD.frmEmployeeEntry.Close()
            JIMS_HRD.frmEmployeeEntry.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navViewEmployee_Click(sender As System.Object, e As System.EventArgs) Handles navViewEmployee.Click
        Me.Cursor = Cursors.WaitCursor
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "ViewEmployee" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfViewEmployeeEntries As New JIMS_HRD.frmGenReports
        cfViewEmployeeEntries.MdiParent = Me
        cfViewEmployeeEntries.Name = "ViewEmployee"
        cfViewEmployeeEntries.Show()
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navNewViolations_Click(sender As System.Object, e As System.EventArgs) Handles navNewViolations.Click
        Me.Cursor = Cursors.WaitCursor
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "NewViolation" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfNewViolationEntry As New JIMS_HRD.frmNewViolation
        cfNewViolationEntry.MdiParent = Me
        cfNewViolationEntry.Name = "NewViolation"
        cfNewViolationEntry.Show()
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navViewViolations_Click(sender As System.Object, e As System.EventArgs) Handles navViewViolations.Click
        'For Each mdiChild As Form In Me.MdiChildren
        'If mdiChild.Name = "ViewViolation" Then
        'mdiChild.Activate()
        'Exit Sub
        'End If
        'Next
        'Dim cfViewViolationEntries As New JIMS_HRD.frmViewViolations
        'cfViewViolationEntries.MdiParent = Me
        'cfViewViolationEntries.Name = "ViewViolation"
        'cfViewViolationEntries.Show()
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmViewViolations Is Nothing Then
            JIMS_HRD.frmViewViolations.Show()
        Else
            JIMS_HRD.frmViewViolations.Close()
            JIMS_HRD.frmViewViolations.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navNewPayroll_Click(sender As System.Object, e As System.EventArgs) Handles navNewPayroll.Click
        OpenJIMSWindow(frmNewPayroll)
    End Sub

    Private Sub navViewPayroll_Click(sender As System.Object, e As System.EventArgs) Handles navViewPayroll.Click
        'For Each mdiChild As Form In Me.MdiChildren
        'If mdiChild.Name = "ViewPayroll" Then
        'mdiChild.Activate()
        'Exit Sub
        'End If
        'Next
        'Dim cfViewPayrollEntries As New JIMS_HRD.frmViewPayroll
        'cfViewPayrollEntries.MdiParent = Me
        'cfViewPayrollEntries.Name = "ViewPayroll"
        'cfViewPayrollEntries.Show()
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmViewPayroll Is Nothing Then
            JIMS_HRD.frmViewPayroll.Show()
        Else
            JIMS_HRD.frmViewPayroll.Close()
            JIMS_HRD.frmViewPayroll.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navViewPayslip_Click(sender As System.Object, e As System.EventArgs) Handles navPrintPayslip.Click
        Me.Cursor = Cursors.WaitCursor
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "ViewPayslip" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfViewPayslipEntries As New JIMS_HRD.frmPrintPayroll
        cfViewPayslipEntries.MdiParent = Me
        cfViewPayslipEntries.Name = "ViewPayslip"
        cfViewPayslipEntries.Show()
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navNewCA_Click(sender As System.Object, e As System.EventArgs) Handles navNewCA.Click
        'For Each mdiChild As Form In Me.MdiChildren
        'If mdiChild.Name = "NewCashAdvance" Then
        'mdiChild.Activate()
        'Exit Sub
        'End If
        'Next
        'Dim cfNewCashAdvance As New JIMS_HRD.frmNewCashAdvance
        'cfNewCashAdvance.MdiParent = Me
        'cfNewCashAdvance.Name = "NewCashAdvance"
        'cfNewCashAdvance.Show()
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmNewCashAdvance Is Nothing Then
            JIMS_HRD.frmNewCashAdvance.Show()
        Else
            JIMS_HRD.frmNewCashAdvance.Close()
            JIMS_HRD.frmNewCashAdvance.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navNewPayments_Click(sender As System.Object, e As System.EventArgs) Handles navNewPayments.Click
        'For Each mdiChild As Form In Me.MdiChildren
        'If mdiChild.Name = "NewCashAdvancePayment" Then
        'mdiChild.Activate()
        'Exit Sub
        'End If
        'Next
        'Dim cfNewCashAdvancePayment As New JIMS_HRD.frmNewCAPayment
        'cfNewCashAdvancePayment.MdiParent = Me
        'cfNewCashAdvancePayment.Name = "NewCashAdvancePayment"
        'cfNewCashAdvancePayment.Show()
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmNewCAPayment Is Nothing Then
            JIMS_HRD.frmNewCAPayment.Show()
        Else
            JIMS_HRD.frmNewCAPayment.Close()
            JIMS_HRD.frmNewCAPayment.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navViewCAPayments_Click(sender As System.Object, e As System.EventArgs) Handles navViewCAPayments.Click
        'Dim cfViewCashAdvanceRecords As New JIMS_HRD.frmViewCARecords
        'cfViewCashAdvanceRecords.MdiParent = Me
        'cfViewCashAdvanceRecords.Name = "ViewCashAdvanceRecords"
        'cfViewCashAdvanceRecords.Show()

        'For Each mdiChild As Form In Me.MdiChildren
        'If mdiChild.Name = "ViewCashAdvanceRecords" Then
        'mdiChild.Activate()
        'End If
        'Next
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmViewCARecords Is Nothing Then
            JIMS_HRD.frmViewCARecords.Show()
        Else
            JIMS_HRD.frmViewCARecords.Close()
            JIMS_HRD.frmViewCARecords.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navNewVoucher_Click(sender As System.Object, e As System.EventArgs) Handles navNewVoucher.Click
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmNewVoucherEntry Is Nothing Then
            JIMS_HRD.frmNewVoucherEntry.Show()
        Else
            JIMS_HRD.frmNewVoucherEntry.Close()
            JIMS_HRD.frmNewVoucherEntry.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navViewVoucher_Click(sender As System.Object, e As System.EventArgs) Handles navViewVoucher.Click
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmSearchVoucherQry Is Nothing Then
            JIMS_HRD.frmSearchVoucherQry.Show()
        Else
            JIMS_HRD.frmSearchVoucherQry.Close()
            JIMS_HRD.frmSearchVoucherQry.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navNewDTR_Click(sender As System.Object, e As System.EventArgs) Handles navNewDTR.Click
        If JIMS_HRD.frmNewDTREntry Is Nothing Then
            JIMS_HRD.frmNewDTREntry.Show()
        Else
            JIMS_HRD.frmNewDTREntry.Close()
            JIMS_HRD.frmNewDTREntry.Show()
        End If
    End Sub

    Private Sub navUserMgt_Click(sender As System.Object, e As System.EventArgs) Handles navUserMgt.Click
        Me.Cursor = Cursors.WaitCursor
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "UserMgt" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfUserMgt As New JIMS_HRD.frmUserMgt
        cfUserMgt.MdiParent = Me
        cfUserMgt.Name = "UserMgt"
        cfUserMgt.Show()
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navConDef_Click(sender As System.Object, e As System.EventArgs) Handles navConDef.Click
        Me.Cursor = Cursors.WaitCursor
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "AppSettings" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfAppSettings As New JIMS_HRD.frmAppSettings
        cfAppSettings.MdiParent = Me
        cfAppSettings.Name = "AppSettings"
        cfAppSettings.Show()
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navChangeLogin_Click(sender As System.Object, e As System.EventArgs) Handles navChangeLogin.Click
        Me.Cursor = Cursors.WaitCursor
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "ChangeLogin" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfChangeLogin As New JIMS_HRD.frmChangeLogin
        cfChangeLogin.MdiParent = Me
        cfChangeLogin.Name = "ChangeLogin"
        cfChangeLogin.Show()
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navShutdown_Click(sender As System.Object, e As System.EventArgs) Handles navShutdown.Click
        Dim response As MsgBoxResult

        response = MsgBox("Are you sure you want to exit the application? ", MsgBoxStyle.Information Or MsgBoxStyle.YesNo, "Exit JIMS-HRD")
        If response = MsgBoxResult.Yes Then
            For Each ChildForm As Form In Me.MdiChildren
                ChildForm.Close()
            Next
            LogOut()
            CleanUpProgram()
        End If
    End Sub

    Private Sub navEmpPER_Click(sender As System.Object, e As System.EventArgs) Handles navEmpPER.Click
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmEmployeePER Is Nothing Then
            JIMS_HRD.frmEmployeePER.Show()
        Else
            JIMS_HRD.frmEmployeePER.Close()
            JIMS_HRD.frmEmployeePER.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navInbound_Click(sender As System.Object, e As System.EventArgs) Handles navInbound.Click
        If JIMS_HRD.frmInbound Is Nothing Then
            JIMS_HRD.frmInbound.Show()
        Else
            JIMS_HRD.frmInbound.Close()
            JIMS_HRD.frmInbound.Show()
        End If
    End Sub

    Private Sub navNewPO_Click(sender As System.Object, e As System.EventArgs) Handles navNewPO.Click
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmNewPurchaseOrder Is Nothing Then
            JIMS_HRD.frmNewPurchaseOrder.Show()
        Else
            JIMS_HRD.frmNewPurchaseOrder.Close()
            JIMS_HRD.frmNewPurchaseOrder.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navViewPO_Click(sender As System.Object, e As System.EventArgs) Handles navViewPO.Click
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmViewPurchaseOrder Is Nothing Then
            JIMS_HRD.frmViewPurchaseOrder.Show()
        Else
            JIMS_HRD.frmViewPurchaseOrder.Close()
            JIMS_HRD.frmViewPurchaseOrder.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navItemMnt_Click(sender As System.Object, e As System.EventArgs) Handles navItemMnt.Click
        Me.Cursor = Cursors.WaitCursor
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "HRItemMaintenance" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfItemMaintenance As New JIMS_HRD.frmHRDSettings
        cfItemMaintenance.MdiParent = Me
        cfItemMaintenance.Name = "HRItemMaintenance"
        cfItemMaintenance.Show()
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub cmdACTSettings_Click(sender As System.Object, e As System.EventArgs) Handles cmdACTSettings.Click
        Me.Cursor = Cursors.WaitCursor
        For Each mdiChild As Form In Me.MdiChildren
            If mdiChild.Name = "ACTItemMaintenance" Then
                mdiChild.Activate()
                Exit Sub
            End If
        Next
        Dim cfACTItemMaintenance As New JIMS_HRD.frmACTSettings
        cfACTItemMaintenance.MdiParent = Me
        cfACTItemMaintenance.Name = "ACTItemMaintenance"
        cfACTItemMaintenance.Show()
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navNewInvoice_Click(sender As System.Object, e As System.EventArgs) Handles navNewInvoice.Click
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmNewInvoiceEntry Is Nothing Then
            JIMS_HRD.frmNewInvoiceEntry.Show()
        Else
            JIMS_HRD.frmNewInvoiceEntry.Close()
            JIMS_HRD.frmNewInvoiceEntry.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navViewInvoice_Click(sender As System.Object, e As System.EventArgs) Handles navViewInvoice.Click
        Me.Cursor = Cursors.WaitCursor
        If JIMS_HRD.frmViewInvoices Is Nothing Then
            JIMS_HRD.frmViewInvoices.Show()
        Else
            JIMS_HRD.frmViewInvoices.Close()
            JIMS_HRD.frmViewInvoices.Show()
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub navModAccount_Click(sender As System.Object, e As System.EventArgs) Handles navModAccount.Click
        OpenJIMSWindow(frmModifyAccount)
    End Sub

    Private Sub navNewWMSItem_Click(sender As System.Object, e As System.EventArgs) Handles navNewWMSItem.Click
        OpenJIMSWindow(frmNewWarehouseItem)
    End Sub

    Private Sub navViewWMSItem_Click(sender As System.Object, e As System.EventArgs) Handles navViewWMSItem.Click
        OpenJIMSWindow(frmViewWarehouseItem2)
    End Sub
End Class
