Public Class frm_MainMenu

    Private Sub frm_MainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Purpose:       Executes upon load of the Main Menu
        'Assumptions:   
        'Affects:       
        'Inputs:        Stores entries from Settings
        'Returns:       N/A
        'Author:        Michael Sanders

        'Set the form so it centers on the user's screen
        CenterToScreen()
        Show()

        txtStatus.Text = "Checking Server Status - Please Wait"
        Refresh()

        Gbl_Waybill_Database_Name = My.Settings.Waybills_DB
        Gbl_Controls_Database_Name = My.Settings.Controls_DB

        ' Load the tablenames that are static
        Gbl_EIA_STCCs_TableName = "EIA_STCCS"

        If gbl_Server_Name = "" Then

            ' Test for Production Server
            If Verify_SQL_Running(My.Settings.ProductionServer) = True Then
                rdo_Production.ForeColor = Color.Green
                gbl_Server_Name = My.Settings.ProductionServer
            Else
                rdo_Production.ForeColor = Color.Red
                rdo_Production.Enabled = False
            End If

            ' Test for Test Server
            If Verify_SQL_Running(My.Settings.TestServer) = True Then
                'rdo_TestServer.Text = "Test Server: " & My.Settings.TestServer.ToString
                rdo_TestServer.ForeColor = Color.Green
                gbl_Server_Name = My.Settings.TestServer
            Else
                rdo_TestServer.ForeColor = Color.Red
                rdo_TestServer.Enabled = False
            End If

            ' Test for Local Server
            If Verify_SQL_Running(My.Settings.VirtualServer) = True Then
                rdo_LocalServer.Text = "Virtual Server: " & My.Settings.VirtualServer.ToString
                rdo_LocalServer.ForeColor = Color.Green
                gbl_Server_Name = My.Settings.VirtualServer.ToString
            Else
                rdo_LocalServer.ForeColor = Color.Red
                rdo_LocalServer.Enabled = False
            End If

            txtStatus.Text = "Please select a server..."
            Refresh()

        Else

            If gbl_Server_Name = My.Settings.ProductionServer.ToString Then
                rdo_Production.Checked = True
            ElseIf gbl_Server_Name = My.Settings.TestServer.ToString Then
                rdo_TestServer.Checked = True
            ElseIf gbl_Server_Name = My.Settings.VirtualServer.ToString Then
                rdo_LocalServer.Checked = True
                rdo_LocalServer.Text = "Virtual Server (" & gbl_Server_Name & ")"
            End If

            txtStatus.Text = "Currently using " & gbl_Server_Name
        End If

    End Sub

    Private Sub btn_Exit_Application_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Exit_Application.Click

        Application.Exit()

    End Sub

    Private Sub btn_URCS_Waybill_Data_Prep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_URCS_Waybill_Data_Prep.Click

        'Purpose:       Executes upon click
        'Assumptions:   
        'Affects:       If server radio button is not selected, display error message, else load subform.
        'Inputs:        
        'Returns:       N/A
        'Author:        Michael Sanders

        If gbl_Server_Name <> "" Then

            ' Open the subform
            Dim frmNew As New frm_URCS_Waybill_Data_Prep()
            frmNew.Show()
            ' Hide the Main Menu
            Hide()

        Else

            MsgBox("You must select a server to proceed.", vbOKOnly, "Error")

        End If

    End Sub

    Private Sub btn_URCS_Post_Processing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_URCS_Post_Processing.Click

        'Purpose:       Executes upon click
        'Assumptions:   
        'Affects:       If server radio button is not selected, display error message, else load subform.
        'Inputs:        
        'Returns:       N/A
        'Author:        Michael Sanders

        If gbl_Server_Name <> "" Then


            ' Open the subform
            Dim frmNew As New frm_Post_Processing_Menu()
            frmNew.Show()
            ' Hide the Main Menu
            Hide()

        Else

            MsgBox("You must select a server to proceed.", vbOKOnly, "Error")

        End If

    End Sub

    Private Sub btn_URCS_Phase3_Costing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_URCS_Phase3_Costing.Click

        'Purpose:       Executes upon click
        'Assumptions:   
        'Affects:       If server radio button is not selected, display error message, else load subform.
        'Inputs:        
        'Returns:       N/A
        'Author:        Michael Sanders

        If gbl_Server_Name <> "" Then
            ' Open the subform
            Dim frmNew As New frmPhase3Main
            frmNew.Show()
            ' Hide the Main Menu
            Hide()

        Else

            MsgBox("You must select a server to proceed.", vbOKOnly, "Error")

        End If

    End Sub

    Private Sub btn_URCS_Phase2_Click(sender As System.Object, e As System.EventArgs) Handles btn_URCS_Phase2.Click

        'Purpose:       Executes upon click
        'Assumptions:   
        'Affects:       If server radio button is not selected, display error message, else load subform.
        'Inputs:        
        'Returns:       N/A
        'Author:        Michael Sanders

        If gbl_Server_Name <> "" Then

            ' Open the subform
            Dim frmNew As New frmPhase2Main
            frmNew.Show()
            ' Hide the Main Menu
            Hide()

        Else

            MsgBox("You must select a server to proceed.", vbOKOnly, "Error")

        End If

    End Sub

    Private Sub btn_Preprocessing_Checks_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Preprocessing_Checks.Click

        'Purpose:       Executes upon click
        'Assumptions:   
        'Affects:       If server radio button is not selected, display error message, else load subform.
        'Inputs:        
        'Returns:       N/A
        'Author:        Michael Sanders

        If gbl_Server_Name <> "" Then

            ' Open the subform
            Dim frmNew As New frm_PreProcessing_Checks
            frmNew.Show()
            ' Hide the Main Menu
            Hide()

        Else

            MsgBox("You must select a server to proceed.", vbOKOnly, "Error")

        End If

    End Sub

    Private Sub btn_Legacy_Costing_Click(sender As Object, e As EventArgs) Handles btn_Legacy_Costing.Click

        'Purpose:       Executes upon click
        'Assumptions:   
        'Affects:       If server radio button is not selected, display error message, else load subform.
        'Inputs:        
        'Returns:       N/A
        'Author:        Michael Sanders

        If gbl_Server_Name <> "" Then

            ' Open the subform
            Dim frmNew As New Legacy_Costing
            frmNew.Show()
            ' Hide the Main Menu
            Hide()

        Else

            MsgBox("You must select a server to proceed.", vbOKOnly, "Error")

        End If

    End Sub

    Private Sub select_Server(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdo_LocalServer.CheckedChanged, rdo_Production.CheckedChanged, rdo_TestServer.CheckedChanged

        'Purpose:       Executes upon click on rdo_LocalServer, rdo_ProductionServer, rdo_TestServer
        'Assumptions:   
        'Affects:       Stores selected server value to memvar gbl_server_name
        'Inputs:        
        'Returns:       N/A
        'Author:        Laura Schneider

        If rdo_LocalServer.Checked Then
            gbl_Server_Name = My.Settings.VirtualServer.ToString
        End If
        If rdo_Production.Checked Then
            gbl_Server_Name = My.Settings.ProductionServer.ToString
        End If
        If rdo_TestServer.Checked Then
            gbl_Server_Name = My.Settings.TestServer.ToString
        End If
        txtStatus.Text = gbl_Server_Name & " will be used."
    End Sub

    Private Sub btn_WB_Sample_Utilities_Menu_Click(sender As Object, e As EventArgs) Handles btn_WB_Sample_Utilities_Menu.Click
        'Purpose:       Executes upon click
        'Assumptions:   
        'Affects:       If server radio button is not selected, display error message, else load subform.
        'Inputs:        
        'Returns:       N/A
        'Author:        Michael Sanders

        If gbl_Server_Name <> "" Then

            ' Open the subform
            Dim frmNew As New frm_WBGeneratorsAndUtilitiesMenu
            frmNew.Show()
            ' Hide the Main Menu
            Hide()

        Else

            MsgBox("You must select a server to proceed.", vbOKOnly, "Error")

        End If
    End Sub
End Class
