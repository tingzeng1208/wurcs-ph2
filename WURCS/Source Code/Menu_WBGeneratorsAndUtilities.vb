Public Class frm_WBGeneratorsAndUtilitiesMenu

    Private Sub WBGeneratorsAndUtilitiesMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim toolTip1 As New ToolTip()
        Dim toolTip2 As New ToolTip()
        Dim toolTip3 As New ToolTip()
        Dim toolTip4 As New ToolTip()
        Dim toolTip5 As New ToolTip()
        Dim toolTip6 As New ToolTip()

        toolTip1.ShowAlways = True
        toolTip2.ShowAlways = True
        toolTip3.ShowAlways = True
        toolTip4.ShowAlways = True
        toolTip5.ShowAlways = True
        toolTip6.ShowAlways = True

        'Center the menu on the user's screen
        Me.CenterToScreen()

        toolTip1.SetToolTip(btn_Participatory_WB_Generator, "Needs QA")
        toolTip2.SetToolTip(btn_State_WB_Generator, "Tested/OK")
        toolTip3.SetToolTip(btn_Annual_WB_Generator, "Tested/OK")
        toolTip4.SetToolTip(btn_By_STCC_WB_Generator, "Tested/OK")
        toolTip5.SetToolTip(btn_Reporting_RR_Generator, "Tested/OK")
        toolTip6.SetToolTip(btn_Trim_SQL_Field_Values, "Needs QA")

    End Sub

    Private Sub btn_Return_To_MainMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_MainMenu.Click
        ' Open the Main Menu Form
        Dim frmNew As New frm_MainMenu()
        frmNew.Show()
        ' Close the Waybills and Utilities Menu
        Me.Close()
    End Sub

    Private Sub btn_Annual_WB_Generator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Annual_WB_Generator.Click
        ' Open the Form
        Dim frmNew As New Annual_WB_Generator()
        frmNew.Show()
        ' Close the Waybills and Utilities Menu
        Me.Close()
    End Sub

    Private Sub btn_State_WB_Generator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_State_WB_Generator.Click
        ' Open the Form
        Dim frmNew As New State_WB_Generator()
        frmNew.Show()
        ' Close the Waybills and Utilities Menu
        Me.Close()
    End Sub

    Private Sub btn_Reporting_RR_Generator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Reporting_RR_Generator.Click
        ' Open the Form
        Dim frmNew As New RR_WB_Generator()
        frmNew.Show()
        ' Close the Waybills and Utilities Menu
        Me.Close()
    End Sub

    Private Sub btn_Participatory_WB_Generator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Participatory_WB_Generator.Click
        ' Open the Form
        Dim frmNew As New Participatory_RR_WB_Generator
        frmNew.Show()
        ' Close the Waybills and Utilities Menu
        Me.Close()
    End Sub

    Private Sub btn_By_STCC_WB_Generator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_By_STCC_WB_Generator.Click
        ' Open the Form
        Dim frmNew As New WB_By_STCC_Generator
        frmNew.Show()
        ' Close the Waybills and Utilities Menu
        Me.Close()
    End Sub

    Private Sub btn_PUWS_Generator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_PUWS_Generator.Click
        ' Open the Form
        Dim frmNew As New PUWS_Generator
        frmNew.Show()
        ' Close the Waybills and Utilities Menu
        Me.Close()
    End Sub

    Private Sub btn_Trim_SQL_Field_Values_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Trim_SQL_Field_Values.Click
        ' Open the Form
        Dim frmNew As New Trim_SQL_Field_Values
        frmNew.Show()
        ' Close the Waybills and Utilities Menu
        Me.Close()
    End Sub

    Private Sub btn_QuarterlyMaskingUtil_Click(sender As System.Object, e As System.EventArgs) Handles btn_QuarterlyMaskingUtil.Click
        ' Open the Form
        Dim frmNew As New QuarterlyMaskingUtil
        frmNew.Show()
        ' Close the Waybills and Utilities Menu
        Me.Close()
    End Sub
End Class