Public Class frm_PreProcessing_Checks

    Private Sub btn_Return_To_MainMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_MainMenu.Click
        ' Open the Main Menu Form
        Dim frmNew As New frm_MainMenu()
        frmNew.Show()
        ' Close the Post Processing Menu
        Me.Close()
    End Sub

    Private Sub frm_PreProcessing_Checks_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

    End Sub

    Private Sub btn_R1_Balance_Check_Click(sender As System.Object, e As System.EventArgs) Handles btn_R1_Balance_Check.Click

        ' Open the R1 Balance Check Form
        Dim frmNew As New frm_R_1_Balance
        frmNew.Show()
        ' Close this Menu
        Me.Close()

    End Sub

    
    Private Sub btn_Trans_Data_Comparison_Click(sender As System.Object, e As System.EventArgs) Handles btn_Trans_Data_Comparison.Click

        ' Open the Trans Data Comparison Form
        Dim frmNew As New frm_Trans_Comparison
        frmNew.Show()
        ' Close this Menu
        Me.Close()

    End Sub
End Class