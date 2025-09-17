Public Class frm_Post_Processing_Menu

    Private Sub btn_Return_To_MainMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_MainMenu.Click
        ' Open the Main Menu Form
        Dim frmNew As New frm_MainMenu()
        frmNew.Show()
        ' Close the Post Processing Menu
        Me.Close()
    End Sub

    Private Sub PostProcessingMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Center the form on the user's screen
        Me.CenterToScreen()
    End Sub

    Private Sub btn_XML_Loader_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_XML_Loader.Click
        Dim frmNew As New frm_XML_Data_Loader
        frmNew.Show()
        Me.Close()
    End Sub

    Private Sub btn_FileToFile_Compare_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FileToFile_Compare.Click
        Dim frmNew As New frm_XML_Comparison
        frmNew.Show()
        Me.Close()
    End Sub


    Private Sub btn_Productivity_Click(sender As System.Object, e As System.EventArgs) Handles btn_Productivity.Click
        Dim frmNew As New frm_Productivity
        frmNew.Show()
        Me.Close()
    End Sub

    Private Sub btn_Add_Make_Whole_Click(sender As Object, e As EventArgs) Handles btn_Add_Make_Whole.Click
        Dim frmNew As New frm_AddMakeWholeToXML
        frmNew.Show()
        Me.Close()
    End Sub
End Class