
Public Class frm_URCS_Waybill_Data_Prep

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button return to main menu click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Return_To_MainMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_MainMenu.Click

        ' Close the Data Prep Menu
        Close()

        ' Shows the Main Menu which was hidden
        frm_MainMenu.Show()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Form urcs waybill data prep load. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub frm_URCS_Waybill_Data_Prep_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Center the form on the user's screen
        CenterToScreen()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button tare weight load click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Tare_Weight_Load_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Tare_Weight_Load.Click
        ' Open the Tare Weight Loader Form
        Dim frmNew As New Tare_Weight_Loader()
        frmNew.Show()
        ' Close the Data Prep Menu
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button initialize urcs year click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Initialize_URCS_Year_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Initialize_URCS_Year.Click
        ' Open the Initialize URCS Year Form
        Dim frmNew As New Initialize_URCS_Year()
        frmNew.Show()
        ' Close the Data Prep Menu
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button r 1 data load click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_R1_Data_Load_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_R1_Data_Load.Click
        Dim frmNew As New R1_Data_Load()
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button create struct 54 data load click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_CS54_Data_Load_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_CS54_Data_Load.Click
        Dim frmNew As New CS54_Data_Loader()
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button aar loader click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_AAR_Loader_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_AAR_Loader.Click
        Dim frmNew As New AAR_Index_Loader()
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button qcs data load click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_QCS_Data_Load_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_QCS_Data_Load.Click
        Dim frmNew As New QCS_Loader()
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button loss and damage load click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Loss_And_Damage_Load_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Loss_And_Damage_Load.Click
        Dim frmNew As New Loss_Damage_Loader()
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button cost of capital load click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Cost_of_Capital_Load_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Cost_of_Capital_Load.Click
        Dim frmNew As New Cost_of_Capital_Loader()
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Event handler. Called by btn_UMF_Data_Load for click events. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_UMF_Data_Load_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frmNew As New UMF_Load_Legacy()
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button waybill data loader click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Waybill_Data_Loader_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Waybill_Data_Loader.Click
        Dim frmNew As New Waybill_Data_Loader()
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button unmasked table updater click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Unmasked_Table_Updater_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Unmasked_Table_Updater.Click
        Dim frmNew As New Unmasked_Table_Update
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button way rrr verification check click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_WayRRR_Verification_Check_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_WayRRR_Verification_Check.Click
        Dim frmNew As New WayRRR_Verification
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button segments table builder click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Segments_Table_Builder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Segments_Table_Builder.Click
        Dim frmNew As New Segments_Table_Builder
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button border crossing segments click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Border_Crossing_Segments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Border_Crossing_Segments.Click
        Dim frmNew As New Border_Crossing_Segments
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button tofc load factors click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_TOFC_Load_Factors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_TOFC_Load_Factors.Click
        Dim frmNew As New TOFC_Load_Factors
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button puws loader click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_PUWS_Loader_Click(sender As System.Object, e As System.EventArgs) Handles btn_PUWS_Loader.Click
        Dim frmNew As New PUWS_Table_Update
        frmNew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Event handler. Called by btn_Convert_To_XML for click events. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Convert_To_XML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frmnew As New Convert_To_XML_Legacy
        frmnew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button stcc code table loader click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_STCC_Code_Table_Loader_Click(sender As System.Object, e As System.EventArgs) Handles btn_STCC_Code_Table_Loader.Click
        Dim frmnew As New STCC_Code_Table_Loader
        frmnew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button csm loader click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_CSM_Loader_Click(sender As Object, e As EventArgs) Handles btn_CSM_Loader.Click
        Dim frmnew As New CSM_Data_Loader
        frmnew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button marks loader click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Marks_Loader_Click(sender As Object, e As EventArgs) Handles btn_Marks_Loader.Click
        Dim frmnew As New Marks_Data_Loader
        frmnew.Show()
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button interim data loader click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/15/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Interim_Data_Loader_Click(sender As Object, e As EventArgs) Handles btn_Interim_Data_Loader.Click
        Dim frmnew As New Interim_Data_Loader
        frmnew.Show()
        Close()
    End Sub
End Class