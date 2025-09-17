Imports System.Data.SqlClient
Public Class Participatory_RR_WB_Generator

    Private Sub Participatory_RR_WB_Generator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim rst As ADODB.Recordset
        Dim mStrSQL As String

        'Set the form so it centers on the user's screen
        CenterToScreen()

        'Set the default values of the checkboxes
        chk_Unmasked.Checked = False
        chk_Litigation.Checked = False
        chk_Litigation.Visible = False
        chk_Unmask_STCC.Checked = False
        chk_Unmask_STCC.Visible = False
        chk_AsReported.Checked = False

        ' Load the Year combobox from the SQL database

        ' Open the SQL connection to the URCS_Controls database
        OpenADOConnection(Gbl_Controls_Database_Name)
        Gbl_WB_Years_TableName = Get_Table_Name_From_SQL("1", "WB_YEARS")

        rst = SetRST()

        mStrSQL = "SELECT wb_year FROM " & Global_Variables.Gbl_WB_Years_TableName & " ORDER BY wb_year DESC"

        rst.Open(mStrSQL, gbl_ADOConnection)
        rst.MoveFirst()

        ' Load the Year values into the combobox
        Do While Not rst.EOF
            cmb_URCS_Year.Items.Add(rst.Fields("wb_year").Value)
            rst.MoveNext()
        Loop

        ' Clean up
        rst.Close()
        rst = Nothing

        'This will populate the railroad combobox

        ' Open the SQL connection using the global variable holding the connection string
        OpenADOConnection(Gbl_Controls_Database_Name)
        Gbl_URCS_WAYRRR_TableName = Get_Table_Name_From_SQL("1", "WAYRRR")
        rst = SetRST()

        mStrSQL = "SELECT RR_NAME FROM " & Gbl_URCS_WAYRRR_TableName & " ORDER BY RR_NAME ASC"

        rst.Open(mStrSQL, gbl_ADOConnection)
        rst.MoveFirst()

        Do While Not rst.EOF
            cmb_Railroad.Items.Add(Trim(rst.Fields("rr_name").Value))
            rst.MoveNext()
        Loop

        ' Clean up
        rst.Close()
        rst = Nothing
        Refresh()

    End Sub

    Private Sub btn_Return_To_WBGeneratorsMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_WBGeneratorsMenu.Click
        ' Open the WBGenerators And Utilities Menu Form
        Dim frmNew As New frm_WBGeneratorsAndUtilitiesMenu()
        frmNew.Show()
        ' Close this Menu
        Me.Close()
    End Sub

    Private Sub cmb_Railroad_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_Railroad.SelectedIndexChanged

        Dim mLoc As Integer
        Dim mStr As String

        mLoc = 0
        mStr = ""

        If txt_Output_FilePath.Text <> "" Then
            mLoc = InStr(txt_Output_FilePath.Text, "\WB")
            mStr = Mid(txt_Output_FilePath.Text, 1, mLoc + 7)
            mStr = mStr & cmb_Railroad.Text & "_"
            If chk_Unmask_STCC.Checked Then
                mLoc = InStr(txt_Output_FilePath.Text, "Unmasked")
            Else
                mLoc = InStr(txt_Output_FilePath.Text, "Masked")
            End If
            mStr = mStr & Mid(txt_Output_FilePath.Text, mLoc)
            txt_Output_FilePath.Text = Replace(mStr, " ", "_")
        End If

    End Sub

    Private Sub btn_Output_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New FolderBrowserDialog
        Dim Year_Selected As Boolean, RR_Selected As Boolean

        Year_Selected = False
        RR_Selected = False

        If cmb_URCS_Year.Text <> "" Then
            Year_Selected = True
        End If

        If cmb_Railroad.Text <> "" Then
            RR_Selected = True
        End If

        If Year_Selected = False Then
            MsgBox("You must set/select the Waybill Year before selecting an output file.", vbOKOnly, "Error!")
            GoTo SkipIt
        End If

        If RR_Selected = False Then
            MsgBox("You must set/select the Railroad before selecting an output file.", vbOKOnly, "Error!")
            GoTo SkipIt
        End If

        fd.Description = "Select the location in which you want the output report placed."

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Output_FilePath.Text = fd.SelectedPath.ToString & "\WB" & Me.cmb_URCS_Year.Text & "_"
            ' Add the description of the record to the filename
            txt_Output_FilePath.Text = txt_Output_FilePath.Text & cmb_Railroad.Text
            If chk_Unmasked.Checked = True Then
                txt_Output_FilePath.Text = txt_Output_FilePath.Text & "_Unmasked"
            Else
                txt_Output_FilePath.Text = txt_Output_FilePath.Text & "_Masked"
            End If
        End If

        ' add the extension
        txt_Output_FilePath.Text = txt_Output_FilePath.Text & ".txt"
        ' Clean up the spaces in the file name
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, " ", "_")
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, ",", "")
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "&", "And")

SkipIt:
    End Sub

    Private Sub chk_Unmasked_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Unmasked.CheckedChanged
        If chk_Unmasked.Checked = True Then
            chk_Litigation.Visible = True
            chk_Litigation.Checked = False
            chk_Unmask_STCC.Visible = True
            chk_Unmask_STCC.Checked = False
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Masked", "Unmasked")
        Else
            chk_Litigation.Checked = False
            chk_Litigation.Visible = False
            chk_Unmask_STCC.Checked = False
            chk_Unmask_STCC.Visible = False
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Unmasked", "Masked")
        End If
    End Sub


    Private Sub cmb_URCS_Year_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_URCS_Year.SelectedIndexChanged
        Dim mLoc As Integer
        Dim mStr As String

        mLoc = 0
        mStr = ""

        If txt_Output_FilePath.Text <> "" Then
            mLoc = InStr(txt_Output_FilePath.Text, "\WB")
            mStr = Mid(txt_Output_FilePath.Text, 1, mLoc + 2)
            mStr = mStr & cmb_URCS_Year.Text
            mStr = mStr & Mid(txt_Output_FilePath.Text, mLoc + 7)
            txt_Output_FilePath.Text = mStr
        End If

        If cmb_URCS_Year.Text <> "" And txt_Output_FilePath.Text <> "" And cmb_Railroad.Text <> "" Then
            btn_Execute.Enabled = True
        End If

    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click
        Dim mstrRR As String
        Dim mRailroad As Integer
        Dim mStrSQL As String
        Dim bolwrite As Boolean

        Dim mDataTable As DataTable

        OpenSQLConnection(Gbl_Controls_Database_Name)

        'Find out what the AARID number is for the railroad name selected on the form
        mDataTable = New DataTable

        mStrSQL = "SELECT AARID FROM " & Gbl_URCS_WAYRRR_TableName & " WHERE RR_NAME = '" & cmb_Railroad.Text & "'"

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        mstrRR = mDataTable.Rows(0)("aarid").ToString

        'Get the database_name and table_name for the URCS_Years value from the database
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(cmb_URCS_Year.Text, "Masked")

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        txt_StatusBox.Text = "Searching for records..."
        Refresh()

        If mRailroad = 105 Or mRailroad = 482 Then
            mStrSQL = "SELECT * FROM " & Global_Variables.Gbl_Masked_TableName & " " &
                "WHERE (report_rr = 105 or report_rr = 482 or " &
                "orr = 105 or orr = 482 or " &
                "jrr1 = 105 or jrr1 = 482 or " &
                "jrr2 = 105 or jrr2 = 482 or " &
                "jrr3 = 105 or jrr3 = 482 or " &
                "jrr4 = 105 or jrr4 = 482 or " &
                "jrr5 = 105 or jrr5 = 482 or " &
                "jrr6 = 105 or jrr6 = 482 or " &
                "trr = 105 or trr = 482)"
        Else
            mStrSQL = "SELECT * FROM " & Global_Variables.Gbl_Masked_TableName & " " &
                "WHERE report_rr = " & mstrRR & " or " &
                "orr = " & mstrRR & " or " &
                "jrr1 = " & mstrRR & " or " &
                "jrr2 = " & mstrRR & " or " &
                "jrr3 = " & mstrRR & " or " &
                "jrr4 = " & mstrRR & " or " &
                "jrr5 = " & mstrRR & " or " &
                "jrr6 = " & mstrRR & " or " &
                "trr = " & mstrRR
        End If

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        bolwrite = False
        If mDataTable.Rows.Count = 0 Then
            MsgBox("No records found for " & cmb_Railroad.Text, vbOKOnly, "Error!")
            Exit Sub
        End If

        bolwrite = MsgBox("The number of records returned is " & mDataTable.Rows.Count.ToString &
                ". Are you sure you want to write the data?", vbYesNo)

        If bolwrite = True Then
            If chk_Unmasked.Checked Then
                txt_StatusBox.Text = "Writing Unmasked Data..."
            Else
                txt_StatusBox.Text = "Writing Masked Data..."
            End If
            Refresh()

            Annual_Sample_By_RR_Writer(cmb_URCS_Year.Text,
                mstrRR.ToString,
                chk_Unmasked.Checked,
                chk_Unmask_STCC.Checked,
                chk_Litigation.Checked,
                Gbl_Waybill_Database_Name,
                Gbl_Masked_TableName,
                "WB" & cmb_URCS_Year.Text & "_Unmasked_Rev",
                txt_Output_FilePath.Text,
                900,
                True,
                cmb_Railroad.Text,
                chk_AsReported.Checked)

            txt_StatusBox.Text = "Done!"
            Refresh()

        Else
            txt_StatusBox.Text = "Aborted"
            Refresh()
        End If

    End Sub

    Private Sub txt_Output_FilePath_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_Output_FilePath.TextChanged

        If cmb_URCS_Year.Text <> "" And txt_Output_FilePath.Text <> "" And cmb_Railroad.Text <> "" Then
            'txt_StatusBox.Text = Count_Waybills_By_Railroad(cmb_URCS_Year.Text, cmb_Railroad.Text)
            btn_Execute.Enabled = True
        End If

    End Sub
End Class