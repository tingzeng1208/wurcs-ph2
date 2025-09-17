Imports System.Data.SqlClient
Public Class WB_By_STCC_Generator

    Private Sub WB_By_STCC_Generator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        CenterToScreen()

        'Set the default values of the checkboxes
        chk_Unmasked.Checked = False
        chk_Unmask_STCC.Checked = False
        chk_Unmask_STCC.Visible = False

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        Gbl_Controls_Database_Name = "URCS_Controls"

        gbl_SQLConnection = New SqlConnection
        gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(Gbl_Controls_Database_Name)
        gbl_SQLConnection.Open()

        Gbl_WB_Years_TableName = Get_Table_Name_From_SQL("1", "WB_Years")

        mDataTable = New DataTable

        mDataTable = Get_Waybill_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            cmb_WB_Years.Items.Add(mDataTable.Rows(mLooper)("wb_year").ToString)
        Next

        mDataTable = Nothing

    End Sub

    Private Sub btn_Return_To_WBGeneratorsMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_WBGeneratorsMenu.Click
        ' Open the WBGenerators And Utilities Menu Form
        Dim frmNew As New frm_WBGeneratorsAndUtilitiesMenu()
        frmNew.Show()
        ' Close this Menu
        Me.Close()
    End Sub

    Private Sub chk_Unmasked_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Unmasked.CheckedChanged
        If chk_Unmasked.Checked = True Then
            chk_Unmask_STCC.Visible = True
            chk_Unmask_STCC.Checked = False
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Masked", "Unmasked")
        Else
            chk_Unmask_STCC.Checked = False
            chk_Unmask_STCC.Visible = False
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Unmasked", "Masked")
        End If
    End Sub

    Private Sub btn_Output_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New FolderBrowserDialog
        Dim Year_Selected As Boolean
        Dim RR_Selected As Boolean

        Year_Selected = False
        RR_Selected = False

        If cmb_WB_Years.Text <> "" Then
            Year_Selected = True
        End If

        If txt_STCC_Code.Text <> "" Then
            RR_Selected = True
        End If

        If Year_Selected = False Then
            MsgBox("You must set/select the Waybill Year before selecting an output file.", vbOKOnly, "Error!")
            GoTo SkipIt
        End If

        If RR_Selected = False Then
            MsgBox("You must enter a STCC Code value before selecting an output file.", vbOKOnly, "Error!")
            GoTo SkipIt
        End If

        fd.Description = "Select the location in which you want the output report placed."

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Output_FilePath.Text = fd.SelectedPath.ToString & "\WB" & Me.cmb_WB_Years.Text & "_"
            ' Add the description of the record to the filename
            txt_Output_FilePath.Text = txt_Output_FilePath.Text & "STCC_" & txt_STCC_Code.Text
            txt_Output_FilePath.Text = txt_Output_FilePath.Text & "_Masked"
        End If
        ' add the extension
        txt_Output_FilePath.Text = txt_Output_FilePath.Text & ".txt"
        ' Clean up the spaces in the file name
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, " ", "_")
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, ",", "")
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "&", "And")

        If chk_Unmasked.Checked = True Then
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Masked", "Unmasked")
        Else
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Unmasked", "Masked")
        End If

SkipIt:
    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click
        Dim mStrSQL As String, mStrWork As String
        Dim mfrmSTCC As String
        Dim mPromptYear As String
        Dim bolWrite As Object
        Dim mPos As Integer
        Dim mMaxRecs As Decimal


        Dim rst As New ADODB.Recordset
        Dim mDataTable As DataTable

        gbl_Database_Name = Get_Database_Name_From_SQL(cmb_WB_Years.Text, "Masked")
        Gbl_Controls_Database_Name = Get_Database_Name_From_SQL(1, "Trans")
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(cmb_WB_Years.Text, "Masked")
        Gbl_Unmasked_Rev_TableName = Get_Table_Name_From_SQL(cmb_WB_Years.Text, "UnMasked_Rev")

        OpenSQLConnection(gbl_Database_Name)
        'OpenADOConnection(Get_Database_Name_From_SQL(cmb_WB_Years.Text, "Masked"))

        ' Store the selected STCC on the form, stripping any wildcard
        mStrWork = txt_STCC_Code.Text
        mPos = InStr(mStrWork, "*")
        If mPos > 0 Then
            mfrmSTCC = Mid(mStrWork, 1, mPos - 1)
        Else
            mfrmSTCC = mStrWork
        End If

        'Initialize the return recordset and run the query to load it
        'rst = SetRST()

        mPromptYear = cmb_WB_Years.Text

        If Mid(mfrmSTCC, 1, 2) = "49" Then
            mStrSQL = "SELECT COUNT(*) AS MyCount " &
                      "FROM " & Global_Variables.Gbl_Masked_TableName & " INNER JOIN " &
                      Global_Variables.Gbl_Unmasked_Rev_TableName & " ON " &
                      Global_Variables.Gbl_Masked_TableName & ".Serial_No = " &
                      Global_Variables.Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_no " &
                      "WHERE " & Global_Variables.Gbl_Masked_TableName & ".STCC_W49 LIKE '" & mfrmSTCC & "%'"
        Else
            mStrSQL = "SELECT COUNT(*) AS MyCount " &
                      "FROM " & Global_Variables.Gbl_Masked_TableName & " INNER JOIN " &
                      Global_Variables.Gbl_Unmasked_Rev_TableName & " ON " &
                      Global_Variables.Gbl_Masked_TableName & ".Serial_No = " &
                      Global_Variables.Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_no " &
                      "WHERE " & Global_Variables.Gbl_Masked_TableName & ".STCC LIKE '" & mfrmSTCC & "%'"
        End If

        txt_StatusBox.Text = "Searching for records..."
        Refresh()

        'rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)
        mDataTable = New DataTable
        txt_StatusBox.Text = ""
        Refresh()

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        bolWrite = False
        If mDataTable.Rows(0)(0) = 0 Then
            bolWrite = MsgBox("No records found. Create empty file?", vbYesNo, "Error!")
            If bolWrite = vbYes Then
                File.Create(txt_Output_FilePath.Text).Dispose()
            End If
            GoTo Done
        End If

        mMaxRecs = mDataTable.Rows(0)(0)

        bolWrite = MsgBox("Found " & CStr(mMaxRecs) & " records for STCC" &
            mfrmSTCC &
            ". Are you sure you want to write the data?", vbYesNo)

        If bolWrite = vbYes Then

            If chk_Unmasked.Checked = True Then
                txt_StatusBox.Text = "Processing Unmasked Data..."
            Else
                txt_StatusBox.Text = "Processing Masked Data..."
            End If
            Refresh()

            'Process the data to the output file
            Annual_Sample_By_STCC(cmb_WB_Years.Text,
                mfrmSTCC,
                chk_Unmasked.Checked,
                chk_Unmask_STCC.Checked,
                Gbl_Waybill_Database_Name,
                Gbl_Masked_TableName,
                Gbl_Unmasked_Rev_TableName,
                txt_Output_FilePath.Text,
                913,
                True,
                False)

            GoTo Done
        End If

Done:
        txt_StatusBox.Text = "Done!"
        Refresh()

Endit:

        mDataTable = Nothing

    End Sub

    Private Sub txt_STCC_Code_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_STCC_Code.TextChanged
        Dim mWorkStr As String

        If txt_Output_FilePath.Text <> "" Then
            mWorkStr = Mid(txt_Output_FilePath.Text, 1, InStr(txt_Output_FilePath.Text, "STCC_") + 5)
            mWorkStr = mWorkStr & txt_STCC_Code.Text & "_UnMasked."

        End If
    End Sub
End Class