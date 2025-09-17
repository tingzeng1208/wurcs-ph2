Imports System.Data.SqlClient
Public Class State_WB_Generator

    Private Sub State_WB_Generator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        CenterToScreen()

        'Set the default values of the checkboxes
        RadioButton_913.Checked = False
        RadioButton_CSV.Checked = False
        CheckBox_Unmasked.Checked = False


        ' Load the Year combobox from the SQL database
        Gbl_Controls_Database_Name = "URCS_Controls"

        gbl_SQLConnection = New SqlConnection
        gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(Gbl_Controls_Database_Name)
        gbl_SQLConnection.Open()

        mDataTable = Get_States_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            lbx_States.Items.Add(mDataTable.Rows(mLooper)("ch_state"))
        Next

        mDataTable = Get_Waybill_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            cmb_WB_Year.Items.Add(mDataTable.Rows(mLooper)("wb_year").ToString)
        Next



    End Sub

    Private Sub btn_Return_To_WBGeneratorsMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_WBGeneratorsMenu.Click
        ' Open the WBGenerators And Utilities Menu Form
        Dim frmNew As New frm_WBGeneratorsAndUtilitiesMenu()
        frmNew.Show()
        ' Close this Menu
        Me.Close()
    End Sub

    Private Sub btn_Output_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New FolderBrowserDialog
        Dim Format_Selected As Boolean, Year_Selected As Boolean, State_Selected As Boolean
        Dim mLooper As Integer
        Dim mworkstr As String

        Year_Selected = False
        Format_Selected = False
        State_Selected = False

        If cmb_WB_Year.Text <> "" Then
            Year_Selected = True
        End If

        If RadioButton_913.Checked = True Or RadioButton_CSV.Checked = True Then
            Format_Selected = True
        End If

        If lbx_States.SelectedItems.Count > 0 Then
            State_Selected = True
        End If

        If Year_Selected = False Then
            MsgBox("You must set/select the Waybill Year before selecting an output file.", vbOKOnly, "Error!")
            GoTo SkipIt
        End If

        If Format_Selected = False Then
            MsgBox("You must set/select the record format (913 or CSV) before selecting an output file.", vbOKOnly, "Error!")
            GoTo SkipIt
        End If

        If State_Selected = False Then
            MsgBox("You must set/select at least one state value before selecting an output file.", vbOKOnly, "Error!")
            GoTo SkipIt
        End If

        fd.Description = "Select the location in which you want the output report placed."

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            mworkstr = fd.SelectedPath.ToString & "\WB" & cmb_WB_Year.Text
            For mLooper = 1 To lbx_States.SelectedItems.Count
                mworkstr = mworkstr & "_" & Trim(lbx_States.SelectedItems.Item(mLooper - 1))
            Next
            mworkstr = mworkstr & "_"
            ' Add the description of the record to the filename
            If RadioButton_913.Checked Then
                mworkstr = mworkstr & "913"
                ' add the description & extension
                mworkstr = mworkstr & "_Masked.txt"
            End If
            If RadioButton_CSV.Checked Then
                mworkstr = mworkstr & "CSV"
                ' add the description & extension
                mworkstr = mworkstr & "_Masked.csv"
            End If
            If CheckBox_Unmasked.Checked = True Then
                mworkstr = Replace(mworkstr, "_Masked", "_UnMasked")
            Else
                mworkstr = Replace(mworkstr, "_UnMasked", "_Masked")
            End If

            txt_Output_FilePath.Text = mworkstr

        End If


SkipIt:

    End Sub

    Private Sub Set_FilePathName(ByVal mPath As String)

        ' Add the description of the record to the filename
        If RadioButton_913.Checked Then
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "_CSV_", "_913_")
        End If
        If RadioButton_CSV.Checked Then
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "_913_", "_CSV_")
        End If

        If CheckBox_Unmasked.Checked = True Then
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "_Masked", "_UnMasked")
        Else
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "_UnMasked", "_Masked")
        End If

    End Sub

    Private Sub RadioButton_913_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_913.CheckedChanged
        Set_FilePathName(txt_Output_FilePath.Text)
    End Sub

    Private Sub RadioButton_CSV_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_CSV.CheckedChanged
        Set_FilePathName(txt_Output_FilePath.Text)
    End Sub

    Private Sub cmb_URCS_Year_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_WB_Year.SelectedIndexChanged

        Dim mLoc As Integer

        mLoc = InStr(txt_Output_FilePath.Text, "WB")
        If txt_Output_FilePath.Text <> "" Then
            Mid(txt_Output_FilePath.Text, mLoc + 2, 4) = cmb_WB_Year.Text
        End If
    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        Dim bolWrite As Object
        Dim mIndex As Long, mMaxRecs As Decimal

        'Perform error checks on form data
        If cmb_WB_Year.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly, "Error!")
            GoTo EndIt
        End If

        If lbx_States.SelectedItems.Count = 0 Then
            MsgBox("You must select the State(s) value.", vbOKOnly, "Error!")
            GoTo EndIt
        End If

        If txt_Output_FilePath.Text = "" Then
            MsgBox("You must select a filename to use or create.", vbOKOnly, "Error!")
            GoTo EndIt
        End If

        'Get the database_name and table_name for the Waybills table from the database
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(cmb_WB_Year.Text, "MASKED")
        Gbl_Unmasked_Rev_TableName = Get_Table_Name_From_SQL(cmb_WB_Year.Text, "UNMASKED_REV")

        ' Open the connection
        OpenADOConnection(Gbl_Waybill_Database_Name)

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        txt_StatusBox.Text = "Searching for records..."
        Refresh()

        mMaxRecs = Count_States_Waybills(cmb_WB_Year.Text, lbx_States.SelectedItems)

        System.Windows.Forms.Cursor.Current = Cursors.Default
        txt_StatusBox.Text = ""
        Refresh()
        bolWrite = False

        If mMaxRecs = 0 Then
            MsgBox("No records found for " & cmb_WB_Year.Text)
        Else
            bolWrite = MsgBox("The number of records returned is " & mMaxRecs &
                ". Are you sure you want to write the data?", vbYesNo)

            If bolWrite = vbYes Then

                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

                mIndex = 0
                txt_StatusBox.Text = "Writing Waybill Data.  Please Wait..."
                Refresh()
                If RadioButton_CSV.Checked Then
                    Annual_Sample_By_Selected_States_Writer(Gbl_Waybill_Database_Name,
                                                  Global_Variables.Gbl_Masked_TableName,
                                                  Global_Variables.Gbl_Unmasked_Rev_TableName,
                                                  lbx_States.SelectedItems,
                                                  CheckBox_Unmasked.Checked,
                                                  mMaxRecs,
                                                  txt_Output_FilePath.Text,
                                                  "CSV")
                Else
                    Annual_Sample_By_Selected_States_Writer(Gbl_Waybill_Database_Name,
                                                  Global_Variables.Gbl_Masked_TableName,
                                                  Global_Variables.Gbl_Unmasked_Rev_TableName,
                                                  lbx_States.SelectedItems,
                                                  CheckBox_Unmasked.Checked,
                                                  mMaxRecs,
                                                  txt_Output_FilePath.Text,
                                                  "913")
                End If
            End If

            System.Windows.Forms.Cursor.Current = Cursors.Default
            txt_StatusBox.Text = "Done!"
            Refresh()
        End If

EndIt:

    End Sub

    Private Sub CheckBox_Unmasked_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_Unmasked.CheckedChanged

        If CheckBox_Unmasked.Checked = True Then
            If MsgBox("Are you sure you want to unmask the revenue values?", vbYesNo, "CAUTION!") = MsgBoxResult.No Then
                CheckBox_Unmasked.Checked = False
            End If
        End If

        If CheckBox_Unmasked.Checked = True Then
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "_Masked", "_UnMasked")
        Else
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "_UnMasked", "_Masked")
        End If

    End Sub

    Private Sub lbx_States_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbx_States.SelectedIndexChanged
        Dim mlooper As Integer
        Dim mworkstr As String

        If Len(txt_Output_FilePath.Text) > 0 Then
            mworkstr = txt_Output_FilePath.Text

            ' keep the directory, but remove anything after "\WByyyy"
            mworkstr = Mid(txt_Output_FilePath.Text, 1, InStr(txt_Output_FilePath.Text, "\WB") + 6)
            'Add the selected state codes
            For mlooper = 1 To lbx_States.SelectedItems.Count
                mworkstr = mworkstr & "_" & Trim(lbx_States.SelectedItems.Item(mlooper - 1))
            Next
            'Add the format
            mworkstr = mworkstr & "_"
            ' Add the description of the record to the filename
            If RadioButton_913.Checked Then
                mworkstr = mworkstr & "900"
                ' add the description & extension
                mworkstr = mworkstr & "_Masked.txt"
            End If
            If RadioButton_CSV.Checked Then
                mworkstr = mworkstr & "CSV"
                ' add the description & extension
                mworkstr = mworkstr & "_Masked.csv"
            End If
            'Indicate whether masked/unmasked
            If CheckBox_Unmasked.Checked = True Then
                mworkstr = Replace(mworkstr, "_Masked", "_UnMasked")
            Else
                mworkstr = Replace(mworkstr, "_UnMasked", "_Masked")
            End If

            txt_Output_FilePath.Text = mworkstr

        End If

    End Sub
End Class