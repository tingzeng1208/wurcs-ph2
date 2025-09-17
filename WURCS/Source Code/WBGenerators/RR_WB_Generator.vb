Imports System.Data.SqlClient
Public Class RR_WB_Generator

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

        If cmb_URCS_Year.Text <> "" And cmb_Railroad.Text <> "" And txt_Output_FilePath.Text <> "" Then
            btn_Execute.Enabled = True
        End If

    End Sub

    Private Sub RR_WB_Generator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim rst As New ADODB.Recordset
        Dim mDatatable As New DataTable
        Dim mStrSQL As String
        Dim mLooper As Integer

        'Set the form so it centers on the user's screen
        CenterToScreen()

        'Set the default values of the checkboxes
        chk_Unmasked.Checked = False
        chk_RedactOthers.Checked = False
        chk_RedactOthers.Visible = False
        chk_Litigation.Checked = False
        chk_Litigation.Visible = False
        chk_Unmask_STCC.Checked = False
        chk_Unmask_STCC.Visible = False

        OpenSQLConnection(My.Settings.Controls_DB)

        Gbl_WB_Years_TableName = Get_Table_Name_From_SQL("1", "WB_YEARS")
        Gbl_URCS_WAYRRR_TableName = Get_Table_Name_From_SQL("1", "WAYRRR")

        mStrSQL = "SELECT wb_year FROM " & Gbl_WB_Years_TableName

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        ' Load the Year values into the combobox
        For mLooper = 0 To mDatatable.Rows.Count - 1
            cmb_URCS_Year.Items.Add(mDatatable.Rows(mLooper)("wb_year"))
        Next

        'This will populate the railroad combobox
        mDatatable = New DataTable
        mStrSQL = "SELECT RR_NAME FROM " & Gbl_URCS_WAYRRR_TableName & " ORDER BY RR_NAME ASC"

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        For mLooper = 0 To mDatatable.Rows.Count - 1
            cmb_Railroad.Items.Add(mDatatable.Rows(mLooper)("rr_name"))
        Next

        mDatatable = Nothing

    End Sub

    Private Sub chk_Unmasked_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Unmasked.CheckedChanged
        If chk_Unmasked.Checked = True Then
            chk_RedactOthers.Visible = True
            chk_RedactOthers.Checked = False
            chk_Litigation.Visible = True
            chk_Litigation.Checked = False
            chk_Unmask_STCC.Visible = True
            chk_Unmask_STCC.Checked = False
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Masked", "Unmasked")
        Else
            chk_RedactOthers.Visible = False
            chk_RedactOthers.Checked = False
            chk_Litigation.Checked = False
            chk_Litigation.Visible = False
            chk_Unmask_STCC.Checked = False
            chk_Unmask_STCC.Visible = False
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Unmasked", "Masked")
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
            txt_Output_FilePath.Text = fd.SelectedPath.ToString & "WB" & cmb_URCS_Year.Text & "_"
            ' Add the description of the record to the filename
            txt_Output_FilePath.Text = txt_Output_FilePath.Text & cmb_Railroad.Text
            txt_Output_FilePath.Text = txt_Output_FilePath.Text & "_Masked"
        End If
        ' add the extension
        txt_Output_FilePath.Text = txt_Output_FilePath.Text & ".txt"
        ' Clean up the crap in the file name
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, " ", "_")
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, ",", "")
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "&", "And")
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, ".", "")

        If chk_Unmasked.Checked = True Then
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Masked", "Unmasked")
        Else
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Unmasked", "Masked")
        End If

        If chk_RedactOthers.Checked = True Then
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Unmasked", "Unmasked_Redacted")
        End If
SkipIt:

    End Sub

    Private Sub cmb_Railroad_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_Railroad.SelectedIndexChanged
        Dim fd As New FolderBrowserDialog
        Dim mLoc As Integer
        Dim mStr As String

        mLoc = 0
        mStr = ""

        If txt_Output_FilePath.Text <> "" Then
            txt_Output_FilePath.Text = fd.SelectedPath.ToString & "\WB" & cmb_URCS_Year.Text & "_"
            ' Add the description of the record to the filename
            cmb_Railroad.Text = Replace(cmb_Railroad.Text, ".", "")
            cmb_Railroad.Text = Replace(cmb_Railroad.Text, " ", "_")
            cmb_Railroad.Text = Replace(cmb_Railroad.Text, ",", "")
            cmb_Railroad.Text = Replace(cmb_Railroad.Text, "&", "And")
            txt_Output_FilePath.Text = txt_Output_FilePath.Text & cmb_Railroad.Text
            txt_Output_FilePath.Text = txt_Output_FilePath.Text & "_Masked"
        End If


        If chk_Unmasked.Checked = True Then
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Masked", "Unmasked")
        Else
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Unmasked", "Masked")
        End If

        If chk_RedactOthers.Checked = True Then
            txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, "Unmasked", "Unmasked_Redacted")
        End If

    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        Dim rst As New ADODB.Recordset
        Dim mDatatable As New DataTable
        Dim mLooper As Integer
        Dim mStrSQL As String, mstrRR As String, mStrOutLine As String
        Dim mSTCC As String, mSTCC_W49 As String
        Dim mPromptYear As String
        Dim bolWrite As Object
        Dim Rec As Long
        Dim fs, outfile
        Dim mRate_Flg As Integer
        Dim mTotal_Rev As Decimal, mU_Rev As Decimal
        Dim mRR_Revs(8) As Decimal

        Dim mWaybill As Class_Waybill

        'Perform error checking on the form controls
        If cmb_URCS_Year.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo EndIt
        End If

        If cmb_Railroad.Text = "" Then
            MsgBox("You must select a railroad.", vbOKOnly)
            GoTo EndIt
        End If

        If txt_Output_FilePath.Text = "" Then
            MsgBox("You must select a filename to use or create.", vbOKOnly)
            GoTo EndIt
        End If

        OpenSQLConnection(My.Settings.Controls_DB)

        Rec = 0

        'Find out what the AARID number is for the railroad name selected on the form

        mStrSQL = "SELECT * FROM " & Gbl_URCS_WAYRRR_TableName & " WHERE RR_NAME = '" & cmb_Railroad.Text & "'"

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        mstrRR = mDatatable.Rows(0)("AARID")

        ' Get the database and table information for the Waybill and Unmasked tables
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(cmb_URCS_Year.Text, "MASKED")
        Gbl_Unmasked_Rev_TableName = Get_Table_Name_From_SQL(cmb_URCS_Year.Text, "UNMASKED_REV")

        OpenSQLConnection(My.Settings.Waybills_DB)

        mPromptYear = cmb_URCS_Year.Text

        If chk_Unmasked.Checked Then
            mStrSQL = "SELECT count(*) " &
                    "FROM " & Gbl_Masked_TableName & " INNER JOIN " &
                    Gbl_Unmasked_Rev_TableName & " ON " &
                    Gbl_Masked_TableName & ".Serial_No = " &
                    Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_no " &
                    "WHERE (" & Gbl_Masked_TableName & ".report_rr = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".ORR = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR1 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR2 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR3 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR4 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR5 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR6 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".TRR = " & mstrRR & ")"
        Else
            mStrSQL = "SELECT count(*) " &
                    "FROM " & Gbl_Masked_TableName & " " &
                    "WHERE (" & Gbl_Masked_TableName & ".report_rr = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".ORR = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR1 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR2 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR3 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR4 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR5 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".JRR6 = " & mstrRR & ") OR (" &
                    Gbl_Masked_TableName & ".TRR = " & mstrRR & ")"

        End If

        txt_StatusBox.Text = "Searching for records..."
        Refresh()

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        bolWrite = False

        If mDatatable.Rows.Count = 0 Then
            MsgBox("No records found for " & cmb_Railroad.Text, vbOKOnly, "Error!")
            GoTo EndIt
        End If

        bolWrite = MsgBox("Are you sure you want to write the data?", vbYesNo)

        If bolWrite = vbYes Then
            ' Open the output file
            fs = CreateObject("Scripting.FileSystemObject")
            outfile = fs.createtextfile(txt_Output_FilePath.Text, True)

            ' Let the user know what is going on
            If Me.chk_Unmasked.Checked Then
                txt_StatusBox.Text = "Building UnMasked Data File..."
            Else
                txt_StatusBox.Text = "Building Masked Data File..."
            End If
            Refresh()

            ' reset the results
            mDatatable = New DataTable

            ' Modify the query
            mStrSQL = Replace(mStrSQL, "count(*)", "*")
            mStrSQL = mStrSQL & " ORDER BY " & Gbl_Masked_TableName & ".serial_no"

            ' Execute the query
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDatatable)
            End Using

            'Display the blank progress Form
            If chk_Unmasked.Checked Then
                txt_StatusBox.Text = "Exporting Unmasked Records"
            Else
                txt_StatusBox.Text = "Exporting Masked Records"
            End If

            For mLooper = 0 To mDatatable.Rows.Count - 1

                If mLooper Mod 100 = 0 Then
                    txt_StatusBox.Text = "Processing record " & (mLooper + 1).ToString & " of " & mDatatable.Rows.Count
                    Application.DoEvents()
                End If

                'Set the variables to zero
                mTotal_Rev = 0
                For mReset = 1 To 8
                    mRR_Revs(mReset) = 0
                Next


                ' Gather the information that will be manipulated if necessary or used by functions
                mSTCC = mDatatable.Rows(mLooper)("stcc")
                mSTCC_W49 = mDatatable.Rows(mLooper)("stcc_w49")

                If chk_Unmasked.Checked Then
                    mTotal_Rev = mDatatable.Rows(mLooper)("total_unmasked_rev")
                    mRR_Revs(1) = mDatatable.Rows(mLooper)("orr_unmasked_rev")
                    mRR_Revs(2) = mDatatable.Rows(mLooper)("jrr1_unmasked_rev")
                    mRR_Revs(3) = mDatatable.Rows(mLooper)("jrr2_unmasked_rev")
                    mRR_Revs(4) = mDatatable.Rows(mLooper)("jrr3_unmasked_rev")
                    mRR_Revs(5) = mDatatable.Rows(mLooper)("jrr4_unmasked_rev")
                    mRR_Revs(6) = mDatatable.Rows(mLooper)("jrr5_unmasked_rev")
                    mRR_Revs(7) = mDatatable.Rows(mLooper)("jrr6_unmasked_rev")
                    mRR_Revs(8) = mDatatable.Rows(mLooper)("trr_unmasked_rev")
                    mU_Rev = mDatatable.Rows(mLooper)("u_rev_unmasked")
                    mRate_Flg = mDatatable.Rows(mLooper)("rate_flg")
                Else
                    mTotal_Rev = mDatatable.Rows(mLooper)("total_rev")
                    mRR_Revs(1) = mDatatable.Rows(mLooper)("orr_rev")
                    mRR_Revs(2) = mDatatable.Rows(mLooper)("jrr1_rev")
                    mRR_Revs(3) = mDatatable.Rows(mLooper)("jrr2_rev")
                    mRR_Revs(4) = mDatatable.Rows(mLooper)("jrr3_rev")
                    mRR_Revs(5) = mDatatable.Rows(mLooper)("jrr4_rev")
                    mRR_Revs(6) = mDatatable.Rows(mLooper)("jrr5_rev")
                    mRR_Revs(7) = mDatatable.Rows(mLooper)("jrr6_rev")
                    mRR_Revs(8) = mDatatable.Rows(mLooper)("trr_rev")
                    mU_Rev = mDatatable.Rows(mLooper)("u_rev")
                    mRate_Flg = 0
                End If

                If chk_Unmask_STCC.Checked Then
                    'leave the STCC fields as they are
                Else
                    If Mid(Trim(mSTCC), 1, 2) = "19" Then
                        mSTCC = "1900000"
                        If Mid(Trim(mSTCC_W49), 1, 2) = "19" Then
                            mSTCC_W49 = "1900000"
                        End If
                        If Mid(Trim(mSTCC_W49), 1, 2) = "49" Then
                            mSTCC_W49 = "4900000"
                        End If
                    End If
                End If

                'initialize the mwaybill class
                mWaybill = New Class_Waybill

                'Pass the serial_no variable to the class loader
                'resulting in the class being loaded with the values from the current record.
                LoadWaybillClass(mDatatable.Rows(mLooper)("serial_no"), mWaybill, mDatatable, 900)

                'load the unmasked/masked variables into the class
                With mWaybill

                    .STCC_W49 = mSTCC_W49
                    .STCC = mSTCC
                    .Rate_Flg = mRate_Flg

                    If .Report_RR = mstrRR Then
                        .ORR_Rev = mRR_Revs(1)
                        .JRR1_Rev = mRR_Revs(2)
                        .JRR2_Rev = mRR_Revs(3)
                        .JRR3_Rev = mRR_Revs(4)
                        .JRR4_Rev = mRR_Revs(5)
                        .JRR5_Rev = mRR_Revs(6)
                        .JRR6_Rev = mRR_Revs(7)
                        .TRR_Rev = mRR_Revs(8)
                        .U_Rev = mU_Rev
                        .Total_Rev = mTotal_Rev
                    Else
                        If chk_RedactOthers.Checked = True Then
                            .Rate_Flg = 0
                            .Report_RR = 0
                            .Tran_Chrg = 0
                            .Misc_Chrg = 0
                            .Total_VC = 0
                            .Total_Dist = 0
                            .Shortline_Miles = 0
                            .U_Rev = 0
                            .Total_Rev = 0

                            If mstrRR <> .ORR Then
                                .ORR = 0
                                .ORR_Rev = 0
                                .ORR_Dist = 0
                                .RR1_VC = 0
                                .ORR_Alpha = ""
                            Else
                                .ORR_Rev = mRR_Revs(1)
                            End If

                            If mstrRR <> .JRR1 Then
                                .JRR1 = 0
                                .JRR1_Rev = 0
                                .JRR1_Dist = 0
                                .RR2_VC = 0
                                .JRR1_Alpha = ""
                            Else
                                .JRR1_Rev = mRR_Revs(2)
                            End If

                            If mstrRR <> .JRR2 Then
                                .JRR2 = 0
                                .JRR2_Rev = 0
                                .JRR2_Dist = 0
                                .RR3_VC = 0
                                .JRR2_Alpha = ""
                            Else
                                .JRR2_Rev = mRR_Revs(3)
                            End If

                            If mstrRR <> .JRR3 Then
                                .JRR3 = 0
                                .JRR3_Rev = 0
                                .JRR3_Dist = 0
                                .RR4_VC = 0
                                .JRR3_Alpha = ""
                            Else
                                .JRR3_Rev = mRR_Revs(4)
                            End If

                            If mstrRR <> .JRR4 Then
                                .JRR4 = 0
                                .JRR4_Rev = 0
                                .JRR4_Dist = 0
                                .RR5_VC = 0
                                .JRR4_Alpha = ""
                            Else
                                .JRR4_Rev = mRR_Revs(5)
                            End If

                            If mstrRR <> .JRR5 Then
                                .JRR5 = 0
                                .JRR5_Rev = 0
                                .JRR5_Dist = 0
                                .RR6_VC = 0
                                .JRR5_Alpha = ""
                            Else
                                .JRR5_Rev = mRR_Revs(6)
                            End If

                            If mstrRR <> .JRR6 Then
                                .JRR6 = 0
                                .JRR6_Rev = 0
                                .JRR6_Dist = 0
                                .RR7_VC = 0
                                .JRR6_Alpha = ""
                            Else
                                .JRR6_Rev = mRR_Revs(7)
                            End If

                            If mstrRR <> .TRR Then
                                .TRR = 0
                                .TRR_Rev = 0
                                .TRR_Dist = 0
                                .RR8_VC = 0
                                .TRR_Alpha = ""
                            Else
                                .TRR_Rev = mRR_Revs(8)
                            End If

                        End If
                    End If
                End With

                'Build the outfile record string
                mStrOutLine = Build913String(mWaybill)

                'the line is built - write it to file
                outfile.writeline(mStrOutLine)

                'unload the waybill info from the class
                mWaybill = Nothing

                Rec = Rec + 1

            Next

            'close the output file
            outfile.Close()

            txt_StatusBox.Text = "Done!"
            Refresh()
        Else

EndIt:
            txt_StatusBox.Text = "Aborted"
            Refresh()

        End If
    End Sub

    Private Sub btn_Return_To_WBGeneratorsMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_WBGeneratorsMenu.Click
        ' Open the WBGenerators And Utilities Menu Form
        Dim frmNew As New frm_WBGeneratorsAndUtilitiesMenu()
        frmNew.Show()
        ' Close this Menu
        Me.Close()
    End Sub

    Private Sub txt_Output_FilePath_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_Output_FilePath.TextChanged

        If cmb_URCS_Year.Text <> "" And cmb_Railroad.Text <> "" And txt_Output_FilePath.Text <> "" Then
            btn_Execute.Enabled = True
        End If

    End Sub
End Class
