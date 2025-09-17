Imports System.Data.SqlClient
Imports System.Text
Public Class Annual_WB_Generator

    Private Sub Annual_WB_Generator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim rst As ADODB.Recordset
        Dim mStrSQL As String

        'Set the form so it centers on the user's screen
        CenterToScreen()

        'Set the default values of the checkboxes
        RadioButton_570.Checked = False
        RadioButton_913.Checked = False

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        OpenADOConnection(Get_Database_Name_From_SQL("1", "WB_Years"))

        rst = SetRST()
        mStrSQL = "SELECT wb_year FROM " & Get_Table_Name_From_SQL("1", "WB_Years") & " ORDER BY wb_year DESC"

        rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)
        rst.MoveFirst()

        ' Load the Year values into the combobox
        Do While Not rst.EOF
            cmb_WB_Year.Items.Add(rst.Fields("wb_year").Value)
            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

    End Sub

    Private Sub btn_Return_To_WBGeneratorsMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_WBGeneratorsMenu.Click

        ' Open the WBGenerators And Utilities Menu Form
        Dim frmNew As New frm_WBGeneratorsAndUtilitiesMenu()
        frmNew.Show()
        ' Close this Menu
        Me.Close()

    End Sub

    Private Sub cmb_URCS_Year_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_WB_Year.SelectedValueChanged

        Dim mArrayPos As Integer

        ' locate the entry position in the tablename array
        mArrayPos = Array1DFindFirst(WBYears, Me.cmb_WB_Year.Text)

        If WBMaskedOnly(mArrayPos) = True Then
            UnmaskedDataCheck.Checked = False
            UnmaskedSTCCCheck.Checked = False

            'If mIsUnmaskedReader = True Then
            UnmaskedDataCheck.Visible = False
            UnmaskedSTCCCheck.Visible = False
        Else
            UnmaskedDataCheck.Visible = True
            UnmaskedSTCCCheck.Visible = True
        End If

    End Sub

    Private Sub RadioButton_913_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_913.CheckedChanged

        UnmaskedDataCheck.Enabled = True
        UnmaskedDataCheck.Visible = True
        UnmaskedSTCCCheck.Enabled = True
        UnmaskedSTCCCheck.Visible = True

        ' Update the output file string
        If InStr(txt_Output_DirPath.Text, "_570_") > 0 Then
            txt_Output_DirPath.Text = Replace(txt_Output_DirPath.Text, "_570_", "_913_")
        ElseIf InStr(txt_Output_DirPath.Text, "_CSV_") > 0 Then
            txt_Output_DirPath.Text = Replace(txt_Output_DirPath.Text, "_CSV_", "_913_")
        End If

    End Sub

    Private Sub RadioButton_570_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_570.CheckedChanged

        UnmaskedDataCheck.Enabled = False
        UnmaskedDataCheck.Visible = False
        UnmaskedSTCCCheck.Enabled = False
        UnmaskedSTCCCheck.Visible = False

        ' Update the output file string
        If InStr(txt_Output_DirPath.Text, "_913_") > 0 Then
            txt_Output_DirPath.Text = Replace(txt_Output_DirPath.Text, "_913_", "_570_")
        ElseIf InStr(txt_Output_DirPath.Text, "_CSV_") > 0 Then
            txt_Output_DirPath.Text = Replace(txt_Output_DirPath.Text, "_CSV_", "_570_")
        End If

    End Sub

    Private Sub RadioButton_CSV_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        UnmaskedDataCheck.Enabled = False
        UnmaskedDataCheck.Visible = False
        UnmaskedSTCCCheck.Enabled = True
        UnmaskedSTCCCheck.Visible = True

        ' Update the output file string
        If InStr(txt_Output_DirPath.Text, "_913_") > 0 Then
            txt_Output_DirPath.Text = Replace(txt_Output_DirPath.Text, "_913_", "_CSV_")
        ElseIf InStr(txt_Output_DirPath.Text, "_570_") > 0 Then
            txt_Output_DirPath.Text = Replace(txt_Output_DirPath.Text, "_570_", "_CSV_")
        End If


    End Sub

    Private Sub UnmaskedSTCCCheck_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UnmaskedSTCCCheck.CheckedChanged

        If UnmaskedSTCCCheck.Checked = True Then
            If MsgBox("Are you SURE you want to produce unmasked 19 Series STCC data?", vbYesNo) = vbYes Then
                'Ok, they want the unmasked data
                MaskedRevenueAndLocationsCheck.Visible = False
                MaskedRevenueAndLocationsCheck.Checked = False
                MaskedRevenueAndLocationsCheck.Enabled = False
                ' Update the output file string
                If UnmaskedDataCheck.Checked = True Or UnmaskedSTCCCheck.Checked = True Then
                    txt_Output_DirPath.Text = Replace(txt_Output_DirPath.Text, "_Masked", "_Unmasked")
                    Refresh()
                End If
            End If
        Else
            If UnmaskedDataCheck.Checked = False Then
                MaskedRevenueAndLocationsCheck.Visible = True
                MaskedRevenueAndLocationsCheck.Checked = False
                MaskedRevenueAndLocationsCheck.Enabled = True
                ' Update the output file string
                If UnmaskedDataCheck.Checked = False And UnmaskedSTCCCheck.Checked = False Then
                    txt_Output_DirPath.Text = Replace(txt_Output_DirPath.Text, "_Unmasked", "_Masked")
                    Refresh()
                End If
            End If
        End If

    End Sub

    Private Sub btn_Output_Folder_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Output_File_Entry.Click

        Dim fd As New FolderBrowserDialog
        Dim Format_Selected As Boolean, Year_Selected As Boolean

        Year_Selected = False
        Format_Selected = False

        If cmb_WB_Year.Text <> "" Then
            Year_Selected = True
        End If

        If RadioButton_913.Checked = True Or RadioButton_570.Checked = True Then
            Format_Selected = True
        End If

        If Year_Selected = False Then
            MsgBox("You must set/select the Waybill Year before selecting an output file.", vbOKOnly, "Error!")
            GoTo SkipIt
        End If

        If Format_Selected = False Then
            MsgBox("You must set/select the record format (913, 570 or CSV) before selecting an output file.", vbOKOnly, "Error!")
            GoTo SkipIt
        End If

        fd.Description = "Select the location in which you want the output report placed."

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Output_DirPath.Text = fd.SelectedPath.ToString & "\WB" & Me.cmb_WB_Year.Text & "_"
            ' Add the description of the record to the filename
            If RadioButton_913.Checked Then
                txt_Output_DirPath.Text = txt_Output_DirPath.Text & "913"
            End If
            If RadioButton_570.Checked Then
                txt_Output_DirPath.Text = txt_Output_DirPath.Text & "570"
            End If
            ' Is the file unmasked data?  If so, add that as well
            If UnmaskedDataCheck.Checked = True Or UnmaskedSTCCCheck.Checked = True Then
                txt_Output_DirPath.Text = txt_Output_DirPath.Text & "_Unmasked"
            Else
                txt_Output_DirPath.Text = txt_Output_DirPath.Text & "_Masked"
            End If
            ' add the extension
            txt_Output_DirPath.Text = txt_Output_DirPath.Text & ".txt"

        End If

SkipIt:

    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        Dim mDatatable As New DataTable
        Dim mStrSQL As String
        Dim Format_Selected As Boolean = False
        Dim mRecFormat As Integer
        Dim bolWrite As Boolean
        Dim mMaxRecs As Single

        Dim mStrOutLine As String, mLooper As Integer
        Dim mSTCC As String, mSTCC_W49 As String, mOrdnance_STCCs() As String = {""}
        'Dim Rec As Long
        Dim outfile As StreamWriter
        Dim mRate_Flg As Integer
        Dim mTotal_Rev As Single, mORR_Rev As Single, mJRR1_Rev As Single
        Dim mJRR2_Rev As Single, mJRR3_Rev As Single, mJRR4_Rev As Single
        Dim mJRR5_Rev As Single, mJRR6_Rev As Single, mTRR_Rev As Single
        Dim mU_Rev As Single
        Dim mDuplicateSerials As Boolean
        Dim mSQLOffset As Integer = 0
        Dim mSQLChunkSize As Integer = 10000

        ' Determine if the masked table is one that is known to contain duplicate serial numbers
        If InStr(cmb_WB_Year.Text, "1984") > 0 Or
            InStr(cmb_WB_Year.Text, "1985") > 0 Or
            InStr(cmb_WB_Year.Text, "1994") > 0 Or
            InStr(cmb_WB_Year.Text, "1995") > 0 Then
            mDuplicateSerials = True
        Else
            mDuplicateSerials = False
        End If

        If RadioButton_913.Checked = True Or RadioButton_570.Checked = True Then
            Format_Selected = True
        End If

        If Format_Selected = False Then
            MsgBox("You must set/select the record format (913 or 570.)", vbOKOnly, "Error!")
            GoTo SkipIt
        End If

        'Perform error checking for the form controls
        If cmb_WB_Year.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo SkipIt
        End If

        If txt_Output_DirPath.Text = "" Then
            MsgBox("You must select a filename to use or create.", vbOKOnly)
            GoTo SkipIt
        End If

        If My.Computer.FileSystem.FileExists(txt_Output_DirPath.Text) Then
            If MsgBox("File already exists!  Overwrite?", vbYesNo, "Warning!") = vbYes Then
                My.Computer.FileSystem.DeleteFile(txt_Output_DirPath.Text)
            Else
                GoTo SkipIt
            End If
        End If

        ' Load the ordbnance STCCs into the array
        OpenSQLConnection(My.Settings.Controls_DB)
        gbl_Table_Name = Get_Table_Name_From_SQL("1", "ORDNANCE_STCCS")

        OpenSQLConnection(My.Settings.Waybills_DB)
        mStrSQL = "SELECT * FROM " & gbl_Table_Name

        mDatatable = New DataTable

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        For mLooper = 0 To mDatatable.Rows.Count - 1
            ReDim Preserve mOrdnance_STCCs(UBound(mOrdnance_STCCs) + 1)
            mOrdnance_STCCs(UBound(mOrdnance_STCCs)) = mDatatable.Rows(mLooper)("STCC")
        Next

        OpenSQLConnection(My.Settings.Waybills_DB)

        'table_names
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(cmb_WB_Year.Text, "Masked")

        ' Verify that this table contains Transaction_No field.
        ' If not, add it to the record.
        If Column_Exist(gbl_Table_Name, "Tracking_No") = False Then
            Column_Add(gbl_Table_Name, "Tracking_No", "BigInt")
        End If

        ' By default, this will be in the same database as the masked data
        If UnmaskedDataCheck.Checked Then
            Gbl_Unmasked_Rev_TableName = Get_Table_Name_From_SQL(cmb_WB_Year.Text, "Unmasked_Rev")
        End If

        ' Find out how many records we're dealing with
        txt_StatusBox.Text = "Searching for records..."
        Refresh()

        mStrSQL = Build_Simple_Count_Records_SQL_Statement(Gbl_Masked_TableName)

        mDatatable = New DataTable

        'Make sure we're in the Waybills database
        OpenSQLConnection(My.Settings.Waybills_DB)

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        mMaxRecs = mDatatable.Rows(0)(0)

        bolWrite = False
        bolWrite = MsgBox("The number of records returned is " & mMaxRecs.ToString &
            ". Are you sure you want to write the data?", vbYesNo)

        If bolWrite = True Then

            txt_StatusBox.Text = "Fetching data from SQL Server - Please wait..."
            Refresh()

            If RadioButton_570.Checked = True Then
                mRecFormat = 570
            ElseIf RadioButton_913.Checked = True Then
                mRecFormat = 913
            Else
                mRecFormat = 9999   'CSV
            End If

            ' Determine if the masked table is one that is known to contain duplicate serial numbers
            If InStr(cmb_WB_Year.Text, "1984") > 0 Or
            InStr(cmb_WB_Year.Text, "1985") > 0 Or
            InStr(cmb_WB_Year.Text, "1994") > 0 Or
            InStr(cmb_WB_Year.Text, "1995") > 0 Then
                mDuplicateSerials = True
            Else
                mDuplicateSerials = False
            End If

            Dim mWaybill As Class_Waybill

            ' Open the output file
            outfile = My.Computer.FileSystem.OpenTextFileWriter(txt_Output_DirPath.Text, True, Encoding.ASCII)

            'We should be all set to begin the SQL work now
            mStrSQL = ""

            Do While True
                If UnmaskedDataCheck.Checked = False Then
                    mStrSQL = "SELECT * FROM " & Gbl_Masked_TableName & " ORDER BY serial_no ASC OFFSET " & mSQLOffset.ToString & " ROWS FETCH NEXT " & mSQLChunkSize.ToString & " ROWS ONLY"
                Else
                    Select Case mDuplicateSerials
                        Case True
                            mStrSQL = "SELECT " & Gbl_Masked_TableName & ".*, " &
                        Gbl_Unmasked_Rev_TableName & ".* " &
                        "FROM " & Gbl_Masked_TableName & " INNER JOIN " &
                        Gbl_Unmasked_Rev_TableName & " ON " &
                        Gbl_Masked_TableName & ".Serial_No = " &
                        Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_no AND " &
                        Gbl_Masked_TableName & ".WB_Num = " &
                        Gbl_Unmasked_Rev_TableName & ".Unmasked_WB_Num" &
                        " ORDER BY " & Gbl_Masked_TableName & ".Serial_No" &
                        " OFFSET " & mSQLOffset.ToString & " ROWS FETCH NEXT " & mSQLChunkSize.ToString & " ROWS ONLY"
                        Case False
                            mStrSQL = "SELECT " & Gbl_Masked_TableName & ".*, " &
                        Gbl_Unmasked_Rev_TableName & ".* " &
                        "FROM " & Gbl_Masked_TableName & " INNER JOIN " &
                        Gbl_Unmasked_Rev_TableName & " ON " &
                        Gbl_Masked_TableName & ".Serial_No = " &
                        Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_no" &
                        " ORDER BY " & Gbl_Masked_TableName & ".Serial_No" &
                        " OFFSET " & mSQLOffset.ToString & " ROWS FETCH NEXT " & mSQLChunkSize.ToString & " ROWS ONLY"
                    End Select

                End If

                ' Execute the SQL command to fetch the records to the DataTable

                Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                    mDatatable = New DataTable
                    daAdapter.Fill(mDatatable)
                End Using

                If mDatatable.Rows.Count = 0 Then
                    Exit Do
                End If

                mStrOutLine = ""

                For mLooper = 0 To mDatatable.Rows.Count - 1

                    'Reset the local variables
                    mTotal_Rev = 0
                    mORR_Rev = 0
                    mJRR1_Rev = 0
                    mJRR2_Rev = 0
                    mJRR3_Rev = 0
                    mJRR4_Rev = 0
                    mJRR5_Rev = 0
                    mJRR6_Rev = 0
                    mTRR_Rev = 0
                    mU_Rev = 0

                    'Load the Masked Revenues into the variables
                    mTotal_Rev = mDatatable.Rows(mLooper)("total_rev")
                    mORR_Rev = mDatatable.Rows(mLooper)("orr_rev")
                    mJRR1_Rev = mDatatable.Rows(mLooper)("jrr1_rev")
                    mJRR2_Rev = mDatatable.Rows(mLooper)("jrr2_rev")
                    mJRR3_Rev = mDatatable.Rows(mLooper)("jrr3_rev")
                    mJRR4_Rev = mDatatable.Rows(mLooper)("jrr4_rev")
                    mJRR5_Rev = mDatatable.Rows(mLooper)("jrr5_rev")
                    mJRR6_Rev = mDatatable.Rows(mLooper)("jrr6_rev")
                    mTRR_Rev = mDatatable.Rows(mLooper)("trr_rev")
                    mU_Rev = mDatatable.Rows(mLooper)("u_rev")

                    'Set the rate flag to zero
                    mRate_Flg = 0

                    If mLooper Mod 100 = 0 Then
                        txt_StatusBox.Text = "Processing record " & (mLooper + mSQLOffset).ToString & " of " & mMaxRecs.ToString & " - " &
                            (((mLooper + mSQLOffset) / mMaxRecs) * 100).ToString("N1") & "%"
                        Application.DoEvents()
                    End If

                    mSTCC = mDatatable.Rows(mLooper)("stcc")
                    mSTCC_W49 = mDatatable.Rows(mLooper)("stcc_w49")

                    If UnmaskedDataCheck.Checked = True Then
                        'Get unmasked values
                        mTotal_Rev = mDatatable.Rows(mLooper)("Total_Unmasked_Rev")
                        mORR_Rev = mDatatable.Rows(mLooper)("ORR_Unmasked_Rev")
                        mJRR1_Rev = mDatatable.Rows(mLooper)("JRR1_Unmasked_Rev")
                        mJRR2_Rev = mDatatable.Rows(mLooper)("JRR2_Unmasked_Rev")
                        mJRR3_Rev = mDatatable.Rows(mLooper)("JRR3_Unmasked_Rev")
                        mJRR4_Rev = mDatatable.Rows(mLooper)("JRR4_Unmasked_Rev")
                        mJRR5_Rev = mDatatable.Rows(mLooper)("JRR5_Unmasked_Rev")
                        mJRR6_Rev = mDatatable.Rows(mLooper)("JRR6_Unmasked_Rev")
                        mTRR_Rev = mDatatable.Rows(mLooper)("TRR_Unmasked_Rev")
                        mU_Rev = mDatatable.Rows(mLooper)("U_Rev_Unmasked")
                        mRate_Flg = mDatatable.Rows(mLooper)("rate_flg")
                    End If

                    ' If the format is 570, the revenues are masked, but the rate_flg is not,
                    If RadioButton_570.Checked Then
                        mRate_Flg = mDatatable.Rows(mLooper)("rate_flg")
                    End If

                    If Array1DFindFirst(mOrdnance_STCCs, mSTCC_W49) <> 0 Then
                        ' we mask ordnance STCCs
                        If Mid(mSTCC, 1, 2) = "19" Then
                            mSTCC = "1900000"
                        End If
                        Select Case Mid(mSTCC_W49, 1, 2)
                            Case "19"
                                mSTCC_W49 = "1900000"
                            Case "49"
                                mSTCC_W49 = "4900000"
                        End Select
                    Else
                        ' As a fail safe, we will apply the change to any possibilty of a STCC 19/49 combination
                        If Mid(mSTCC, 1, 2) = "19" Then
                            mSTCC = "1900000"
                            Select Case Mid(mSTCC_W49, 1, 2)
                                Case "19"
                                    mSTCC_W49 = "1900000"
                                Case "49"
                                    mSTCC_W49 = "4900000"
                            End Select
                        End If
                    End If

                    ' Initialize the mWaybill Class
                    mWaybill = New Class_Waybill

                    With mWaybill
                        'Load the unmasked/masked variables to the class
                        .STCC_W49 = mSTCC_W49
                        .Rate_Flg = mRate_Flg
                        .STCC = mSTCC
                        If chk_ZeroRevenues.Checked = True Then
                            .Total_Rev = 0
                            .ORR_Rev = 0
                            .JRR1_Rev = 0
                            .JRR2_Rev = 0
                            .JRR3_Rev = 0
                            .JRR4_Rev = 0
                            .JRR5_Rev = 0
                            .JRR6_Rev = 0
                            .TRR_Rev = 0
                            .U_Rev = 0
                        Else
                            .Total_Rev = mTotal_Rev
                            .ORR_Rev = mORR_Rev
                            .JRR1_Rev = mJRR1_Rev
                            .JRR2_Rev = mJRR2_Rev
                            .JRR3_Rev = mJRR3_Rev
                            .JRR4_Rev = mJRR4_Rev
                            .JRR5_Rev = mJRR5_Rev
                            .JRR6_Rev = mJRR6_Rev
                            .TRR_Rev = mTRR_Rev
                            .U_Rev = mU_Rev
                        End If

                    End With

                    'Pass the serial_no variable to the class loader
                    'resulting in the class being loaded with the rest of
                    'the values from the current resultset record.

                    LoadWaybillClass(mDatatable.Rows(mLooper)("serial_no"), mWaybill, mDatatable, 913)

                    'Build the outfile Record string depending on selection
                    If RadioButton_913.Checked Then
                        mStrOutLine = Build913String(mWaybill)
                    End If

                    If RadioButton_570.Checked Then
                        mStrOutLine = Build570String(mWaybill)
                    End If

                    outfile.WriteLine(mStrOutLine)

                Next

                'Increment offset for next chunk
                mSQLOffset = mSQLOffset + 10000

            Loop

            'close the output file
            outfile.Flush()
            outfile.Close()

            mDatatable = Nothing
            mWaybill = Nothing

            txt_StatusBox.Text = "Done!"
            Refresh()

        Else

            txt_StatusBox.Text = "Aborted."
            Refresh()

        End If



SkipIt:

    End Sub

    Private Sub UnmaskedDataCheck_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UnmaskedDataCheck.CheckedChanged

        If UnmaskedDataCheck.Checked = True Then
            If MsgBox("Are you SURE you want to produce unmasked revenue data?", vbYesNo) = vbYes Then
                'Ok, they want the unmasked data so they can't have PUWS data
                MaskedRevenueAndLocationsCheck.Visible = False
                MaskedRevenueAndLocationsCheck.Checked = False
                MaskedRevenueAndLocationsCheck.Enabled = False
                ' Update the output file string
                If UnmaskedDataCheck.Checked = True Or UnmaskedSTCCCheck.Checked = True Then
                    txt_Output_DirPath.Text = Replace(txt_Output_DirPath.Text, "_Masked", "_Unmasked")
                    Refresh()
                End If
            End If
        Else
            MaskedRevenueAndLocationsCheck.Visible = True
            MaskedRevenueAndLocationsCheck.Checked = False
            MaskedRevenueAndLocationsCheck.Enabled = True
            ' Update the output file string
            If UnmaskedDataCheck.Checked = False And UnmaskedSTCCCheck.Checked = False Then
                txt_Output_DirPath.Text = Replace(txt_Output_DirPath.Text, "_Unmasked", "_Masked")
                Refresh()
            End If
        End If

    End Sub

    Private Sub chk_ZeroRevenues_CheckedChanged(sender As Object, e As EventArgs) Handles chk_ZeroRevenues.CheckedChanged

    End Sub
End Class