Imports System.Text
Public Class PUWS_Generator

    Private Sub PUWS_Generator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'This routine retrieves the choices of year values available from the
        'SQL lookup table and loads the choices into the year combo box control
        'on the form.

        Dim rst As ADODB.Recordset
        Dim mStrSQL As String
        Dim mIndex As Integer

        CenterToScreen()
        mIndex = 0

        Gbl_Controls_Database_Name = Get_Database_Name_From_SQL("1", "Wb_Years")
        OpenADOConnection(Gbl_Controls_Database_Name)

        'Build the SQL statement
        gbl_Table_Name = Get_Table_Name_From_SQL("1", "WB_Years")
        mStrSQL = "SELECT wb_year FROM " & gbl_Table_Name

        'Open recordset and add to the combobox
        rst = SetRST()
        rst.Open(mStrSQL, gbl_ADOConnection)

        'Load the values from the resultset into the combobox control on the
        'form.
        rst.MoveFirst()
        Do While Not rst.EOF
            cmb_URCSYear.Items.Add(rst.Fields("wb_year").Value)
            rst.MoveNext()
            mIndex = mIndex + 1
        Loop

        rst.Close()
        rst = Nothing

    End Sub

    Private Sub btn_Output_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = False
        fd.Filter = "Txt Files|*.txt|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Output_FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        Dim rst As ADODB.Recordset
        Dim mStrSQL As String
        Dim mWriteIt As Integer
        Dim mMaxRecs As Single, mThisRec As Single
        Dim outfile As System.IO.StreamWriter
        Dim mOutStr As StringBuilder

        'Perform error checking for the form controls
        If cmb_URCSYear.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo SkipIt
        End If

        If txt_Output_FilePath.Text = "" Then
            MsgBox("You must select a filename to use or create.", vbOKOnly)
            GoTo SkipIt
        End If

        'Get the database_name and table_name for the URCS_Years value from the database
        Gbl_Waybill_Database_Name = Get_Database_Name_From_SQL(cmb_URCSYear.Text, "PUWS_MASKED")
        Gbl_PUWS_Masked_Tablename = Get_Table_Name_From_SQL(cmb_URCSYear.Text, "PUWS_MASKED")

        OpenADOConnection(Gbl_Waybill_Database_Name)

        rst = SetRST()

        mStrSQL = "Select * FROM " & Gbl_PUWS_Masked_Tablename

        txt_StatusBox.Text = "Fetching data from server..."
        Refresh()

        ' Execute the query
        rst.Open(mStrSQL, gbl_ADOConnection)

        If rst.RecordCount < 1 Then
            MsgBox("No records found for " & cmb_URCSYear.Text, vbOKOnly, "Error!")
            txt_StatusBox.Text = "Done!"
            Refresh()
            GoTo SkipIt
        End If

        mMaxRecs = rst.RecordCount
        mThisRec = 1

        mWriteIt = MsgBox("The number of records returned is " & mMaxRecs.ToString &
            ". Are you sure you want to write the data?", vbYesNo)

        If mWriteIt = vbYes Then

            ' open the output filestream
            outfile = My.Computer.FileSystem.OpenTextFileWriter(txt_Output_FilePath.Text, False, System.Text.ASCIIEncoding.ASCII)

            rst.MoveFirst()

            Do While Not rst.EOF

                If mThisRec Mod 100 = 0 Then
                    txt_StatusBox.Text = CStr(Math.Round((mThisRec / rst.RecordCount) * 100, 1)) & "% - Writing Records to File"
                    Refresh()
                    Application.DoEvents()
                End If

                'Build the output record (without the serial_no!)
                With rst
                    mOutStr = New StringBuilder(247)
                    mOutStr.Append(Field_Left(.Fields("wb_date").Value, 6))
                    mOutStr.Append(Field_Left(.Fields("acct_period").Value, 4))
                    mOutStr.Append(Field_Right(.Fields("u_cars").Value, 4))
                    mOutStr.Append(Field_Left(.Fields("car_own").Value, 1))
                    mOutStr.Append(Field_Left(.Fields("car_typ").Value, 4))
                    mOutStr.Append(Field_Left(.Fields("mech").Value, 4))
                    mOutStr.Append(Field_Right(.Fields("stb_car_typ").Value, 2))
                    mOutStr.Append(Field_Left(.Fields("tofc_serv_code").Value, 3))
                    mOutStr.Append(Field_Right(.Fields("u_tc_units").Value, 4))
                    mOutStr.Append(Field_Left(.Fields("tofc_own_code").Value, 1))
                    mOutStr.Append(Field_Left(.Fields("tofc_unit_type").Value, 1))
                    mOutStr.Append(Field_Left(.Fields("haz_bulk").Value, 1))
                    mOutStr.Append(Field_Left(.Fields("stcc").Value, 5))
                    mOutStr.Append(Field_Right(.Fields("bill_wght_tons").Value, 7))
                    mOutStr.Append(Field_Right(.Fields("act_wght").Value, 7))
                    mOutStr.Append(Field_Right(.Fields("u_rev").Value, 9))
                    mOutStr.Append(Field_Right(.Fields("tran_chrg").Value, 9))
                    mOutStr.Append(Field_Right(.Fields("misc_chrg").Value, 9))
                    mOutStr.Append(Field_Right(.Fields("intra_state_code").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("type_move").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("all_rail_code").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("move_via_water").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("transit_code").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("truck_for_rail").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("rebill").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("shortline_miles").Value, 4))
                    mOutStr.Append(Field_Right(.Fields("stratum").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("subsample").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("exp_factor").Value, 5))
                    mOutStr.Append(Field_Right(.Fields("exp_factor_th").Value, 3))
                    mOutStr.Append(Field_Right(.Fields("jf").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("o_bea").Value, 3))
                    mOutStr.Append(Field_Right(.Fields("o_ft").Value, 1))
                    mOutStr.Append(Field_Left(.Fields("jct1_st").Value, 2))
                    mOutStr.Append(Field_Left(.Fields("jct2_st").Value, 2))
                    mOutStr.Append(Field_Left(.Fields("jct3_st").Value, 2))
                    mOutStr.Append(Field_Left(.Fields("jct4_st").Value, 2))
                    mOutStr.Append(Field_Left(.Fields("jct5_st").Value, 2))
                    mOutStr.Append(Field_Left(.Fields("jct6_st").Value, 2))
                    mOutStr.Append(Field_Left(.Fields("jct7_st").Value, 2))
                    mOutStr.Append(Field_Left(.Fields("jct8_st").Value, 2))
                    mOutStr.Append(Field_Left(.Fields("jct9_st").Value, 2))
                    mOutStr.Append(Field_Right(.Fields("t_bea").Value, 3))
                    mOutStr.Append(Field_Right(.Fields("t_ft").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("report_period").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("car_cap").Value, 5))
                    mOutStr.Append(Field_Right(.Fields("nom_car_cap").Value, 3))
                    mOutStr.Append(Field_Right(.Fields("tare").Value, 4))
                    mOutStr.Append(Field_Right(.Fields("outside_l").Value, 5))
                    mOutStr.Append(Field_Right(.Fields("outside_w").Value, 4))
                    mOutStr.Append(Field_Right(.Fields("outside_h").Value, 4))
                    mOutStr.Append(Field_Right(.Fields("ex_outside_h").Value, 4))
                    mOutStr.Append(Field_Left(.Fields("type_wheel").Value, 1))
                    mOutStr.Append(Field_Left(.Fields("no_axles").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("draft_gear").Value, 2))
                    mOutStr.Append(Field_Right(.Fields("art_units").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("err_code1").Value, 2))
                    mOutStr.Append(Field_Right(.Fields("err_code2").Value, 2))
                    mOutStr.Append(Space(46))
                    mOutStr.Append(Field_Left(.Fields("error_flg").Value, 1))
                    mOutStr.Append(Field_Right(.Fields("cars").Value, 6))
                    mOutStr.Append(Field_Right(.Fields("tons").Value, 9))
                    mOutStr.Append(Field_Right(.Fields("total_rev").Value, 11))
                    mOutStr.Append(Field_Right(.Fields("tc_units").Value, 6))

                End With

                outfile.WriteLine(mOutStr)

                rst.MoveNext()
                mThisRec = mThisRec + 1

            Loop

            'close the output file
            outfile.Close()

        End If

        rst.Close()
        rst = Nothing

        txt_StatusBox.Text = "Done!"
        Refresh()

SkipIt:

    End Sub

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the WBGenerators And Utilities Menu Form
        Dim frmNew As New frm_WBGeneratorsAndUtilitiesMenu()
        frmNew.Show()
        ' Close this Menu
        Close()
    End Sub
End Class