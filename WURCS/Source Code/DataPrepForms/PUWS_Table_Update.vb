Public Class PUWS_Table_Update

    Private Sub PUWS_Table_Update_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        mDataTable = Get_URCS_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            cmb_URCSYear.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
        Next

        mDataTable = Nothing

    End Sub

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the previous Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close this form
        Me.Close()
    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click
        Dim rst As New ADODB.Recordset
        Dim mStrSQL As String
        Dim InLine As String
        Dim mOverWrite As Integer
        Dim mMaxRecs As Long, mThisRec As Long
        Dim PUWS As Class_PUWS
        Dim mMasking_Factor As Single
        Dim mU_Masked_Rev As Integer, mTotal_Masked_Rev As Integer

        Me.CenterToScreen()
        mMaxRecs = 0
        mThisRec = 0

        ' Verify that the user has selected a year and a text file
        If Me.cmb_URCSYear.Text = "" Then
            MsgBox("You must select a year value to process!", vbOKOnly, "Error!")
            GoTo QuitIt
        End If

        If Me.txt_Input_FilePath.Text = "" Then
            MsgBox("You must select a file to process!", vbOKOnly, "Error!")
            GoTo QuitIt
        End If

        ' Get the table and database names for the Masked Waybills file
        Global_Variables.Gbl_Masked_TableName = Get_Table_Name_From_SQL(Me.cmb_URCSYear.Text, "MASKED")

        ' Get the table name for the Unmasked_Rev table. The database should remain the same
        Global_Variables.Gbl_Unmasked_Rev_TableName = Get_Table_Name_From_SQL(Me.cmb_URCSYear.Text, "UNMASKED_REV")

        ' Get the table name for the PUWS tables
        Global_Variables.Gbl_PUWS_Masked_Tablename = Get_Table_Name_From_SQL(Me.cmb_URCSYear.Text, "PUWS_MASKED")
        Global_Variables.Gbl_PUWS_Masking_Factors_Tablename = Get_Table_Name_From_SQL(Me.cmb_URCSYear.Text, "PUWS_MASKING_FACTORS")

        ' Open/Check
        OpenADOConnection(Global_Variables.Gbl_Waybill_Database_Name)

        If VerifyTableExist(Global_Variables.Gbl_Waybill_Database_Name, Global_Variables.Gbl_PUWS_Masked_Tablename) = False Then
            mOverWrite = vbYes
        Else
            mOverWrite = MsgBox("Found " & Count_Any_Table(Global_Variables.Gbl_Waybill_Database_Name, Global_Variables.Gbl_PUWS_Masked_Tablename).ToString & " records.  Overwrite?",
                            vbYesNo, "WARNING!")
        End If

        If mOverWrite = vbYes Then
            Me.txt_StatusBox.Text = "Wiping existing PUWS data..."
            Refresh()

            ' Faster to truncate the table than to drop/recreate
            If VerifyTableExist(Global_Variables.Gbl_Waybill_Database_Name, Global_Variables.Gbl_PUWS_Masked_Tablename) Then
                mStrSQL = "TRUNCATE TABLE " & Global_Variables.Gbl_PUWS_Masked_Tablename
                Global_Variables.gbl_ADOConnection.Execute(mStrSQL)
            Else
                ' Recreate the tables
                mStrSQL = "CREATE TABLE " & Global_Variables.Gbl_PUWS_Masked_Tablename &
                    " ( " &
                    "[WB_Date] [nvarchar](6) NULL," &
                    "[acct_period] [nvarchar](4) NULL," &
                    "[u_cars] [int] NULL," &
                    "[car_own] [nvarchar](1) NULL," &
                    "[car_typ] [nvarchar](4) NULL," &
                    "[mech] [nvarchar](4) NULL," &
                    "[stb_car_typ] [tinyint] NULL," &
                    "[tofc_serv_code] [nvarchar](3) NULL," &
                    "[u_tc_units] [int] NULL," &
                    "[tofc_own_code] [nvarchar](1) NULL," &
                    "[tofc_unit_type] [nvarchar](1) NULL," &
                    "[haz_bulk] [nvarchar](1) NULL," &
                    "[stcc] [nvarchar](7) NULL," &
                    "[bill_wght_tons] [int] NULL," &
                    "[act_wght] [int] NULL," &
                    "[u_rev] [int] NULL," &
                    "[tran_chrg] [int] NULL," &
                    "[misc_chrg] [int] NULL," &
                    "[intra_state_code] [int] NULL," &
                    "[type_move] [int] NULL," &
                    "[all_rail_code] [int] NULL," &
                    "[move_via_water] [int] NULL," &
                    "[transit_code] [int] NULL," &
                    "[truck_for_rail] [int] NULL," &
                    "[rebill] [int] NULL," &
                    "[shortline_miles] [int] NULL," &
                    "[stratum] [int] NULL," &
                    "[subsample] [int] NULL," &
                    "[exp_factor] [int] NULL," &
                    "[exp_factor_th] [int] NULL," &
                    "[jf] [int] NULL," &
                    "[o_bea] [smallint] NULL," &
                    "[o_ft] [smallint] NULL," &
                    "[jct1_st] [nvarchar](2) NULL," &
                    "[jct2_st] [nvarchar](2) NULL," &
                    "[jct3_st] [nvarchar](2) NULL," &
                    "[jct4_st] [nvarchar](2) NULL," &
                    "[jct5_st] [nvarchar](2) NULL," &
                    "[jct6_st] [nvarchar](2) NULL," &
                    "[jct7_st] [nvarchar](2) NULL," &
                    "[jct8_st] [nvarchar](2) NULL," &
                    "[jct9_st] [nvarchar](2) NULL," &
                    "[t_bea] [smallint] NULL," &
                    "[t_ft] [smallint] NULL," &
                    "[report_period] [tinyint] NULL," &
                    "[car_cap] [int] NULL," &
                    "[nom_car_cap] [smallint] NULL," &
                    "[tare] [smallint] NULL," &
                    "[outside_l] [int] NULL," &
                    "[outside_w] [smallint] NULL," &
                    "[outside_h] [smallint] NULL," &
                    "[ex_outside_h] [smallint] NULL," &
                    "[type_wheel] [nvarchar](1) NULL," &
                    "[no_axles] [nvarchar](1) NULL," &
                    "[draft_gear] [tinyint] NULL," &
                    "[art_units] [int] NULL," &
                    "[err_code1] [tinyint] NULL," &
                    "[err_code2] [tinyint] NULL," &
                    "[error_flg] [nvarchar](1) NULL," &
                    "[cars] [int] NULL," &
                    "[tons] [int] NULL," &
                    "[total_rev] [decimal](18, 0) NULL," &
                    "[tc_units] [int] NULL," &
                    "[serial_no] [varchar](6) NOT NULL) ON [PRIMARY]"

                Global_Variables.gbl_ADOConnection.Execute(mStrSQL)

                ' Now we need to create the index for the new table
                mStrSQL = "ALTER TABLE " & Global_Variables.Gbl_PUWS_Masked_Tablename &
                    " ADD CONSTRAINT pk_" & Global_Variables.Gbl_PUWS_Masked_Tablename &
                    " PRIMARY KEY CLUSTERED " &
                    " (Serial_No) WITH (STATISTICS_NORECOMPUTE = OFF," &
                    " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, " &
                    " ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"

                Global_Variables.gbl_ADOConnection.Execute(mStrSQL)

            End If

            If VerifyTableExist(Global_Variables.Gbl_Waybill_Database_Name, Global_Variables.Gbl_PUWS_Masking_Factors_Tablename) Then
                mStrSQL = "TRUNCATE TABLE " & Global_Variables.Gbl_PUWS_Masking_Factors_Tablename
                Global_Variables.gbl_ADOConnection.Execute(mStrSQL)
            Else
                ' We'll create the masking factor table now
                mStrSQL = "CREATE TABLE " & Global_Variables.Gbl_PUWS_Masking_Factors_Tablename &
                    " ( " &
                    "Serial_No nvarchar(6) NOT NULL, " &
                    "Masking_Factor [decimal](3, 2) NOT NULL) ON [PRIMARY]"

                Global_Variables.gbl_ADOConnection.Execute(mStrSQL)

                ' Now we need to create the index for the new table
                mStrSQL = "ALTER TABLE " & Global_Variables.Gbl_PUWS_Masking_Factors_Tablename &
                    " ADD CONSTRAINT pk_" & Global_Variables.Gbl_PUWS_Masking_Factors_Tablename &
                    " PRIMARY KEY CLUSTERED " &
                    " (Serial_No) WITH (STATISTICS_NORECOMPUTE = OFF," &
                    " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, " &
                    " ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"

                Global_Variables.gbl_ADOConnection.Execute(mStrSQL)

                ' Finally, we need to set/restrict the rights for the masking factors table
                mStrSQL = "GRANT SELECT ON " & Global_Variables.Gbl_PUWS_Masking_Factors_Tablename &
                    " TO Unmasked_Revenue_Read_Access"

                Global_Variables.gbl_ADOConnection.Execute(mStrSQL)

                mStrSQL = "DENY SELECT ON " & Global_Variables.Gbl_PUWS_Masking_Factors_Tablename &
                " TO Unmasked_Revenue_Deny_Access"

                Global_Variables.gbl_ADOConnection.Execute(mStrSQL)

            End If
        Else
            ' The user chose not to overwrite.  Abort the operation.
            GoTo EndIt
        End If

        ' Load the text file to the Masked_PUWS table
        'Open the text file and check to make sure that it is the correct length
        Me.txt_StatusBox.Text = "Examining the input file..."
        Refresh()

        mMaxRecs = 0

        Using reader As StreamReader = New StreamReader(Me.txt_Input_FilePath.Text)
            ' Read one line from file
            InLine = reader.ReadLine

            If Len(InLine) <> 253 Then
                MsgBox("The Text file appears to be invalid - Line Length = " & Len(InLine).ToString, vbOKOnly, "Error!")
                GoTo QuitIt
            End If

            ' Count the number of records
            mMaxRecs = mMaxRecs + 1

            Do While Not reader.EndOfStream = True
                reader.ReadLine()
                mMaxRecs = mMaxRecs + 1
            Loop

            reader.Dispose()

        End Using

        'reopen it to resposition the pointer
        'Read to the end of the stream, adding data to the sql table

        mThisRec = 0
        Using reader As StreamReader = New StreamReader(Me.txt_Input_FilePath.Text)

            Do While Not reader.EndOfStream = True
                InLine = reader.ReadLine()
                mThisRec = mThisRec + 1

                If mThisRec Mod 100 = 0 Then
                    Me.txt_StatusBox.Text = "Processing record " & mThisRec.ToString & " of " & mMaxRecs.ToString & " records..."
                    Refresh()
                    Application.DoEvents()
                End If

                PUWS = New Class_PUWS
                With PUWS
                    ' Note that the Substring option is zero-based
                    .WB_Date = InLine.Substring(0, 6)
                    .Acct_Period = InLine.Substring(6, 4)
                    .U_Cars = ReturnInteger(InLine.Substring(10, 4))
                    .Car_Own = InLine.Substring(14, 1)
                    .Car_Typ = InLine.Substring(15, 4)
                    .Mech = InLine.Substring(19, 4)
                    .STB_Car_Type = ReturnInteger(InLine.Substring(23, 2))
                    .TOFC_Serv_Code = InLine.Substring(25, 3)
                    .U_TC_Units = ReturnInteger(InLine.Substring(28, 4))
                    .TOFC_Own_Code = InLine.Substring(32, 1)
                    .TOFC_Unit_Type = InLine.Substring(33, 1)
                    .Haz_Bulk = InLine.Substring(34, 1)
                    .STCC = InLine.Substring(35, 5)
                    .Bill_Wght_Tons = ReturnInteger(InLine.Substring(40, 7))
                    .Act_Wght = ReturnInteger(InLine.Substring(47, 7))
                    .U_Rev = ReturnInteger(InLine.Substring(54, 9))
                    .Tran_Chrg = ReturnInteger(InLine.Substring(63, 9))
                    .Misc_Chrg = ReturnInteger(InLine.Substring(72, 9))
                    .Intra_State_Code = ReturnInteger(InLine.Substring(81, 1))
                    .Type_Move = ReturnInteger(InLine.Substring(82, 1))
                    .All_Rail_Code = ReturnInteger(InLine.Substring(83, 1))
                    .Move_Via_Water = ReturnInteger(InLine.Substring(84, 1))
                    .Transit_Code = ReturnInteger(InLine.Substring(85, 1))
                    .Truck_For_Rail = ReturnInteger(InLine.Substring(86, 1))
                    .Rebill = ReturnInteger(InLine.Substring(87, 1))
                    .Shortline_Miles = ReturnInteger(InLine.Substring(88, 4))
                    .Stratum = ReturnInteger(InLine.Substring(92, 1))
                    .Subsample = ReturnInteger(InLine.Substring(93, 1))
                    .Exp_Factor = ReturnInteger(InLine.Substring(94, 5))
                    .Exp_Factor_Th = ReturnInteger(InLine.Substring(99, 3))
                    .JF = ReturnInteger(InLine.Substring(102, 1))
                    .O_BEA = ReturnInteger(InLine.Substring(103, 3))
                    .O_FT = ReturnInteger(InLine.Substring(106, 1))
                    .JCT1_ST = InLine.Substring(107, 2)
                    .JCT2_ST = InLine.Substring(109, 2)
                    .JCT3_ST = InLine.Substring(111, 2)
                    .JCT4_ST = InLine.Substring(113, 2)
                    .JCT5_ST = InLine.Substring(115, 2)
                    .JCT6_ST = InLine.Substring(117, 2)
                    .JCT7_ST = InLine.Substring(119, 2)
                    .JCT8_ST = InLine.Substring(121, 2)
                    .JCT9_ST = InLine.Substring(123, 2)
                    .T_BEA = ReturnInteger(InLine.Substring(125, 3))
                    .T_FT = ReturnInteger(InLine.Substring(128, 1))
                    .Report_Period = ReturnInteger(InLine.Substring(129, 1))
                    .Car_Cap = ReturnInteger(InLine.Substring(130, 5))
                    .Nom_Car_Cap = ReturnInteger(InLine.Substring(135, 3))
                    .Tare = ReturnInteger(InLine.Substring(138, 4))
                    .Outside_L = ReturnInteger(InLine.Substring(142, 5))
                    .Outside_W = ReturnInteger(InLine.Substring(147, 4))
                    .Outside_H = ReturnInteger(InLine.Substring(151, 4))
                    .Ex_Outside_H = ReturnInteger(InLine.Substring(155, 4))
                    .Type_Wheel = InLine.Substring(159, 1)
                    .No_Axles = InLine.Substring(160, 1)
                    .Draft_Gear = ReturnInteger(InLine.Substring(161, 2))
                    .Art_Units = ReturnInteger(InLine.Substring(163, 1))
                    .Err_Code1 = ReturnInteger(InLine.Substring(164, 2))
                    .Err_Code2 = ReturnInteger(InLine.Substring(166, 2))
                    'Cols 169-214 are blank
                    .Error_Flg = InLine.Substring(214, 1)
                    .Cars = ReturnInteger(InLine.Substring(215, 6))
                    .Tons = ReturnInteger(InLine.Substring(221, 9))
                    .Total_Rev = ReturnInteger(InLine.Substring(230, 11))
                    .TC_Units = ReturnInteger(InLine.Substring(241, 6))
                    .Serial_No = InLine.Substring(247, 6)

                    'Get the revenues from SQL server for this serial number for anything other than UP & PAL
                    mStrSQL = "SELECT serial_no, total_rev, u_rev, report_rr, rate_flg, exp_factor_th, u_rev_unmasked " &
                        " FROM " & Global_Variables.Gbl_Masked_TableName & " INNER JOIN " &
                        Global_Variables.Gbl_Unmasked_Rev_TableName & " ON " &
                        Global_Variables.Gbl_Masked_TableName & ".Serial_No = " &
                        Global_Variables.Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_No " &
                        "WHERE serial_no = '" & .Serial_No & "'"

                    ' Check/Open the SQL connection
                    OpenADOConnection(Global_Variables.Gbl_Waybill_Database_Name)

                    rst = SetRST()
                    rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)

                    ' Set the default values for the revenue fields to memvars
                    mTotal_Masked_Rev = rst.Fields("total_rev").Value
                    mU_Masked_Rev = rst.Fields("u_rev").Value
                    mMasking_Factor = 1

                    If rst.Fields("rate_flg").Value > 0 Then
                        'We need to mask the unmasked revenues in the new table for contract moves only

                        ' Set up the Random Number Generator
                        Randomize()

                        Select Case rst.Fields("report_rr").Value

                            ' NS - 555
                            Case 555
                                ' Not used at this time

                                ' CSXT - 712
                            Case 712
                                ' Not used at this time

                                ' BNSF - 777
                            Case 777
                                ' Not used at this time

                                ' CP - 105
                            Case 105
                                ' Not used at this time

                            Case 802, 907
                                ' UP & P&L (P&L added for 2013+)

                                Do Until mMasking_Factor > 1.04 And mMasking_Factor <= 1.55
                                    mMasking_Factor = 1 + System.Math.Round(Rnd(), 2, MidpointRounding.AwayFromZero)
                                    'mMasking_Factor = System.Math.Round(mMasking_Factor, 2, MidpointRounding.AwayFromZero)
                                Loop

                                ' For these roads we mask the unmasked revenues
                                If mU_Masked_Rev > 0 And mTotal_Masked_Rev > 0 Then
                                    mU_Masked_Rev = System.Math.Round(rst.Fields("u_rev_unmasked").Value * mMasking_Factor, 0, MidpointRounding.AwayFromZero)
                                    mTotal_Masked_Rev = mU_Masked_Rev * rst.Fields("exp_factor_th").Value
                                End If

                            Case Else

                                ' Not used at this time

                        End Select

                    End If

                    ' Now we need to record/write the masking factor to its table
                    mStrSQL = "INSERT INTO " & Global_Variables.Gbl_PUWS_Masking_Factors_Tablename & " (" &
                            "serial_no, masking_factor) VALUES ("
                    mStrSQL = mStrSQL & "'" & rst.Fields("Serial_No").Value & "', "
                    mStrSQL = mStrSQL & Format(mMasking_Factor, "0.00") & ")"

                    ' Check/Open the SQL connection
                    OpenADOConnection(Global_Variables.Gbl_Waybill_Database_Name)
                    'execute the SQL statement
                    Global_Variables.gbl_ADOConnection.Execute(mStrSQL)

                    ' Close the rst - we're done with it
                    rst.Close()
                    rst = Nothing

                    ' Start building the SQL statement to load the entire record into the table.
                    mStrSQL = "INSERT INTO " & Global_Variables.Gbl_PUWS_Masked_Tablename & " ("
                    mStrSQL = mStrSQL & "WB_Date, "
                    mStrSQL = mStrSQL & "Acct_Period, "
                    mStrSQL = mStrSQL & "U_Cars, "
                    mStrSQL = mStrSQL & "Car_Own, "
                    mStrSQL = mStrSQL & "Car_Typ, "
                    mStrSQL = mStrSQL & "Mech, "
                    mStrSQL = mStrSQL & "STB_Car_Typ, "
                    mStrSQL = mStrSQL & "TOFC_Serv_Code, "
                    mStrSQL = mStrSQL & "U_TC_Units, "
                    mStrSQL = mStrSQL & "TOFC_Own_Code, "
                    mStrSQL = mStrSQL & "TOFC_Unit_Type, "
                    mStrSQL = mStrSQL & "Haz_Bulk, "
                    mStrSQL = mStrSQL & "STCC, "
                    mStrSQL = mStrSQL & "Bill_Wght_Tons, "
                    mStrSQL = mStrSQL & "Act_Wght, "
                    mStrSQL = mStrSQL & "U_Rev, "
                    mStrSQL = mStrSQL & "Tran_Chrg, "
                    mStrSQL = mStrSQL & "Misc_Chrg, "
                    mStrSQL = mStrSQL & "Intra_State_Code, "
                    mStrSQL = mStrSQL & "Type_Move, "
                    mStrSQL = mStrSQL & "All_Rail_Code, "
                    mStrSQL = mStrSQL & "Move_Via_Water, "
                    mStrSQL = mStrSQL & "Transit_Code, "
                    mStrSQL = mStrSQL & "Truck_For_Rail, "
                    mStrSQL = mStrSQL & "Rebill, "
                    mStrSQL = mStrSQL & "Shortline_Miles, "
                    mStrSQL = mStrSQL & "Stratum, "
                    mStrSQL = mStrSQL & "Subsample, "
                    mStrSQL = mStrSQL & "Exp_Factor, "
                    mStrSQL = mStrSQL & "Exp_Factor_Th, "
                    mStrSQL = mStrSQL & "JF, "
                    mStrSQL = mStrSQL & "O_BEA, "
                    mStrSQL = mStrSQL & "O_FT, "
                    mStrSQL = mStrSQL & "JCT1_ST, "
                    mStrSQL = mStrSQL & "JCT2_ST, "
                    mStrSQL = mStrSQL & "JCT3_ST, "
                    mStrSQL = mStrSQL & "JCT4_ST, "
                    mStrSQL = mStrSQL & "JCT5_ST, "
                    mStrSQL = mStrSQL & "JCT6_ST, "
                    mStrSQL = mStrSQL & "JCT7_ST, "
                    mStrSQL = mStrSQL & "JCT8_ST, "
                    mStrSQL = mStrSQL & "JCT9_ST, "
                    mStrSQL = mStrSQL & "T_BEA, "
                    mStrSQL = mStrSQL & "T_FT, "
                    mStrSQL = mStrSQL & "Report_Period, "
                    mStrSQL = mStrSQL & "Car_Cap, "
                    mStrSQL = mStrSQL & "Nom_Car_Cap, "
                    mStrSQL = mStrSQL & "Tare, "
                    mStrSQL = mStrSQL & "Outside_L, "
                    mStrSQL = mStrSQL & "Outside_W, "
                    mStrSQL = mStrSQL & "Outside_H, "
                    mStrSQL = mStrSQL & "Ex_Outside_H, "
                    mStrSQL = mStrSQL & "Type_Wheel, "
                    mStrSQL = mStrSQL & "No_Axles, "
                    mStrSQL = mStrSQL & "Draft_Gear, "
                    mStrSQL = mStrSQL & "Art_Units, "
                    mStrSQL = mStrSQL & "Err_Code1, "
                    mStrSQL = mStrSQL & "Err_Code2, "
                    mStrSQL = mStrSQL & "Error_Flg, "
                    mStrSQL = mStrSQL & "Cars, "
                    mStrSQL = mStrSQL & "Tons, "
                    mStrSQL = mStrSQL & "Total_Rev, "
                    mStrSQL = mStrSQL & "TC_Units, "
                    mStrSQL = mStrSQL & "Serial_No"

                    mStrSQL = mStrSQL & ") VALUES ("

                    mStrSQL = mStrSQL & "'" & .WB_Date.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .Acct_Period.ToString & "',"
                    mStrSQL = mStrSQL & .U_Cars.ToString & ","
                    mStrSQL = mStrSQL & "'" & .Car_Own.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .Car_Typ.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .Mech.ToString & "',"
                    mStrSQL = mStrSQL & .STB_Car_Type.ToString & ","
                    mStrSQL = mStrSQL & "'" & .TOFC_Serv_Code.ToString & "',"
                    mStrSQL = mStrSQL & .U_TC_Units.ToString & ","
                    mStrSQL = mStrSQL & "'" & .TOFC_Own_Code.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .TOFC_Unit_Type.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .Haz_Bulk.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .STCC.ToString & "',"
                    mStrSQL = mStrSQL & .Bill_Wght_Tons.ToString & ","
                    mStrSQL = mStrSQL & .Act_Wght.ToString & ","
                    mStrSQL = mStrSQL & mU_Masked_Rev.ToString & ","
                    mStrSQL = mStrSQL & .Tran_Chrg.ToString & ","
                    mStrSQL = mStrSQL & .Misc_Chrg.ToString & ","
                    mStrSQL = mStrSQL & .Intra_State_Code.ToString & ","
                    mStrSQL = mStrSQL & .Type_Move.ToString & ","
                    mStrSQL = mStrSQL & .All_Rail_Code.ToString & ","
                    mStrSQL = mStrSQL & .Move_Via_Water.ToString & ","
                    mStrSQL = mStrSQL & .Transit_Code.ToString & ","
                    mStrSQL = mStrSQL & .Truck_For_Rail.ToString & ","
                    mStrSQL = mStrSQL & .Rebill.ToString & ","
                    mStrSQL = mStrSQL & .Shortline_Miles.ToString & ","
                    mStrSQL = mStrSQL & .Stratum.ToString & ","
                    mStrSQL = mStrSQL & .Subsample.ToString & ","
                    mStrSQL = mStrSQL & .Exp_Factor.ToString & ","
                    mStrSQL = mStrSQL & .Exp_Factor_Th.ToString & ","
                    mStrSQL = mStrSQL & .JF.ToString & ","
                    mStrSQL = mStrSQL & .O_BEA.ToString & ","
                    mStrSQL = mStrSQL & .O_FT.ToString & ","
                    mStrSQL = mStrSQL & "'" & .JCT1_ST.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .JCT2_ST.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .JCT3_ST.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .JCT4_ST.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .JCT5_ST.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .JCT6_ST.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .JCT7_ST.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .JCT8_ST.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .JCT9_ST.ToString & "',"
                    mStrSQL = mStrSQL & .T_BEA.ToString & ","
                    mStrSQL = mStrSQL & .T_FT.ToString & ","
                    mStrSQL = mStrSQL & .Report_Period.ToString & ","
                    mStrSQL = mStrSQL & .Car_Cap.ToString & ","
                    mStrSQL = mStrSQL & .Nom_Car_Cap.ToString & ","
                    mStrSQL = mStrSQL & .Tare.ToString & ","
                    mStrSQL = mStrSQL & .Outside_L.ToString & ","
                    mStrSQL = mStrSQL & .Outside_W.ToString & ","
                    mStrSQL = mStrSQL & .Outside_H.ToString & ","
                    mStrSQL = mStrSQL & .Ex_Outside_H.ToString & ","
                    mStrSQL = mStrSQL & "'" & .Type_Wheel.ToString & "',"
                    mStrSQL = mStrSQL & "'" & .No_Axles.ToString & "',"
                    mStrSQL = mStrSQL & .Draft_Gear.ToString & ","
                    mStrSQL = mStrSQL & .Art_Units.ToString & ","
                    mStrSQL = mStrSQL & .Err_Code1.ToString & ","
                    mStrSQL = mStrSQL & .Err_Code2.ToString & ","
                    'Cols 169-214 are blank
                    mStrSQL = mStrSQL & "'" & .Error_Flg.ToString & "',"
                    mStrSQL = mStrSQL & .Cars.ToString & ","
                    mStrSQL = mStrSQL & .Tons.ToString & ","
                    mStrSQL = mStrSQL & mTotal_Masked_Rev.ToString & ","
                    mStrSQL = mStrSQL & .TC_Units.ToString & ","
                    mStrSQL = mStrSQL & "'" & .Serial_No.ToString & "'"

                    mStrSQL = mStrSQL & ")"

                    ' Check/Open the SQL connection
                    OpenADOConnection(Global_Variables.Gbl_Waybill_Database_Name)

                    ' Execute the command
                    Global_Variables.gbl_ADOConnection.Execute(mStrSQL)

                End With

            Loop

            reader.Dispose()

        End Using

        Me.txt_StatusBox.Text = "Done!"
        Refresh()

EndIt:
QuitIt:

    End Sub

    Private Sub btn_Input_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Text Files|*.txt|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_FilePath.Text = fd.FileName
        End If
    End Sub
End Class