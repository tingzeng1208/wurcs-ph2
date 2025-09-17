Imports System.Data.SqlClient
Imports System.Text
Public Class Segments_Table_Builder

    Private Sub Segments_Table_Builder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        CenterToScreen()

        ' Load the Year combobox from the SQL database
        mDataTable = Get_URCS_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            cmb_URCS_Year.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
        Next

        mDataTable = Nothing

    End Sub

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close this Form
        Close()
    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        Dim mStrSQL As StringBuilder, mStrOutSQL As StringBuilder
        Dim mBolWrite As Boolean, mIntermodal As Boolean
        Dim mSglMaxRecs As Single, mThisRec As Single
        Dim mShipTypes(8) As String
        Dim mLooper As Integer, mLooper2 As Integer
        Dim mVC_Value As Integer
        Dim mSQLCmd As New SqlCommand
        Dim mDatatable As New DataTable

        Gbl_Masked_TableName = Get_Table_Name_From_SQL(cmb_URCS_Year.Text, "MASKED")
        Gbl_Segments_TableName = Get_Table_Name_From_SQL(cmb_URCS_Year.Text, "SEGMENTS")
        Gbl_Unmasked_Segments_TableName = Get_Table_Name_From_SQL(cmb_URCS_Year.Text, "UNMASKED_SEGMENTS")
        Gbl_Unmasked_Rev_TableName = Get_Table_Name_From_SQL(cmb_URCS_Year.Text, "UNMASKED_REV")

        Insert_AuditTrail_Record("URCS_Controls",
                                 "Updated WB" & cmb_URCS_Year.Text & "_Segments and WB" & cmb_URCS_Year.Text & "_Unmasked_Segments Tables.")

        Gbl_Waybill_Database_Name = Get_Database_Name_From_SQL(cmb_URCS_Year.Text, "MASKED")

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        'Initialize the return datatable and run the query to load it

        txt_StatusBox.Text = "Searching for records..."
        Refresh()

        mSglMaxRecs = Count_Any_Table(Gbl_Waybill_Database_Name, Gbl_Segments_TableName)

        txt_StatusBox.Text = ""
        Refresh()

        If mSglMaxRecs > 0 Then
            mBolWrite = MsgBox("Are you sure that you want to overwrite " & mSglMaxRecs.ToString & " Segment Records for " & cmb_URCS_Year.Text & "?", vbYesNo, "Caution!")
        Else
            mBolWrite = MsgBox("Are you sure that you want to create Segment Records for " & cmb_URCS_Year.Text & "?", vbYesNo, "Caution!")
        End If

        If mBolWrite = True Then
            Cursor.Current = Cursors.WaitCursor

            Insert_AuditTrail_Record("URCS_Controls",
                                     "Inserted Segments data into WB" & cmb_URCS_Year.Text & "_Segments and WB" & cmb_URCS_Year.Text & "_Unmasked_Segments.")

            txt_StatusBox.Text = "Wiping the output tables"
            Refresh()

            OpenSQLConnection(Gbl_Waybill_Database_Name)

            ' We will delete/drop the WBxxxx_Segments table here

            ' Check to see if the table exists before trying the TRUNCATE command
            If VerifyTableExist(Gbl_Waybill_Database_Name, Gbl_Segments_TableName) Then
                mStrOutSQL = New StringBuilder
                mStrOutSQL.Append("TRUNCATE TABLE " & Gbl_Segments_TableName)
                OpenSQLConnection(Gbl_Waybill_Database_Name)
                mSQLCmd = New SqlCommand
                mSQLCmd.CommandType = CommandType.Text
                mSQLCmd.CommandText = mStrOutSQL.ToString
                mSQLCmd.Connection = gbl_SQLConnection
                mSQLCmd.ExecuteNonQuery()
            Else
                mStrOutSQL = New StringBuilder
                mStrOutSQL.Append("CREATE TABLE ")
                mStrOutSQL.Append(Gbl_Segments_TableName)
                mStrOutSQL.Append("( ")
                mStrOutSQL.Append("Serial_No varchar(6) NOT NULL, ")
                mStrOutSQL.Append("Seg_no tinyint NOT NULL, ")
                mStrOutSQL.Append("Total_Segs tinyint NULL, ")
                mStrOutSQL.Append("RR_Num smallint NULL, ")
                mStrOutSQL.Append("RR_Alpha nvarchar(5) NULL, ")
                mStrOutSQL.Append("RR_Dist int NULL, ")
                mStrOutSQL.Append("RR_Cntry nvarchar(2) NULL, ")
                mStrOutSQL.Append("RR_Rev decimal(18, 0) NULL, ")
                mStrOutSQL.Append("RR_VC int NULL, ")
                mStrOutSQL.Append("Seg_Type nvarchar(2) NULL, ")
                mStrOutSQL.Append("From_Node int NULL, ")
                mStrOutSQL.Append("To_Node int NULL, ")
                mStrOutSQL.Append("From_Loc nvarchar(9) NULL, ")
                mStrOutSQL.Append("From_St nvarchar(5) NULL, ")
                mStrOutSQL.Append("To_Loc nvarchar(9) NULL, ")
                mStrOutSQL.Append("To_St nvarchar(5) NULL) on [PRIMARY]")
                OpenSQLConnection(Gbl_Waybill_Database_Name)

                mSQLCmd = New SqlCommand
                mSQLCmd.CommandType = CommandType.Text
                mSQLCmd.CommandText = mStrOutSQL.ToString
                mSQLCmd.Connection = gbl_SQLConnection
                mSQLCmd.ExecuteNonQuery()

                ' Now we need to create the index for the table we just created.
                mStrOutSQL = New StringBuilder
                mStrOutSQL.Append("ALTER TABLE ")
                mStrOutSQL.Append(Gbl_Segments_TableName)
                mStrOutSQL.Append(" ADD CONSTRAINT pk_")
                mStrOutSQL.Append(Gbl_Segments_TableName)
                mStrOutSQL.Append(" PRIMARY KEY CLUSTERED ")
                mStrOutSQL.Append("(Serial_No, Seg_No) WITH (STATISTICS_NORECOMPUTE= OFF, ")
                mStrOutSQL.Append("IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ")
                mStrOutSQL.Append("ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]")
                mSQLCmd = New SqlCommand
                mSQLCmd.CommandType = CommandType.Text
                mSQLCmd.CommandText = mStrOutSQL.ToString
                mSQLCmd.Connection = gbl_SQLConnection
                mSQLCmd.ExecuteNonQuery()

            End If

            ' And now do the same for the WBxxxx_Unmasked_Segments table as well
            If VerifyTableExist(Gbl_Waybill_Database_Name, Gbl_Unmasked_Segments_TableName) Then
                mStrOutSQL = New StringBuilder
                mStrOutSQL.Append("TRUNCATE TABLE ")
                mStrOutSQL.Append(Gbl_Unmasked_Segments_TableName)
                mSQLCmd = New SqlCommand
                mSQLCmd.CommandType = CommandType.Text
                mSQLCmd.CommandText = mStrOutSQL.ToString
                mSQLCmd.Connection = gbl_SQLConnection
                mSQLCmd.ExecuteNonQuery()
            Else

                ' And create the Unmasked table for each segment
                mStrOutSQL = New StringBuilder
                mStrOutSQL.Append("CREATE TABLE ")
                mStrOutSQL.Append(Gbl_Unmasked_Segments_TableName)
                mStrOutSQL.Append(" ( ")
                mStrOutSQL.Append("Serial_No varchar(6) NOT NULL, ")
                mStrOutSQL.Append("Seg_no tinyint NOT NULL, ")
                mStrOutSQL.Append("RR_Unmasked_Rev decimal(18, 0) NULL) on [PRIMARY]")
                OpenSQLConnection(Gbl_Waybill_Database_Name)
                mSQLCmd = New SqlCommand
                mSQLCmd.CommandType = CommandType.Text
                mSQLCmd.CommandText = mStrOutSQL.ToString
                mSQLCmd.Connection = gbl_SQLConnection
                mSQLCmd.ExecuteNonQuery()

                ' And finally, create the index and rights for the Unmasked Segments table
                mStrOutSQL = New StringBuilder
                mStrOutSQL.Append("ALTER TABLE ")
                mStrOutSQL.Append(Gbl_Unmasked_Segments_TableName)
                mStrOutSQL.Append(" ADD CONSTRAINT pk_" & Gbl_Unmasked_Segments_TableName)
                mStrOutSQL.Append(" PRIMARY KEY CLUSTERED ")
                mStrOutSQL.Append("(Serial_No, Seg_No) WITH (STATISTICS_NORECOMPUTE= OFF, ")
                mStrOutSQL.Append("IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ")
                mStrOutSQL.Append("ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]")
                OpenSQLConnection(Gbl_Waybill_Database_Name)
                mSQLCmd = New SqlCommand
                mSQLCmd.CommandType = CommandType.Text
                mSQLCmd.CommandText = mStrOutSQL.ToString
                mSQLCmd.Connection = gbl_SQLConnection
                mSQLCmd.ExecuteNonQuery()

                mStrOutSQL = New StringBuilder
                mStrOutSQL.Append("GRANT SELECT ON " & Gbl_Unmasked_Segments_TableName)
                mStrOutSQL.Append(" TO Unmasked_Revenue_Read_Access")
                OpenSQLConnection(Gbl_Waybill_Database_Name)
                mSQLCmd = New SqlCommand
                mSQLCmd.CommandType = CommandType.Text
                mSQLCmd.CommandText = mStrOutSQL.ToString
                mSQLCmd.Connection = gbl_SQLConnection
                mSQLCmd.ExecuteNonQuery()

                mStrOutSQL = New StringBuilder
                mStrOutSQL.Append("DENY SELECT ON " & Gbl_Unmasked_Segments_TableName)
                mStrOutSQL.Append(" TO Unmasked_Revenue_Deny_Access")
                OpenSQLConnection(Gbl_Waybill_Database_Name)
                mSQLCmd = New SqlCommand
                mSQLCmd.CommandType = CommandType.Text
                mSQLCmd.CommandText = mStrOutSQL.ToString
                mSQLCmd.Connection = gbl_SQLConnection
                mSQLCmd.ExecuteNonQuery()

            End If

            txt_StatusBox.Text = "Getting Waybill Data From Server"
            Refresh()

            mStrSQL = New StringBuilder
            mStrSQL.Append("SELECT " & Gbl_Masked_TableName)
            mStrSQL.Append(".Serial_no, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JF, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".ORR, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT1, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR1, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT2, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR2, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT3, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR3, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT4, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR4, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT5, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR5, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT6, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR6, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT7, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".TRR, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".O_FSAC, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".T_FSAC, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".ORR_ALPHA, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR1_ALPHA, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR2_ALPHA, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR3_ALPHA, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR4_ALPHA, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR5_ALPHA, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR6_ALPHA, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".TRR_ALPHA, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".ORR_REV, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR1_REV, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR2_REV, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR3_REV, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR4_REV, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR5_REV, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR6_REV, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".TRR_REV, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".ORR_DIST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR1_DIST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR2_DIST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR3_DIST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR4_DIST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR5_DIST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR6_DIST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".TRR_DIST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".TOTAL_DIST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".REBILL, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".O_ST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT1_ST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT2_ST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT3_ST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT4_ST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT5_ST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT6_ST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JCT7_ST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".T_ST, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".ONET, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".NET1, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".NET2, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".NET3, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".NET4, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".NET5, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".NET6, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".NET7, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".TNET, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".ORR_CNTRY, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR1_CNTRY, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR2_CNTRY, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR3_CNTRY, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR4_CNTRY, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR5_CNTRY, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".JRR6_CNTRY, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".TRR_CNTRY, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".RR1_VC, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".RR2_VC, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".RR3_VC, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".RR4_VC, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".RR5_VC, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".RR6_VC, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".RR7_VC, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".RR8_VC, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".STB_CAR_TYP, ")
            mStrSQL.Append(Gbl_Masked_TableName & ".TOFC_SERV_CODE, ")
            mStrSQL.Append(Gbl_Unmasked_Rev_TableName & ".ORR_Unmasked_Rev, ")
            mStrSQL.Append(Gbl_Unmasked_Rev_TableName & ".JRR1_Unmasked_Rev, ")
            mStrSQL.Append(Gbl_Unmasked_Rev_TableName & ".JRR2_Unmasked_Rev, ")
            mStrSQL.Append(Gbl_Unmasked_Rev_TableName & ".JRR3_Unmasked_Rev, ")
            mStrSQL.Append(Gbl_Unmasked_Rev_TableName & ".JRR4_Unmasked_Rev, ")
            mStrSQL.Append(Gbl_Unmasked_Rev_TableName & ".JRR5_Unmasked_Rev, ")
            mStrSQL.Append(Gbl_Unmasked_Rev_TableName & ".JRR6_Unmasked_Rev, ")
            mStrSQL.Append(Gbl_Unmasked_Rev_TableName & ".TRR_Unmasked_Rev ")

            ' Now we add the join properties between the 2 tables

            mStrSQL.Append("FROM " & Gbl_Masked_TableName & " INNER JOIN ")
            mStrSQL.Append(Gbl_Unmasked_Rev_TableName & " ON ")
            mStrSQL.Append(Gbl_Masked_TableName & ".Serial_No = ")
            mStrSQL.Append(Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_No")

            Using daAdapter As New SqlDataAdapter(mStrSQL.ToString, gbl_SQLConnection)
                daAdapter.Fill(mDatatable)
            End Using

            mThisRec = 1

            For mLooper = 0 To mDatatable.Rows.Count - 1

                If mThisRec Mod 100 = 0 Then
                    txt_StatusBox.Text = "Loading record " & mThisRec.ToString & " of " & mDatatable.Rows.Count.ToString & "..."
                    Refresh()
                    Application.DoEvents()
                End If

                '********************************************************************************
                ' Determine if the move is Intermodal
                '********************************************************************************
                mIntermodal = False
                Select Case mDatatable.Rows(mLooper)("stb_car_typ")
                    Case 46, 49, 52, 54
                        If Len(Trim(ReturnString(mDatatable.Rows(mLooper)("tofc_serv_code"), 0))) > 0 Then
                            mIntermodal = True
                        End If
                End Select

                ''********************************************************************************
                ' Now we set the shiptypes for the waybill record
                '********************************************************************************
                If mDatatable.Rows(mLooper)("jf") = 0 Then
                    If ((mDatatable.Rows(mLooper)("total_dist") / 10) < 8.5) And (mIntermodal = False) Then
                        mShipTypes(0) = "IA"
                    Else
                        Select Case mDatatable.Rows(mLooper)("rebill")
                            Case 0
                                mShipTypes(0) = "OT"
                            Case 1
                                mShipTypes(0) = "OD"
                            Case 2
                                mShipTypes(0) = "RD"
                            Case 3
                                mShipTypes(0) = "RT"
                        End Select
                    End If
                Else
                    If (mDatatable.Rows(mLooper)("total_dist") / 10 >= 8.5) Or (mIntermodal = True) Then
                        '********************************************************************************
                        ' Set the default shiptype array values
                        ' Array position 0 is the ORR
                        ' Positions 1-6 are the JRRs
                        ' Position 7 is the TRR
                        '********************************************************************************
                        mShipTypes(0) = "OD"
                        For mLooper2 = 1 To 6
                            mShipTypes(mLooper2) = "RD"
                        Next mLooper2
                        mShipTypes(7) = "RT"
                    Else
                        For mLooper2 = 0 To 7
                            mShipTypes(mLooper2) = "IR"
                        Next mLooper2
                    End If

                    Select Case mDatatable.Rows(mLooper)("rebill")
                        Case 1
                            mShipTypes(7) = "RD"
                        Case 2
                            mShipTypes(0) = "RD"
                            mShipTypes(7) = "RD"
                        Case 3
                            mShipTypes(0) = "RD"
                    End Select

                End If

                'If (rst.Fields("jf").Value = 1) And (rst.Fields("rebill").Value = 0) Then
                '    If ((rst.Fields("total_dist").Value / 10) < 8.5) And (mIntermodal = False) Then
                '        '********************************************************************************
                '        ' Only set this to "IR" if there are only two railroads reported in the movement,
                '        ' it is not a rebill, it travels less than 8.5 miles, and it is not intermodal.
                '        '********************************************************************************
                '        For mLooper = 0 To 7
                '            mShipTypes(mLooper) = "IR"
                '        Next mLooper
                '    End If
                'End If

                '********************************************************************************
                ' Make accodations for the VC if it is blank
                '********************************************************************************
                If IsDBNull(mDatatable.Rows(mLooper)("RR1_VC")) = True Then
                    mVC_Value = 0
                Else
                    mVC_Value = mDatatable.Rows(mLooper)("RR1_VC")
                End If

                '********************************************************************************
                ' Make accodations for the VC if it is blank
                '********************************************************************************
                If IsDBNull(mDatatable.Rows(mLooper)("RR1_VC")) = True Then
                    mVC_Value = 0
                Else
                    mVC_Value = mDatatable.Rows(mLooper)("RR1_VC")
                End If

                '********************************************************************************
                '* Write the unique records to the Segments table
                '********************************************************************************

                Select Case mDatatable.Rows(mLooper)("JF")
                    Case 0

                        ' Add the record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            1,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("ORR"),
                            ReturnString(mDatatable.Rows(mLooper)("ORR_Alpha"), 0),
                            mDatatable.Rows(mLooper)("ORR_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("ORR_Cntry"), 0),
                            mDatatable.Rows(mLooper)("ORR_Rev"),
                            mVC_Value,
                            mShipTypes(0),
                            mDatatable.Rows(mLooper)("ONET"),
                            mDatatable.Rows(mLooper)("TNET"),
                            Field_Right(mDatatable.Rows(mLooper)("ORR"), 3) & "-" & Field_Right(mDatatable.Rows(mLooper)("O_FSAC"), 5),
                            mDatatable.Rows(mLooper)("O_ST"),
                            Field_Right(mDatatable.Rows(mLooper)("TRR"), 3) & "-" & Field_Right(mDatatable.Rows(mLooper)("T_FSAC"), 5),
                            mDatatable.Rows(mLooper)("T_ST")).ToString)

                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            1,
                            mDatatable.Rows(mLooper)("ORR_Unmasked_Rev"))).ToString()

                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                    Case 1
                        ' Add the record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            1,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("ORR"),
                            mDatatable.Rows(mLooper)("ORR_Alpha"),
                            mDatatable.Rows(mLooper)("ORR_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("ORR_Cntry"), 0),
                            mDatatable.Rows(mLooper)("ORR_Rev"),
                            mVC_Value,
                            mShipTypes(0),
                            mDatatable.Rows(mLooper)("ONET"),
                            mDatatable.Rows(mLooper)("NET1"),
                            Field_Right(mDatatable.Rows(mLooper)("ORR"), 3) & "-" & Field_Right(mDatatable.Rows(mLooper)("O_FSAC"), 5),
                            mDatatable.Rows(mLooper)("O_ST"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT1"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT1_ST"), 0)))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            1,
                            mDatatable.Rows(mLooper)("ORR_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR2_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR2_VC")
                        End If

                        ' Add the 2nd record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            2,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("TRR"),
                            mDatatable.Rows(mLooper)("TRR_Alpha"),
                            mDatatable.Rows(mLooper)("TRR_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("TRR_Cntry"), 0),
                            mDatatable.Rows(mLooper)("TRR_Rev"),
                            mDatatable.Rows(mLooper)("RR2_VC"),
                            mShipTypes(7),
                            mDatatable.Rows(mLooper)("NET1"),
                            mDatatable.Rows(mLooper)("TNET"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT1"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT1_ST"), 0),
                            Field_Right(mDatatable.Rows(mLooper)("TRR"), 3) & "-" & Field_Right(mDatatable.Rows(mLooper)("T_FSAC"), 5),
                            mDatatable.Rows(mLooper)("T_ST")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 2nd record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            2,
                            mDatatable.Rows(mLooper)("TRR_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                    Case 2

                        ' The first Segment and Unmasked Segment records will be added later

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR2_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR2_VC")
                        End If

                        ' Add the 2nd record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            2,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("JRR1"),
                            mDatatable.Rows(mLooper)("JRR1_Alpha"),
                            mDatatable.Rows(mLooper)("JRR1_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("JRR1_Cntry"), 0),
                            mDatatable.Rows(mLooper)("JRR1_Rev"),
                            mVC_Value,
                            mShipTypes(1),
                            mDatatable.Rows(mLooper)("NET1"),
                            mDatatable.Rows(mLooper)("NET2"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT1"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT1_ST"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT2"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT2_ST"), 0)))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 2nd record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            2,
                            mDatatable.Rows(mLooper)("JRR1_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR3_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR3_VC")
                        End If

                        ' Add the 3rd record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            3,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("TRR"),
                            mDatatable.Rows(mLooper)("TRR_Alpha"),
                            mDatatable.Rows(mLooper)("TRR_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("TRR_Cntry"), 0),
                            mDatatable.Rows(mLooper)("TRR_Rev"),
                            mVC_Value,
                            mShipTypes(7),
                            mDatatable.Rows(mLooper)("NET2"),
                            mDatatable.Rows(mLooper)("TNET"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT2"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT2_ST"), 0),
                            Field_Right(mDatatable.Rows(mLooper)("TRR"), 3) & "-" & Field_Right(mDatatable.Rows(mLooper)("T_FSAC"), 5),
                            mDatatable.Rows(mLooper)("T_ST")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 3rd record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            3,
                            mDatatable.Rows(mLooper)("TRR_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                    Case 3

                        ' The first 2 Segment and Unmasked Segment records will be added later

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR3_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR3_VC")
                        End If

                        ' Add the 3rd record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            3,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("JRR2"),
                            mDatatable.Rows(mLooper)("JRR2_Alpha"),
                            mDatatable.Rows(mLooper)("JRR2_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("JRR2_Cntry"), 0),
                            mDatatable.Rows(mLooper)("JRR2_Rev"),
                            mDatatable.Rows(mLooper)("RR3_VC"),
                            mShipTypes(2),
                            mDatatable.Rows(mLooper)("NET2"),
                            mDatatable.Rows(mLooper)("NET3"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT2"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT2_ST"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT3"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT3_ST"), 0)))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 3rd record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            3,
                            mDatatable.Rows(mLooper)("JRR2_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR4_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR4_VC")
                        End If

                        ' Add the 4th record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            4,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("TRR"),
                            mDatatable.Rows(mLooper)("TRR_Alpha"),
                            mDatatable.Rows(mLooper)("TRR_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("TRR_Cntry"), 0),
                            mDatatable.Rows(mLooper)("TRR_Rev"),
                            mVC_Value,
                            mShipTypes(7),
                            mDatatable.Rows(mLooper)("NET3"),
                            mDatatable.Rows(mLooper)("TNET"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT3"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT3_ST"), 0),
                            Field_Right(mDatatable.Rows(mLooper)("TRR"), 3) & "-" & Field_Right(mDatatable.Rows(mLooper)("T_FSAC"), 5),
                            mDatatable.Rows(mLooper)("T_ST")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 4th record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            4,
                            mDatatable.Rows(mLooper)("TRR_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                    Case 4

                        ' The first 3 Segment and Unmasked Segment records will be added later

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR4_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR4_VC")
                        End If

                        ' Add the 4th record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            4,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("JRR3"),
                            mDatatable.Rows(mLooper)("JRR3_Alpha"),
                            mDatatable.Rows(mLooper)("JRR3_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("JRR3_Cntry"), 0),
                            mDatatable.Rows(mLooper)("JRR3_Rev"),
                            mVC_Value,
                            mShipTypes(3),
                            mDatatable.Rows(mLooper)("NET3"),
                            mDatatable.Rows(mLooper)("NET4"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT3"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT3_ST"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT4"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT4_ST"), 0)))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 4th record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            4,
                            mDatatable.Rows(mLooper)("JRR3_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR5_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR5_VC")
                        End If

                        ' Add the 5th record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            5,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("TRR"),
                            mDatatable.Rows(mLooper)("TRR_Alpha"),
                            mDatatable.Rows(mLooper)("TRR_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("TRR_Cntry"), 0),
                            mDatatable.Rows(mLooper)("TRR_Rev"),
                            mVC_Value,
                            mShipTypes(7),
                            mDatatable.Rows(mLooper)("NET4"),
                            mDatatable.Rows(mLooper)("TNET"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT4"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT4_ST"), 0),
                            Field_Right(mDatatable.Rows(mLooper)("TRR"), 3) & "-" & Field_Right(mDatatable.Rows(mLooper)("T_FSAC"), 5),
                            mDatatable.Rows(mLooper)("T_ST")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 4th record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            5,
                            mDatatable.Rows(mLooper)("TRR_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                    Case 5

                        ' The first 4 Segment and Unmasked Segment records will be added later

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR5_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR5_VC")
                        End If

                        ' Add the 5th record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                        Gbl_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        5,
                        mDatatable.Rows(mLooper)("JF") + 1,
                        mDatatable.Rows(mLooper)("JRR4"),
                        mDatatable.Rows(mLooper)("JRR4_Alpha"),
                        mDatatable.Rows(mLooper)("JRR4_Dist"),
                        ReturnString(mDatatable.Rows(mLooper)("JRR4_Cntry"), 0),
                        mDatatable.Rows(mLooper)("JRR4_Rev"),
                        mVC_Value,
                        mShipTypes(4),
                        mDatatable.Rows(mLooper)("NET4"),
                        mDatatable.Rows(mLooper)("NET5"),
                        ReturnString(mDatatable.Rows(mLooper)("JCT4"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT4_ST"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT5"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT5_ST"), 0)))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 5th record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            5,
                            mDatatable.Rows(mLooper)("JRR4_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR6_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR6_VC")
                        End If

                        ' Add the 6th record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            6,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("TRR"),
                            mDatatable.Rows(mLooper)("TRR_Alpha"),
                            mDatatable.Rows(mLooper)("TRR_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("TRR_Cntry"), 0),
                            mDatatable.Rows(mLooper)("TRR_Rev"),
                            mDatatable.Rows(mLooper)("RR6_VC"),
                            mShipTypes(7),
                            mDatatable.Rows(mLooper)("NET5"),
                            mDatatable.Rows(mLooper)("TNET"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT5"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT5_ST"), 0),
                            Field_Right(mDatatable.Rows(mLooper)("TRR"), 3) & "-" & Field_Right(mDatatable.Rows(mLooper)("T_FSAC"), 5),
                            mDatatable.Rows(mLooper)("T_ST")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 6th record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            6,
                            mDatatable.Rows(mLooper)("TRR_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                    Case 6

                        ' The first 5 Segment and Unmasked Segment records will be added later

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR6_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR6_VC")
                        End If

                        ' Add the 6th record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            6,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("JRR5"),
                            mDatatable.Rows(mLooper)("JRR5_Alpha"),
                            mDatatable.Rows(mLooper)("JRR5_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("JRR5_Cntry"), 0),
                            mDatatable.Rows(mLooper)("JRR5_Rev"),
                            mVC_Value,
                            mShipTypes(5),
                            mDatatable.Rows(mLooper)("NET5"),
                            mDatatable.Rows(mLooper)("NET6"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT5"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT5_ST"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT6"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT6_ST"), 0)))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 6th record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            6,
                            mDatatable.Rows(mLooper)("JRR5_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR7_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR7_VC")
                        End If

                        ' Add the 7th record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            7,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("TRR"),
                            mDatatable.Rows(mLooper)("TRR_Alpha"),
                            mDatatable.Rows(mLooper)("TRR_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("TRR_Cntry"), 0),
                            mDatatable.Rows(mLooper)("TRR_Rev"),
                            mVC_Value,
                            mShipTypes(7),
                            mDatatable.Rows(mLooper)("NET6"),
                            mDatatable.Rows(mLooper)("TNET"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT6"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT6_ST"), 0),
                            ReturnString(Field_Right(mDatatable.Rows(mLooper)("TRR"), 3) & "-" & Field_Right(mDatatable.Rows(mLooper)("T_FSAC"), 5), 0),
                            ReturnString(mDatatable.Rows(mLooper)("T_ST"), 0)))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 7th record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            7,
                            mDatatable.Rows(mLooper)("TRR_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                    Case 7

                        ' The first 6 Segment and Unmasked Segment records will be added later

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR7_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR7_VC")
                        End If

                        ' Add the 7th record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            7,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("JRR6"),
                            mDatatable.Rows(mLooper)("JRR6_Alpha"),
                            mDatatable.Rows(mLooper)("JRR6_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("JRR6_Cntry"), 0),
                            mDatatable.Rows(mLooper)("JRR6_Rev"),
                            mVC_Value,
                            mShipTypes(6),
                            mDatatable.Rows(mLooper)("NET6"),
                            mDatatable.Rows(mLooper)("NET7"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT6"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT6_ST"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT7"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT7_ST"), 0)))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 7th record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            7,
                            mDatatable.Rows(mLooper)("JRR6_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        '********************************************************************************
                        ' Make accodations for the VC if it is blank
                        '********************************************************************************
                        If IsDBNull(mDatatable.Rows(mLooper)("RR8_VC")) = True Then
                            mVC_Value = 0
                        Else
                            mVC_Value = mDatatable.Rows(mLooper)("RR8_VC")
                        End If

                        ' Add the 8th record to the Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Segment_SQL(
                            Gbl_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            8,
                            mDatatable.Rows(mLooper)("JF") + 1,
                            mDatatable.Rows(mLooper)("TRR"),
                            ReturnString(mDatatable.Rows(mLooper)("TRR_Alpha"), 0),
                            mDatatable.Rows(mLooper)("TRR_Dist"),
                            ReturnString(mDatatable.Rows(mLooper)("TRR_Cntry"), 0),
                            mDatatable.Rows(mLooper)("TRR_Rev"),
                            mVC_Value,
                            mShipTypes(7),
                            mDatatable.Rows(mLooper)("NET7"),
                            mDatatable.Rows(mLooper)("TNET"),
                            ReturnString(mDatatable.Rows(mLooper)("JCT7"), 0),
                            ReturnString(mDatatable.Rows(mLooper)("JCT7_ST"), 0),
                            Field_Right(mDatatable.Rows(mLooper)("TRR"), 3) & "-" & Field_Right(mDatatable.Rows(mLooper)("T_FSAC"), 5),
                            ReturnString(mDatatable.Rows(mLooper)("T_ST"), 0)))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                        'Add the 8th record to the Unmasked Segments table
                        mStrSQL = New StringBuilder
                        mStrSQL.Append(Build_Unmasked_Segment_SQL(
                            Gbl_Unmasked_Segments_TableName,
                            mDatatable.Rows(mLooper)("Serial_No"),
                            8,
                            mDatatable.Rows(mLooper)("TRR_Unmasked_Rev")))
                        OpenSQLConnection(Gbl_Waybill_Database_Name)
                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = mStrSQL.ToString
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.ExecuteNonQuery()

                End Select

                If mDatatable.Rows(mLooper)("JF") > 1 Then
                    ' Write the first Segment and Unmasked Segment records
                    ' Add the record to the Segments table

                    '********************************************************************************
                    ' Make accodations for the VC if it is blank
                    '********************************************************************************
                    If IsDBNull(mDatatable.Rows(mLooper)("RR1_VC")) = True Then
                        mVC_Value = 0
                    Else
                        mVC_Value = mDatatable.Rows(mLooper)("RR1_VC")
                    End If

                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Segment_SQL(
                        Gbl_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        1,
                        mDatatable.Rows(mLooper)("JF") + 1,
                        mDatatable.Rows(mLooper)("ORR"),
                        mDatatable.Rows(mLooper)("ORR_Alpha"),
                        mDatatable.Rows(mLooper)("ORR_Dist"),
                        ReturnString(mDatatable.Rows(mLooper)("ORR_Cntry"), 0),
                        mDatatable.Rows(mLooper)("ORR_Rev"),
                        mVC_Value,
                        mShipTypes(0),
                        mDatatable.Rows(mLooper)("ONET"),
                        mDatatable.Rows(mLooper)("NET1"),
                        Field_Right(mDatatable.Rows(mLooper)("ORR"), 3) & "-" & Field_Right(mDatatable.Rows(mLooper)("O_FSAC"), 5),
                        mDatatable.Rows(mLooper)("O_ST"),
                        ReturnString(mDatatable.Rows(mLooper)("JCT1"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT1_ST"), 0)))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                    'Add the record to the Unmasked Segments table
                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Unmasked_Segment_SQL(
                        Gbl_Unmasked_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        1,
                        mDatatable.Rows(mLooper)("ORR_Unmasked_Rev")))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                End If

                If mDatatable.Rows(mLooper)("JF") > 2 Then
                    ' Write the 2nd Segment and Unmasked Segment records

                    '********************************************************************************
                    ' Make accodations for the VC if it is blank
                    '********************************************************************************
                    If IsDBNull(mDatatable.Rows(mLooper)("RR2_VC")) = True Then
                        mVC_Value = 0
                    Else
                        mVC_Value = mDatatable.Rows(mLooper)("RR2_VC")
                    End If

                    ' Add the 2nd record to the Segments table
                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Segment_SQL(
                        Gbl_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        2,
                        mDatatable.Rows(mLooper)("JF") + 1,
                        mDatatable.Rows(mLooper)("JRR1"),
                        mDatatable.Rows(mLooper)("JRR1_Alpha"),
                        mDatatable.Rows(mLooper)("JRR1_Dist"),
                        ReturnString(mDatatable.Rows(mLooper)("JRR1_Cntry"), 0),
                        mDatatable.Rows(mLooper)("JRR1_Rev"),
                        mVC_Value,
                        mShipTypes(1),
                        mDatatable.Rows(mLooper)("NET1"),
                        mDatatable.Rows(mLooper)("NET2"),
                        ReturnString(mDatatable.Rows(mLooper)("JCT1"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT1_ST"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT2"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT2_ST"), 0)))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                    'Add the 2nd record to the Unmasked Segments table
                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Unmasked_Segment_SQL(
                        Gbl_Unmasked_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        2,
                        mDatatable.Rows(mLooper)("JRR1_Unmasked_Rev")))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                End If

                If mDatatable.Rows(mLooper)("JF") > 3 Then
                    ' Write the 3rd Segment and Unmasked Segment records

                    '********************************************************************************
                    ' Make accodations for the VC if it is blank
                    '********************************************************************************
                    If IsDBNull(mDatatable.Rows(mLooper)("RR3_VC")) = True Then
                        mVC_Value = 0
                    Else
                        mVC_Value = mDatatable.Rows(mLooper)("RR3_VC")
                    End If

                    ' Add the 3rd record to the Segments table
                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Segment_SQL(
                        Gbl_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        3,
                        mDatatable.Rows(mLooper)("JF") + 1,
                        mDatatable.Rows(mLooper)("JRR2"),
                        mDatatable.Rows(mLooper)("JRR2_Alpha"),
                        mDatatable.Rows(mLooper)("JRR2_Dist"),
                        ReturnString(mDatatable.Rows(mLooper)("JRR2_Cntry"), 0),
                        mDatatable.Rows(mLooper)("JRR2_Rev"),
                        mVC_Value,
                        mShipTypes(2),
                        mDatatable.Rows(mLooper)("NET2"),
                        mDatatable.Rows(mLooper)("NET3"),
                        ReturnString(mDatatable.Rows(mLooper)("JCT2"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT2_ST"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT3"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT3_ST"), 0)))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                    'Add the 3rd record to the Unmasked Segments table
                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Unmasked_Segment_SQL(
                        Gbl_Unmasked_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        3,
                        mDatatable.Rows(mLooper)("JRR2_Unmasked_Rev")))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                End If

                If mDatatable.Rows(mLooper)("JF") > 4 Then
                    ' Write the 4th Segment and Unmasked Segment records

                    '********************************************************************************
                    ' Make accodations for the VC if it is blank
                    '********************************************************************************
                    If IsDBNull(mDatatable.Rows(mLooper)("RR4_VC")) = True Then
                        mVC_Value = 0
                    Else
                        mVC_Value = mDatatable.Rows(mLooper)("RR4_VC")
                    End If

                    ' Add the 4th record to the Segments table
                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Segment_SQL(
                        Gbl_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        4,
                        mDatatable.Rows(mLooper)("JF") + 1,
                        mDatatable.Rows(mLooper)("JRR3"),
                        mDatatable.Rows(mLooper)("JRR3_Alpha"),
                        mDatatable.Rows(mLooper)("JRR3_Dist"),
                        ReturnString(mDatatable.Rows(mLooper)("JRR3_Cntry"), 0),
                        mDatatable.Rows(mLooper)("JRR3_Rev"),
                        mVC_Value,
                        mShipTypes(3),
                        mDatatable.Rows(mLooper)("NET3"),
                        mDatatable.Rows(mLooper)("NET4"),
                        ReturnString(mDatatable.Rows(mLooper)("JCT3"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT3_ST"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT4"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT4_ST"), 0)))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                    'Add the 4th record to the Unmasked Segments table
                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Unmasked_Segment_SQL(
                        Gbl_Unmasked_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        4,
                        mDatatable.Rows(mLooper)("JRR3_Unmasked_Rev")))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                End If

                If mDatatable.Rows(mLooper)("JF") > 5 Then
                    ' Write the 5th Segment and Unmasked Segment records

                    '********************************************************************************
                    ' Make accodations for the VC if it is blank
                    '********************************************************************************
                    If IsDBNull(mDatatable.Rows(mLooper)("RR5_VC")) = True Then
                        mVC_Value = 0
                    Else
                        mVC_Value = mDatatable.Rows(mLooper)("RR5_VC")
                    End If

                    ' Add the 5th record to the Segments table
                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Segment_SQL(
                        Gbl_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        5,
                        mDatatable.Rows(mLooper)("JF") + 1,
                        mDatatable.Rows(mLooper)("JRR4"),
                        mDatatable.Rows(mLooper)("JRR4_Alpha"),
                        mDatatable.Rows(mLooper)("JRR4_Dist"),
                        ReturnString(mDatatable.Rows(mLooper)("JRR4_Cntry"), 0),
                        mDatatable.Rows(mLooper)("JRR4_Rev"),
                        mVC_Value,
                        mShipTypes(4),
                        mDatatable.Rows(mLooper)("NET4"),
                        mDatatable.Rows(mLooper)("NET5"),
                        ReturnString(mDatatable.Rows(mLooper)("JCT4"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT4_ST"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT5"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT5_ST"), 0)))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                    'Add the 4th record to the Unmasked Segments table
                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Unmasked_Segment_SQL(
                        Gbl_Unmasked_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        5,
                        mDatatable.Rows(mLooper)("JRR4_Unmasked_Rev")))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                End If

                If mDatatable.Rows(mLooper)("JF") > 6 Then
                    ' Write the 6th Segment and Unmasked Segment records

                    '********************************************************************************
                    ' Make accodations for the VC if it is blank
                    '********************************************************************************
                    If IsDBNull(mDatatable.Rows(mLooper)("RR6_VC")) = True Then
                        mVC_Value = 0
                    Else
                        mVC_Value = mDatatable.Rows(mLooper)("RR6_VC")
                    End If

                    ' Add the 6th record to the Segments table
                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Segment_SQL(
                        Gbl_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        6,
                        mDatatable.Rows(mLooper)("JF") + 1,
                        mDatatable.Rows(mLooper)("JRR5"),
                        mDatatable.Rows(mLooper)("JRR5_Alpha"),
                        mDatatable.Rows(mLooper)("JRR5_Dist"),
                        ReturnString(mDatatable.Rows(mLooper)("JRR5_Cntry"), 0),
                        mDatatable.Rows(mLooper)("JRR5_Rev"),
                        mVC_Value,
                        mShipTypes(5),
                        mDatatable.Rows(mLooper)("NET5"),
                        mDatatable.Rows(mLooper)("NET6"),
                        ReturnString(mDatatable.Rows(mLooper)("JCT5"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT5_ST"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT6"), 0),
                        ReturnString(mDatatable.Rows(mLooper)("JCT6_ST"), 0)))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                    'Add the 6th record to the Unmasked Segments table
                    mStrSQL = New StringBuilder
                    mStrSQL.Append(Build_Unmasked_Segment_SQL(
                        Gbl_Unmasked_Segments_TableName,
                        mDatatable.Rows(mLooper)("Serial_No"),
                        6,
                        mDatatable.Rows(mLooper)("JRR5_Unmasked_Rev")))
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.CommandText = mStrSQL.ToString
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.ExecuteNonQuery()

                End If

                mThisRec = mThisRec + 1

            Next

            txt_StatusBox.Text = "Done!"
            Refresh()

        Else
            txt_StatusBox.Text = "Run Aborted."
            Refresh()
        End If

    End Sub

    Public Shared Function NotNull(Of T)(ByVal Value As T, ByVal DefaultValue As T) As T
        If Value Is Nothing OrElse IsDBNull(Value) Then
            Return DefaultValue
        Else
            Return Value
        End If
    End Function
End Class