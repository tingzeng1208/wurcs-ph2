Imports System.Data.SqlClient
Imports System.Text
Public Class Unmasked_Table_Update

    Public mTableName As String
    Public mMaskedDataOnly As Boolean
    Public mUnmaskedDataFlag As Boolean

    Private Sub Unmasked_Table_Update_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable
        Dim mIndex As Integer

        CenterToScreen()

        mIndex = 0

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
        ' Open the Main Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close this Menu
        Close()
    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click
        '***************************************************************************
        ' This program reads the WByear_Masked and then builds the SQL statement to
        ' update the UnMasked_Rev table data.
        '***************************************************************************

        Dim mStrSQL As Stringbuilder
        Dim mStrOutSQL As String, mSTCC As String
        Dim mSTCC_W49 As String, mCar_Init As String, mTState As String
        Dim mAcctMonth As Integer, mAcctYear As Integer
        Dim mDate As Date, mYear As Integer, mMaxRecs As Decimal
        Dim mSerial_No As String, mWbNum As Long, Rec As Long, mFSAC As Long
        Dim mUCarNum As Long, mTCNum As Long, mUCars As Long
        Dim mRate_Flg As Integer, mRptRR As Integer, morr As Integer, mjrr1 As Integer
        Dim mjrr2 As Integer, mjrr3 As Integer, mjrr4 As Integer, mjrr5 As Integer
        Dim mjrr6 As Integer, mtrr As Integer, mJF As Integer
        Dim mTotal_Rev As Object, mORR_Rev As Object, mJRR1_Rev As Object
        Dim mJRR2_Rev As Object, mJRR3_Rev As Object, mJRR4_Rev As Object
        Dim mJRR5_Rev As Object, mJRR6_Rev As Object, mTRR_Rev As Object
        Dim mU_Rev As Long, mExp_Factor_Th As Integer, mTotal_Unmask_Rev As Long
        Dim mWorkDate As Date
        Dim mOverWrite As Integer

        Dim mDataTable As DataTable
        Dim mSQLCmd As SqlCommand

        'Load the arrays for unmasking functions
        LoadArrayData()

        'Load the table information and find the masked table name
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(cmb_URCSYear.Text, "MASKED")
        OpenSQLConnection(Get_Database_Name_From_SQL(cmb_URCSYear.Text, "MASKED"))

        mDataTable = New DataTable

        txt_StatusBox.Text = "Searching for records..."
        Refresh()

        mStrSQL = New StringBuilder
        mStrSQL.Append("SELECT jf FROM dbo." & Gbl_Masked_TableName)

        Using daAdapter As New SqlDataAdapter(mStrSQL.ToString, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        txt_StatusBox.Text = ""
        Refresh()

        mMaxRecs = mDataTable.Rows.Count

        mOverWrite = MsgBox("The number of records returned is " & mMaxRecs.ToString &
                ". Are you sure you want to write the data?", vbYesNo)

        If mOverWrite = vbYes Then

            Insert_AuditTrail_Record("URCS_Controls",
                                     "Inserted Unmasked Waybill Data to WB" & cmb_URCSYear.Text & "_Unmasked_Rev.")

            txt_StatusBox.Text = "Wiping/Creating the output table"
            Refresh()

            Gbl_Unmasked_Rev_TableName = Get_Table_Name_From_SQL(cmb_URCSYear.Text, "UNMASKED_REV")
            mStrOutSQL = ""

            ' Determine if the table exists
            If VerifyTableExist(Gbl_Waybill_Database_Name, Gbl_Unmasked_Rev_TableName) Then
                ' We can erase the table's contents
                mStrOutSQL = "TRUNCATE TABLE " & Gbl_Unmasked_Rev_TableName
            Else
                ' We need to create it (we'll copy the previous year's table structure)
                mStrOutSQL = "SELECT * INTO " & Gbl_Unmasked_Rev_TableName & " FROM WB" &
                    CStr(CInt(cmb_URCSYear.Text) - 1) & "_Unmasked_Rev WHERE 1=2"
            End If

            ' Tell SQL to execute the command
            mSQLCmd = New SqlCommand
            mSQLCmd.CommandType = CommandType.Text
            mSQLCmd.Connection = gbl_SQLConnection
            mSQLCmd.CommandText = mStrOutSQL
            mSQLCmd.ExecuteNonQuery()

            txt_StatusBox.Text = "Getting Waybill Data From Server"
            Refresh()

            mStrSQL = New StringBuilder
            mStrSQL.Append("SELECT ")
            mStrSQL.Append("serial_no,")
            mStrSQL.Append("report_rr,")
            mStrSQL.Append("orr,")
            mStrSQL.Append("jrr1,")
            mStrSQL.Append("jrr2,")
            mStrSQL.Append("jrr3,")
            mStrSQL.Append("jrr4,")
            mStrSQL.Append("jrr5,")
            mStrSQL.Append("jrr6,")
            mStrSQL.Append("trr,")
            mStrSQL.Append("rate_flg,")
            mStrSQL.Append("stcc,")
            mStrSQL.Append("stcc_w49,")
            mStrSQL.Append("jf,")
            mStrSQL.Append("total_rev,")
            mStrSQL.Append("orr_rev,")
            mStrSQL.Append("jrr1_rev,")
            mStrSQL.Append("jrr2_rev,")
            mStrSQL.Append("jrr3_rev,")
            mStrSQL.Append("jrr4_rev,")
            mStrSQL.Append("jrr5_rev,")
            mStrSQL.Append("jrr6_rev,")
            mStrSQL.Append("trr_rev,")
            mStrSQL.Append("wb_num,")
            mStrSQL.Append("u_car_init,")
            mStrSQL.Append("wb_date,")
            mStrSQL.Append("acct_period,")
            mStrSQL.Append("t_fsac,")
            mStrSQL.Append("u_car_num,")
            mStrSQL.Append("u_tc_num,")
            mStrSQL.Append("t_st,")
            mStrSQL.Append("u_cars,")
            mStrSQL.Append("u_rev,")
            mStrSQL.Append("exp_factor_th")
            mStrSQL.Append(" FROM " & Get_Table_Name_From_SQL(cmb_URCSYear.Text, "MASKED"))

            OpenSQLConnection(Get_Database_Name_From_SQL(cmb_URCSYear.Text, "MASKED"))

            mDataTable = New DataTable
            Using daAdapter As New SqlDataAdapter(mStrSQL.ToString, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            mMaxRecs = mDataTable.Rows.Count
            Rec = 1

            'Loop through the datatable until we hit the end of the set

            For mlooper = 0 To mDataTable.Rows.Count - 1

                Application.DoEvents()

                mSerial_No = '0'
                mWbNum = 0
                mRptRR = 0
                morr = 0
                mjrr1 = 0
                mjrr2 = 0
                mjrr3 = 0
                mjrr4 = 0
                mjrr5 = 0
                mjrr6 = 0
                mtrr = 0
                mRate_Flg = 0
                mSTCC = 0
                mSTCC_W49 = 0
                mJF = 0
                mTotal_Rev = 0
                mORR_Rev = 0
                mJRR1_Rev = 0
                mJRR2_Rev = 0
                mJRR3_Rev = 0
                mJRR4_Rev = 0
                mJRR5_Rev = 0
                mJRR6_Rev = 0
                mTRR_Rev = 0
                mCar_Init = " "
                mDate = New Date
                mYear = 0
                mAcctMonth = 0
                mAcctYear = 0
                mFSAC = 0
                mUCarNum = 0
                mTCNum = 0
                mTState = 0
                mUCars = 0
                mU_Rev = 0
                mExp_Factor_Th = 0

                If Rec Mod 100 = 0 Then
                    txt_StatusBox.Text = "Writing record " & CStr(Rec) & " of " & mMaxRecs
                    Refresh()
                End If

                mSerial_No = ReturnString(mDataTable.Rows(mlooper)("serial_no"))
                mWbNum = mDataTable.Rows(mlooper)("wb_num")
                mCar_Init = ReturnString(Strings.Left(mDataTable.Rows(mlooper)("u_car_init"), 1), 1)
                mDate = ReturnDate(mDataTable.Rows(mlooper)("wb_date"))
                mYear = cmb_URCSYear.Text
                If Len(mDataTable.Rows(mlooper)("acct_period")) = 5 Then
                    mAcctMonth = CInt(Strings.Left(mDataTable.Rows(mlooper)("acct_period"), 1))
                Else
                    mAcctMonth = CInt(Strings.Left(mDataTable.Rows(mlooper)("acct_period"), 2))
                End If
                mAcctYear = CInt(Strings.Right(mDataTable.Rows(mlooper)("acct_period"), 4))
                ' Make sure that we get a 4 position year by faking it
                mWorkDate = CDate("1/1/" & CStr(mAcctYear))
                mAcctYear = Year(mWorkDate)
                mFSAC = mDataTable.Rows(mlooper)("t_fsac")
                mUCarNum = mDataTable.Rows(mlooper)("u_car_num")
                mTCNum = CDec(mDataTable.Rows(mlooper)("u_tc_num"))
                mTState = ReturnString(mDataTable.Rows(mlooper)("t_st"), 2)
                mUCars = mDataTable.Rows(mlooper)("u_cars")
                mRptRR = mDataTable.Rows(mlooper)("report_rr")
                mRate_Flg = mDataTable.Rows(mlooper)("rate_flg")
                mSTCC_W49 = mDataTable.Rows(mlooper)("stcc_w49")
                mTotal_Rev = CLng(mDataTable.Rows(mlooper)("total_rev"))
                mORR_Rev = CLng(mDataTable.Rows(mlooper)("orr_rev"))
                mJRR1_Rev = CLng(mDataTable.Rows(mlooper)("jrr1_rev"))
                mJRR2_Rev = CLng(mDataTable.Rows(mlooper)("jrr2_rev"))
                mJRR3_Rev = CLng(mDataTable.Rows(mlooper)("jrr3_rev"))
                mJRR4_Rev = CLng(mDataTable.Rows(mlooper)("jrr4_rev"))
                mJRR5_Rev = CLng(mDataTable.Rows(mlooper)("jrr5_rev"))
                mJRR6_Rev = CLng(mDataTable.Rows(mlooper)("jrr6_rev"))
                mTRR_Rev = CLng(mDataTable.Rows(mlooper)("trr_rev"))
                mU_Rev = mDataTable.Rows(mlooper)("u_rev")
                mExp_Factor_Th = mDataTable.Rows(mlooper)("exp_factor_th")

                'start building the SQL statement to run later
                mStrOutSQL = "INSERT INTO " & Gbl_Unmasked_Rev_TableName & "(" &
                        "[unmasked_serial_no], "

                'We have years that require the waybill number for a key reference (1994, 1995)
                Select Case Val(cmb_URCSYear.Text)
                    Case 1984, 1985, 1994, 1995
                        mStrOutSQL = mStrOutSQL & "[unmasked_wb_num], "
                End Select

                ' Back to building the string
                mStrOutSQL = mStrOutSQL &
                    "[total_unmasked_rev], " &
                    "[orr_unmasked_rev], " &
                    "[jrr1_unmasked_rev], " &
                    "[jrr2_unmasked_rev], " &
                    "[jrr3_unmasked_rev], " &
                    "[jrr4_unmasked_rev], " &
                    "[jrr5_unmasked_rev], " &
                    "[jrr6_unmasked_rev], " &
                    "[trr_unmasked_rev], " &
                    "[u_rev_unmasked]"

                mStrOutSQL = mStrOutSQL & ") VALUES ("                               'do not delete

                ' Unmask the rev values based on reporting railroad where rate_flg = 1
                ' If the flag isn't set then the Rev values remain the same as read

                If mRate_Flg = 1 Then

                    Select Case CInt(mRptRR)

                        Case 105    'CP
                            If Val(cmb_URCSYear.Text) < 1992 Then
                                'Do nothing - they didn't mask rates prior to Year 1992
                            Else
                                mU_Rev = UnmaskCPValue(0, mSTCC_W49, mU_Rev)
                            End If

                        Case 555    'NS
                            If Val(cmb_URCSYear.Text) < 1990 Then
                                'Do nothing - they didn't mask rates prior to Year 1990
                            Else
                                mU_Rev = UnmaskNSValue(0, mDate, mAcctYear, mAcctMonth, mFSAC, mWbNum, mUCarNum, mTCNum, mU_Rev)
                            End If

                        Case 712    'CSX
                            If Val(cmb_URCSYear.Text) < 1990 Then
                                'Do nothing - they didn't mask rates prior to Year 1990
                            Else
                                mU_Rev = UnmaskCSXValue(mSerial_No, mSTCC_W49, mWbNum, mAcctMonth, mAcctYear, mU_Rev)
                            End If

                        Case 777    'BNSF
                            If Val(cmb_URCSYear.Text) < 2000 Then
                                'Do nothing - they didn't mask rates prior to Year 2000
                            Else
                                mU_Rev = UnmaskBNSFValue(0, mCar_Init, mAcctMonth, mU_Rev)
                            End If

                        Case 802  'UP
                            If Val(cmb_URCSYear.Text) < 1990 Then
                                'Do nothing - they didn't mask rates prior to Year 1990
                            Else
                                mU_Rev = UnmaskUPValue(mSerial_No, Val(mSTCC_W49), Year(mDate), mWbNum, mU_Rev)
                            End If

                        Case Else   'All other railroads
                            If Val(cmb_URCSYear.Text) < 2001 Then
                                'Do nothing - they didn't mask rates prior to Year 2001
                            Else
                                mU_Rev = UnmaskGenericValue(mSerial_No, mSTCC_W49, mDate, mYear, mU_Rev)
                            End If
                    End Select

                    If mTotal_Rev > 0 Then
                        mTotal_Unmask_Rev = mU_Rev * mExp_Factor_Th
                        mORR_Rev = (mORR_Rev / mTotal_Rev) * mTotal_Unmask_Rev
                        mORR_Rev = Math.Round(mORR_Rev)
                        mJRR1_Rev = (mJRR1_Rev / mTotal_Rev) * mTotal_Unmask_Rev
                        mJRR1_Rev = Math.Round(mJRR1_Rev)
                        mJRR2_Rev = (mJRR2_Rev / mTotal_Rev) * mTotal_Unmask_Rev
                        mJRR2_Rev = Math.Round(mJRR2_Rev)
                        mJRR3_Rev = (mJRR3_Rev / mTotal_Rev) * mTotal_Unmask_Rev
                        mJRR3_Rev = Math.Round(mJRR3_Rev)
                        mJRR4_Rev = (mJRR4_Rev / mTotal_Rev) * mTotal_Unmask_Rev
                        mJRR4_Rev = Math.Round(mJRR4_Rev)
                        mJRR5_Rev = (mJRR5_Rev / mTotal_Rev) * mTotal_Unmask_Rev
                        mJRR5_Rev = Math.Round(mJRR5_Rev)
                        mJRR6_Rev = (mJRR6_Rev / mTotal_Rev) * mTotal_Unmask_Rev
                        mJRR6_Rev = Math.Round(mJRR6_Rev)
                        mTRR_Rev = (mTRR_Rev / mTotal_Rev) * mTotal_Unmask_Rev
                        mTRR_Rev = Math.Round(mTRR_Rev)
                        If mJF > 0 Then
                            mTRR_Rev = mTRR_Rev + (mTotal_Unmask_Rev - (mORR_Rev + mJRR1_Rev +
                                    mJRR2_Rev + mJRR3_Rev + mJRR4_Rev + mJRR5_Rev + mJRR6_Rev + mTRR_Rev))
                        End If
                        mTotal_Rev = CDec(mTotal_Unmask_Rev)
                    End If
                End If

                ' lay in the values starting with the serial_no to the back half of the SQL statement
                mStrOutSQL = mStrOutSQL & "'" & CStr(mSerial_No) & "', "

                ' Accomodate the problem years again
                Select Case Val(cmb_URCSYear.Text)
                    Case 1984, 1985, 1994, 1995
                        mStrOutSQL = mStrOutSQL & CStr(mWbNum) & ", "
                End Select

                mStrOutSQL = mStrOutSQL & mTotal_Rev & ", "
                mStrOutSQL = mStrOutSQL & mORR_Rev & ", "
                mStrOutSQL = mStrOutSQL & mJRR1_Rev & ", "
                mStrOutSQL = mStrOutSQL & mJRR2_Rev & ", "
                mStrOutSQL = mStrOutSQL & mJRR3_Rev & ", "
                mStrOutSQL = mStrOutSQL & mJRR4_Rev & ", "
                mStrOutSQL = mStrOutSQL & mJRR5_Rev & ", "
                mStrOutSQL = mStrOutSQL & mJRR6_Rev & ", "
                mStrOutSQL = mStrOutSQL & mTRR_Rev & ", "
                mStrOutSQL = mStrOutSQL & mU_Rev & ")"

                'execute the SQL statement
                mSQLCmd = New SqlCommand
                mSQLCmd.Connection = gbl_SQLConnection
                mSQLCmd.CommandType = CommandType.Text
                mSQLCmd.CommandText = mStrOutSQL.ToString
                mSQLCmd.ExecuteNonQuery()

                Rec = Rec + 1

            Next

            txt_StatusBox.Text = "Done!"
            Refresh()

        End If

EndIt:
    End Sub
End Class