Imports System.Data.SqlClient
Imports System.Text
Module Sample_Generator_Routines

    Sub Annual_Sample_By_State_Writer(ByVal mDatabase_Name As String,
                                      ByVal mTable_Name As String,
                                      ByVal mUnMasked_Table_Name As String,
                                      ByVal mState As String,
                                      ByVal mUnmasked As Boolean,
                                      ByVal mRecords As Decimal,
                                      ByVal mOutputFilePath As String,
                                      ByVal mFormat As String)

        Dim mStrSQL As String, mStrOutLine As String
        Dim mSTCC As String, mSTCC_W49 As String
        Dim mWorkState As String
        Dim outfile
        Dim Rec As Long
        Dim mRate_Flg As Integer
        Dim mTotal_Rev As Single, mORR_Rev As Single, mJRR1_Rev As Single
        Dim mJRR2_Rev As Single, mJRR3_Rev As Single, mJRR4_Rev As Single
        Dim mJRR5_Rev As Single, mJRR6_Rev As Single, mTRR_Rev As Single
        Dim mU_Rev As Single

        Dim mWaybill As Class_Waybill
        Dim rst As ADODB.Recordset

        mWorkState = Trim(mState)

        ' Open/Check the SQL connection
        OpenADOConnection(mDatabase_Name)

        'Open the output file - create it if necessary and always append
        outfile = My.Computer.FileSystem.OpenTextFileWriter(mOutputFilePath, False, System.Text.ASCIIEncoding.ASCII)

        ''Display the blank progress Form
        'ProgressBar_Form.Text = "Fetching Records from SQL Server.  Please Wait"
        'ProgressBar_Form.Show()

        ''Add the parameters for the progress bar control
        'ProgressBar_Form.ProgressBar.Min = 0
        'ProgressBar_Form.ProgressBar.Max = 100
        'ProgressBar_Form.ProgressBar.Value = 0

        'Build the SQL select statement.

        '*********************************************************************************
        ' Modified 10 Aug 2015 by Michael Sanders
        ' Modification is to SQL select statement to pick up records where the state flag
        ' or the originating, terminating or junction alpha fields contains the state
        ' two-letter alpha field matches the mState argument.
        '*********************************************************************************
        ' Original Method prior to 2015 processing
        'If mUnmasked = True Then
        '    mStrSQL = "SELECT * FROM " & mTable_Name & " INNER JOIN " & mUnMasked_Table_Name & _
        '        " ON " & mTable_Name & ".Serial_No = " & mUnMasked_Table_Name & ".Unmasked_Serial_No" & _
        '        " WHERE (" & Trim(mState) & "_Flg = 1)"
        'Else
        '    mStrSQL = "SELECT * FROM " & mTable_Name & " WHERE (" & _
        '    Trim(mState) & "_Flg = 1)"
        'End If

        ' New mothod for years 2015+
        If mUnmasked = True Then
            mStrSQL = "SELECT * FROM " & mTable_Name & " INNER JOIN " & mUnMasked_Table_Name &
                " ON " & mTable_Name & ".Serial_No = " & mUnMasked_Table_Name & ".Unmasked_Serial_No" &
                " WHERE (" & mWorkState & "_Flg = 1) " &
                "Or ((O_ST = '" & mWorkState & "') " &
                "Or (JCT1_ST = '" & mWorkState & "') " &
                "Or (JCT2_ST = '" & mWorkState & "') " &
                "Or (JCT3_ST = '" & mWorkState & "') " &
                "Or (JCT4_ST = '" & mWorkState & "') " &
                "Or (JCT5_ST = '" & mWorkState & "') " &
                "Or (JCT6_ST = '" & mWorkState & "') " &
                "Or (JCT7_ST = '" & mWorkState & "') " &
                "Or (T_ST = '" & mWorkState & "'))"
        Else
            mStrSQL = "SELECT * FROM " & mTable_Name & " WHERE (" &
            mWorkState & "_Flg = 1) " &
                "Or ((O_ST = '" & mWorkState & "') " &
                "Or (JCT1_ST = '" & mWorkState & "') " &
                "Or (JCT2_ST = '" & mWorkState & "') " &
                "Or (JCT3_ST = '" & mWorkState & "') " &
                "Or (JCT4_ST = '" & mWorkState & "') " &
                "Or (JCT5_ST = '" & mWorkState & "') " &
                "Or (JCT6_ST = '" & mWorkState & "') " &
                "Or (JCT7_ST = '" & mWorkState & "') " &
                "Or (T_ST = '" & mWorkState & "'))"
        End If

        rst = SetRST()

        rst.Open(mStrSQL, gbl_ADOConnection)

        'Display the blank progress Form
        'ProgressBar_Form.Text = "Exporting " & CStr(mRecords) & " Records for " & mState & "..."
        'ProgressBar_Form.Show()

        If rst.RecordCount > 0 Then
            Rec = 1

            rst.MoveFirst()

            'loop through the recordset until we hit the end
            Do While rst.EOF <> True

                If Rec Mod 100 = 0 Then
                    'ProgressBar_Form.Text = CStr(Math.Round((Rec / rst.RecordCount) * 100, 1)) & "% - Writing Records to File for " &
                    '    mState & "..."
                    'ProgressBar_Form.ProgressBar.Value = Math.Round((Rec / mRecords) * 100, 1)
                    'ProgressBar_Form.ProgressBar.TextShow = ProgBar.ProgBarPlus.eTextShow.Percent
                    Application.DoEvents()
                End If

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
                mTotal_Rev = rst.Fields("total_rev").Value
                mORR_Rev = rst.Fields("orr_rev").Value
                mJRR1_Rev = rst.Fields("jrr1_rev").Value
                mJRR2_Rev = rst.Fields("jrr2_rev").Value
                mJRR3_Rev = rst.Fields("jrr3_rev").Value
                mJRR4_Rev = rst.Fields("jrr4_rev").Value
                mJRR5_Rev = rst.Fields("jrr5_rev").Value
                mJRR6_Rev = rst.Fields("jrr6_rev").Value
                mTRR_Rev = rst.Fields("trr_rev").Value
                mU_Rev = rst.Fields("u_rev").Value

                'Set the rate flag to zero
                mRate_Flg = 0

                'Initialize the Waybill class
                mWaybill = New Class_Waybill

                If mUnmasked = True Then
                    'Get unmasked values
                    mTotal_Rev = rst.Fields("Total_Unmasked_Rev").Value
                    mORR_Rev = rst.Fields("ORR_Unmasked_Rev").Value
                    mJRR1_Rev = rst.Fields("JRR1_Unmasked_Rev").Value
                    mJRR2_Rev = rst.Fields("JRR2_Unmasked_Rev").Value
                    mJRR3_Rev = rst.Fields("JRR3_Unmasked_Rev").Value
                    mJRR4_Rev = rst.Fields("JRR4_Unmasked_Rev").Value
                    mJRR5_Rev = rst.Fields("JRR5_Unmasked_Rev").Value
                    mJRR6_Rev = rst.Fields("JRR6_Unmasked_Rev").Value
                    mTRR_Rev = rst.Fields("TRR_Unmasked_Rev").Value
                    mU_Rev = rst.Fields("U_Rev_Unmasked").Value
                    mRate_Flg = rst.Fields("rate_flg").Value
                End If

                'We do, however, have to mask the STCC codes
                mSTCC = rst.Fields("stcc").Value
                mSTCC_W49 = rst.Fields("stcc_w49").Value
                If Left(Trim(mSTCC), 2) = "19" Then
                    mSTCC = "1900000"
                    If Left(Trim(mSTCC_W49), 2) = "19" Then
                        mSTCC_W49 = "1900000"
                    End If
                    If Left(Trim(mSTCC_W49), 2) = "49" Then
                        mSTCC_W49 = "4900000"
                    End If
                End If

                With mWaybill
                    'Load the unmasked/masked variables to the class
                    .STCC_W49 = mSTCC_W49
                    .Rate_Flg = mRate_Flg
                    .STCC = mSTCC
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
                End With

                'pass the serial_no variable to the class loader
                'resulting in the remaining values being loaded to the class.
                LoadWaybillClass(rst.Fields("serial_no").Value, mWaybill, rst, mFormat)

                'build the outfile record string
                mStrOutLine = ""
                Select Case mFormat
                    Case "913", "900"
                        mStrOutLine = Build913String(mWaybill)
                End Select

                'the line is now built and can be written to the output file
                outfile.writeline(mStrOutLine)

                'unload the mWaybill class
                mWaybill = Nothing

                Rec = Rec + 1
                rst.MoveNext()

            Loop

            'close the output file
            outfile.Close()

            ' Close the Progress Bar form
            'ProgressBar_Form.Close()

        End If

        rst.Close()
        rst = Nothing

    End Sub

    Sub Annual_Sample_By_Selected_States_Writer(ByVal mDatabase_Name As String,
                                      ByVal mTable_Name As String,
                                      ByVal mUnMasked_Table_Name As String,
                                      ByVal mSelectedStates As ListBox.SelectedObjectCollection,
                                      ByVal mUnmasked As Boolean,
                                      ByVal mRecords As Decimal,
                                      ByVal mOutputFilePath As String,
                                      ByVal mFormat As String)

        '*********************************************************************************
        ' Created 12 Jun 2018 by Michael Sanders
        ' Created to allow for selection of multiple state records and to utilize DataTable
        ' instead of ADODB RecordSet.  Original source code copy saved to zip file v1
        '*********************************************************************************

        Dim mStrSQL As String, mStrOutLine As String
        Dim mSTCC As String, mSTCC_W49 As String
        Dim outfile
        Dim Rec As Long
        Dim mRate_Flg As Integer
        Dim mTotal_Rev As Single, mORR_Rev As Single, mJRR1_Rev As Single
        Dim mJRR2_Rev As Single, mJRR3_Rev As Single, mJRR4_Rev As Single
        Dim mJRR5_Rev As Single, mJRR6_Rev As Single, mTRR_Rev As Single
        Dim mU_Rev As Single

        Dim mWaybill As Class_Waybill
        Dim mDatatable As New DataTable

        ' Open/Check the SQL connection
        OpenSQLConnection(mDatabase_Name)

        'Open the output file - create it if necessary and always append
        outfile = My.Computer.FileSystem.OpenTextFileWriter(mOutputFilePath, False, Encoding.ASCII)

        'Build the SQL select statement.

        '*********************************************************************************
        ' Modified 10 Aug 2015 by Michael Sanders
        ' Modification is to SQL select statement to pick up records where the state flag
        ' or the originating, terminating or junction alpha fields contains the state
        ' two-letter alpha field matches the mState argument.
        '*********************************************************************************
        ' Original Method prior to 2015 processing
        'If mUnmasked = True Then
        '    mStrSQL = "SELECT * FROM " & mTable_Name & " INNER JOIN " & mUnMasked_Table_Name & _
        '        " ON " & mTable_Name & ".Serial_No = " & mUnMasked_Table_Name & ".Unmasked_Serial_No" & _
        '        " WHERE (" & Trim(mState) & "_Flg = 1)"
        'Else
        '    mStrSQL = "SELECT * FROM " & mTable_Name & " WHERE (" & _
        '    Trim(mState) & "_Flg = 1)"
        'End If

        ' New method
        If mUnmasked = True Then
            mStrSQL = "SELECT * FROM " & Trim(mTable_Name) & " INNER JOIN " & Trim(mUnMasked_Table_Name) &
                " ON " & Trim(mTable_Name) & ".Serial_No = " & Trim(mUnMasked_Table_Name) & ".Unmasked_Serial_No" &
                " WHERE "
        Else
            mStrSQL = "SELECT * FROM " & mTable_Name & " WHERE "
        End If

        For mLooper = 1 To mSelectedStates.Count                                                    'Count is 1 based
            mStrSQL = mStrSQL & Trim(mSelectedStates.Item(mLooper - 1).ToString) & "_Flg = 1 OR "   'Position is zero based
        Next mLooper

        ' That takes care of the state flgs.  Now to address the origin, termination and junction states
        ' Origin states
        For mLooper = 1 To mSelectedStates.Count
            mStrSQL = mStrSQL & "O_ST = '" & Trim(mSelectedStates.Item(mLooper - 1).ToString) & "' OR "
        Next mLooper

        ' Termination States
        For mLooper = 1 To mSelectedStates.Count
            mStrSQL = mStrSQL & "T_ST = '" & Trim(mSelectedStates.Item(mLooper - 1).ToString) & "' OR "
        Next

        ' Junction states
        For mLooper = 1 To mSelectedStates.Count

            For mJunctionLooper = 1 To 7
                mStrSQL = mStrSQL & "JCT" & mJunctionLooper.ToString & "_ST = '" & Trim(mSelectedStates.Item(mLooper - 1).ToString) & "'"
                If mJunctionLooper <> 7 Then
                    mStrSQL = mStrSQL & " OR "
                End If
            Next mJunctionLooper

            If mLooper <> mSelectedStates.Count Then
                mStrSQL = mStrSQL & " OR "
            End If

        Next mLooper

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        If mDatatable.Rows.Count > 0 Then
            Rec = 1

            'loop through the recordset until we hit the end
            For mLooper = 0 To mDatatable.Rows.Count - 1

                Application.DoEvents()

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

                If mUnmasked = True Then
                    'load the Unmasked values into the memvars
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
                Else
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
                End If

                'We do, however, have to mask the STCC codes
                mSTCC = mDatatable.Rows(mLooper)("stcc")
                mSTCC_W49 = mDatatable.Rows(mLooper)("stcc_w49")
                If Left(Trim(mSTCC), 2) = "19" Then
                    mSTCC = "1900000"
                    If Left(Trim(mSTCC_W49), 2) = "19" Then
                        mSTCC_W49 = "1900000"
                    End If
                    If Left(Trim(mSTCC_W49), 2) = "49" Then
                        mSTCC_W49 = "4900000"
                    End If
                End If

                'Initialize the Waybill class
                mWaybill = New Class_Waybill

                With mWaybill
                    'Load the unmasked/masked variables to the class
                    .STCC_W49 = mSTCC_W49
                    .Rate_Flg = mRate_Flg
                    .STCC = mSTCC
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
                End With

                'pass the serial_no variable to the class loader
                'resulting in the remaining values being loaded to the class.
                Load_Waybill_Class_From_Datarow(mWaybill, mDatatable.Rows(mLooper))

                'build the outfile record string
                mStrOutLine = ""
                Select Case mFormat
                    Case "913"
                        mStrOutLine = Build913String(mWaybill)
                End Select

                'the line is now built and can be written to the output file
                outfile.writeline(mStrOutLine)

                'unload the mWaybill class
                mWaybill = Nothing

                Rec = Rec + 1

            Next

            'close the output file
            outfile.Close()

        End If

        mDatatable = Nothing

    End Sub

    Sub Annual_Sample_By_RR_Writer(ByVal mYear As Integer,
        ByVal mRailroad As Integer,
        ByVal mUnMasked As Boolean,
        ByVal mUnMasked_STCC As Boolean,
        ByVal UnmaskedLitigation As Boolean,
        ByVal mDatabaseName As String,
        ByVal mMaskedDataTableName As String,
        ByVal mUnMaskedDataTableName As String,
        ByVal mOutputFilePath As String,
        ByVal mFormat As Integer,
        ByVal mShowProgress As Boolean,
        ByVal mRailroadName As String,
        ByVal mAsReported As Boolean)

        Dim mStrSQL As String, mStrOutLine As String
        Dim mSTCC As String, mSTCC_W49 As String
        Dim outfile As StreamWriter
        Dim mRate_Flg As Integer
        Dim mTotal_Rev As Decimal, mORR_Rev As Decimal, mJRR1_Rev As Decimal
        Dim mJRR2_Rev As Decimal, mJRR3_Rev As Decimal, mJRR4_Rev As Decimal
        Dim mJRR5_Rev As Decimal, mJRR6_Rev As Decimal, mTRR_Rev As Decimal
        Dim mU_Rev As Decimal
        Dim mRec As Integer

        Dim mDataTable As DataTable
        Dim mWaybill As Class_Waybill

        ' Open/Check the SQL connection
        OpenSQLConnection(mDatabaseName)

        mDataTable = New DataTable

        mStrSQL = ""

        Select Case mUnMasked
            Case True
                ' We have to make accomodations for CP having both 105 and 482 codes
                If (mRailroad = 105) Or (mRailroad = 482) Then
                    mStrSQL = "SELECT " &
                        "dbo.WB" & mMaskedDataTableName & ".*, " &
                        "dbo.WB" & mUnMaskedDataTableName & ".*, " &
                        "report_rr AS Expr1 " &
                        "FROM dbo." & mMaskedDataTableName & " INNER JOIN " &
                        "dbo." & mUnMaskedDataTableName & " ON " &
                        "dbo." & mMaskedDataTableName & ".Serial_No = " &
                        "dbo." & mUnMaskedDataTableName & ".Unmasked_Serial_no And " &
                        "dbo." & mUnMaskedDataTableName & ".Unmasked_WB_Num " &
                        "WHERE (dbo." & mMaskedDataTableName & ".report_rr = 105 Or " &
                        "dbo." & mMaskedDataTableName & ".report_rr = 482 Or " &
                        "dbo." & mMaskedDataTableName & ".orr = 105 Or dbo." & mMaskedDataTableName & ".orr = 482 Or " &
                        "dbo." & mMaskedDataTableName & ".jrr1 = 105 Or dbo." & mMaskedDataTableName & ".jrr1 = 482 Or " &
                        "dbo." & mMaskedDataTableName & ".jrr2 = 105 Or dbo." & mMaskedDataTableName & ".jrr2 = 482 Or " &
                        "dbo." & mMaskedDataTableName & ".jrr3 = 105 Or dbo." & mMaskedDataTableName & ".jrr3 = 482 Or " &
                        "dbo." & mMaskedDataTableName & ".jrr4 = 105 Or dbo." & mMaskedDataTableName & ".jrr4 = 482 Or " &
                        "dbo." & mMaskedDataTableName & ".jrr5 = 105 Or dbo." & mMaskedDataTableName & ".jrr5 = 482 Or " &
                        "dbo." & mMaskedDataTableName & ".jrr6 = 105 Or dbo." & mMaskedDataTableName & ".jrr6 = 482 Or " &
                        "dbo." & mMaskedDataTableName & ".trr = 105 Or dbo." & mMaskedDataTableName & ".trr = 482)"
                Else
                    mStrSQL = "SELECT dbo." & mMaskedDataTableName & ".*, dbo." & mUnMaskedDataTableName & ".* " &
                        "FROM dbo." & mUnMaskedDataTableName & " INNER JOIN " &
                        "dbo." & mMaskedDataTableName & " ON dbo." & mMaskedDataTableName & ".Serial_No = " &
                        "dbo." & mUnMaskedDataTableName & ".Unmasked_Serial_no " &
                        "WHERE (dbo." & mMaskedDataTableName & ".report_rr = " & mRailroad & " Or " &
                        "dbo." & mMaskedDataTableName & ".orr = " & mRailroad & " Or " &
                        "dbo." & mMaskedDataTableName & ".jrr1 = " & mRailroad & " Or " &
                        "dbo." & mMaskedDataTableName & ".jrr2 = " & mRailroad & " Or " &
                        "dbo." & mMaskedDataTableName & ".jrr3 = " & mRailroad & " Or " &
                        "dbo." & mMaskedDataTableName & ".jrr4 = " & mRailroad & " Or " &
                        "dbo." & mMaskedDataTableName & ".jrr5 = " & mRailroad & " Or " &
                        "dbo." & mMaskedDataTableName & ".jrr6 = " & mRailroad & " Or " &
                        "dbo." & mMaskedDataTableName & ".trr = " & mRailroad & ")"
                End If
            Case False
                If (mRailroad = 105) Or (mRailroad = 482) Then
                    mStrSQL = "SELECT * FROM dbo." & mMaskedDataTableName & " " &
                        "WHERE (report_rr = 105 Or report_rr = 482 Or " &
                        "orr = 105 Or orr = 482 Or " &
                        "jrr1 = 105 Or jrr1 = 482 Or " &
                        "jrr2 = 105 Or jrr2 = 482 Or " &
                        "jrr3 = 105 Or jrr3 = 482 Or " &
                        "jrr4 = 105 Or jrr4 = 482 Or " &
                        "jrr5 = 105 Or jrr5 = 482 Or " &
                        "jrr6 = 105 Or jrr6 = 482 Or " &
                        "trr = 105 Or trr = 482)"
                Else
                    mStrSQL = "SELECT * FROM dbo." & mMaskedDataTableName & " " &
                        "WHERE report_rr = " & mRailroad & " Or " &
                        "orr = " & mRailroad & " Or " &
                        "jrr1 = " & mRailroad & " Or " &
                        "jrr2 = " & mRailroad & " Or " &
                        "jrr3 = " & mRailroad & " Or " &
                        "jrr4 = " & mRailroad & " Or " &
                        "jrr5 = " & mRailroad & " Or " &
                        "jrr6 = " & mRailroad & " Or " &
                        "trr = " & mRailroad
                End If
        End Select

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        'Perform the write only if the record count > 0
        If mDataTable.Rows.Count > 0 Then

            mRec = 1

            ' Delete the output file if it exists
            If My.Computer.FileSystem.FileExists(mOutputFilePath) Then
                My.Computer.FileSystem.DeleteFile(mOutputFilePath)
            End If

            ' Create the output file
            outfile = New StreamWriter(mOutputFilePath, False)

            mRec = 1

            'loop through the recordset until we hit the end of the set
            For mLooper = 0 To mDataTable.Rows.Count - 1


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
                mRate_Flg = 0

                ' Gather the information that will be manipulated if necessary or used by functions
                mSTCC = mDataTable.Rows(mLooper)("stcc")
                mSTCC_W49 = mDataTable.Rows(mLooper)("stcc_w49")

                If mRailroad = mDataTable.Rows(mLooper)("report_rr") Then
                    Select Case mUnMasked
                        Case True
                            mTotal_Rev = mDataTable.Rows(mLooper)("total_unmasked_rev")
                            mORR_Rev = mDataTable.Rows(mLooper)("orr_unmasked_rev")
                            mJRR1_Rev = mDataTable.Rows(mLooper)("jrr1_unmasked_rev")
                            mJRR2_Rev = mDataTable.Rows(mLooper)("jrr2_unmasked_rev")
                            mJRR3_Rev = mDataTable.Rows(mLooper)("jrr3_unmasked_rev")
                            mJRR4_Rev = mDataTable.Rows(mLooper)("jrr4_unmasked_rev")
                            mJRR5_Rev = mDataTable.Rows(mLooper)("jrr5_unmasked_rev")
                            mJRR6_Rev = mDataTable.Rows(mLooper)("jrr6_unmasked_rev")
                            mTRR_Rev = mDataTable.Rows(mLooper)("trr_unmasked_rev")
                            mRate_Flg = mDataTable.Rows(mLooper)("rate_flg")
                            mU_Rev = mDataTable.Rows(mLooper)("u_rev_unmasked")
                        Case False
                            mTotal_Rev = mDataTable.Rows(mLooper)("total_rev")
                            mORR_Rev = mDataTable.Rows(mLooper)("orr_rev")
                            mJRR1_Rev = mDataTable.Rows(mLooper)("jrr1_rev")
                            mJRR2_Rev = mDataTable.Rows(mLooper)("jrr2_rev")
                            mJRR3_Rev = mDataTable.Rows(mLooper)("jrr3_rev")
                            mJRR4_Rev = mDataTable.Rows(mLooper)("jrr4_rev")
                            mJRR5_Rev = mDataTable.Rows(mLooper)("jrr5_rev")
                            mJRR6_Rev = mDataTable.Rows(mLooper)("jrr6_rev")
                            mTRR_Rev = mDataTable.Rows(mLooper)("trr_rev")
                            mU_Rev = mDataTable.Rows(mLooper)("u_rev")
                            If mAsReported = True Then
                                mRate_Flg = mDataTable.Rows(mLooper)("rate_flg")
                            Else
                                mRate_Flg = 0
                            End If
                    End Select
                Else
                    Select Case UnmaskedLitigation
                        Case True
                            mTotal_Rev = mDataTable.Rows(mLooper)("total_unmasked_rev")
                            mORR_Rev = mDataTable.Rows(mLooper)("orr_unmasked_rev")
                            mJRR1_Rev = mDataTable.Rows(mLooper)("jrr1_unmasked_rev")
                            mJRR2_Rev = mDataTable.Rows(mLooper)("jrr2_unmasked_rev")
                            mJRR3_Rev = mDataTable.Rows(mLooper)("jrr3_unmasked_rev")
                            mJRR4_Rev = mDataTable.Rows(mLooper)("jrr4_unmasked_rev")
                            mJRR5_Rev = mDataTable.Rows(mLooper)("jrr5_unmasked_rev")
                            mJRR6_Rev = mDataTable.Rows(mLooper)("jrr6_unmasked_rev")
                            mTRR_Rev = mDataTable.Rows(mLooper)("trr_unmasked_rev")
                            mRate_Flg = mDataTable.Rows(mLooper)("rate_flg")
                            mU_Rev = mDataTable.Rows(mLooper)("u_rev_unmasked")
                        Case False
                            mTotal_Rev = mDataTable.Rows(mLooper)("total_rev")
                            mORR_Rev = mDataTable.Rows(mLooper)("orr_rev")
                            mJRR1_Rev = mDataTable.Rows(mLooper)("jrr1_rev")
                            mJRR2_Rev = mDataTable.Rows(mLooper)("jrr2_rev")
                            mJRR3_Rev = mDataTable.Rows(mLooper)("jrr3_rev")
                            mJRR4_Rev = mDataTable.Rows(mLooper)("jrr4_rev")
                            mJRR5_Rev = mDataTable.Rows(mLooper)("jrr5_rev")
                            mJRR6_Rev = mDataTable.Rows(mLooper)("jrr6_rev")
                            mTRR_Rev = mDataTable.Rows(mLooper)("trr_rev")
                            mU_Rev = mDataTable.Rows(mLooper)("u_rev")
                            If mAsReported = True Then
                                mRate_Flg = mDataTable.Rows(mLooper)("rate_flg")
                            Else
                                mRate_Flg = 0
                            End If
                    End Select
                End If

                If mUnMasked_STCC = True Then
                    'leave the STCC fields as they are
                Else
                    If Left(Trim(mSTCC), 2) = "19" Then
                        mSTCC = "1900000"
                        If Left(Trim(mSTCC_W49), 2) = "19" Then
                            mSTCC_W49 = "1900000"
                        End If
                        If Left(Trim(mSTCC_W49), 2) = "49" Then
                            mSTCC_W49 = "4900000"
                        End If
                    End If
                End If

                'initialize the mwaybill class
                mWaybill = New Class_Waybill

                'Pass the serial_no variable to the class loader
                'resulting in the class being loaded with the
                'values from the current resultset record.
                LoadWaybillClass(mDataTable.Rows(mLooper)("serial_no"), mWaybill, mDataTable, mFormat)

                'load/replace the unmasked/masked variables into the class
                mWaybill.STCC_W49 = mSTCC_W49
                mWaybill.STCC = mSTCC
                mWaybill.Rate_Flg = mRate_Flg
                mWaybill.Total_Rev = mTotal_Rev
                mWaybill.ORR_Rev = mORR_Rev
                mWaybill.JRR1_Rev = mJRR1_Rev
                mWaybill.JRR2_Rev = mJRR2_Rev
                mWaybill.JRR3_Rev = mJRR3_Rev
                mWaybill.JRR4_Rev = mJRR4_Rev
                mWaybill.JRR5_Rev = mJRR5_Rev
                mWaybill.JRR6_Rev = mJRR6_Rev
                mWaybill.TRR_Rev = mTRR_Rev
                mWaybill.U_Rev = mU_Rev

                'Build the outfile record string
                mStrOutLine = Build913String(mWaybill)

                'the line is built - write it to file
                outfile.WriteLine(mStrOutLine)

                'unload the waybill info from the class
                mWaybill = Nothing

                mRec = mRec + 1
            Next

            'close the output file
            outfile.Flush()
            outfile.Close()

        End If

    End Sub

    Sub Annual_Sample_By_STCC(ByVal mYear As Integer,
        ByVal mSTCC As String,
        ByVal mUnMasked As Boolean,
        ByVal mUnMasked_STCC As Boolean,
        ByVal mDatabaseName As String,
        ByVal mMaskedDataTableName As String,
        ByVal mUnMaskedDataTableName As String,
        ByVal mOutputFilePath As String,
        ByVal mFormat As Integer,
        ByVal mDisplayProgress As Boolean,
        ByVal mOverwrite As Boolean)

        Dim mStrSQL As String, mStrOutLine As String
        Dim mThisSTCC_W49 As String, mThisSTCC As String
        Dim Rec As Long
        Dim outfile
        Dim mRate_Flg As Integer
        Dim mTotal_Rev As Decimal, mORR_Rev As Decimal, mJRR1_Rev As Decimal
        Dim mJRR2_Rev As Decimal, mJRR3_Rev As Decimal, mJRR4_Rev As Decimal
        Dim mJRR5_Rev As Decimal, mJRR6_Rev As Decimal, mTRR_Rev As Decimal
        Dim mU_Rev As Decimal
        Dim mLooper As Integer

        Dim mDataTable As DataTable
        Dim mWaybill As Class_Waybill

        ' Open/Check the SQL connection
        OpenSQLConnection(mDatabaseName)

        mDataTable = New DataTable

        ' Get the records

        If Mid(mSTCC, 1, 2) = "49" Then
            If mUnMasked = True Then
                mStrSQL = "SELECT * " &
                        "FROM dbo." & mMaskedDataTableName & " INNER JOIN " &
                        "dbo." & mUnMaskedDataTableName & " ON " &
                        "dbo." & mMaskedDataTableName & ".Serial_No = " &
                        "dbo." & mUnMaskedDataTableName & ".Unmasked_Serial_no " &
                        "WHERE (dbo." & mMaskedDataTableName & ".STCC_W49 Like '" & mSTCC & "%')"
            Else
                mStrSQL = "SELECT * FROM dbo." & mMaskedDataTableName &
                        " WHERE (STCC_W49 LIKE '" & mSTCC & "%')"
            End If
        Else
            If mUnMasked = True Then
                mStrSQL = "SELECT * " &
                        "FROM dbo." & mMaskedDataTableName & " INNER JOIN " &
                        "dbo." & mUnMaskedDataTableName & " ON " &
                        "dbo." & mMaskedDataTableName & ".Serial_No = " &
                        "dbo." & mUnMaskedDataTableName & ".Unmasked_Serial_no " &
                        "WHERE (dbo." & mMaskedDataTableName & ".STCC LIKE '" & mSTCC & "%')"
            Else
                mStrSQL = "SELECT * FROM dbo." & mMaskedDataTableName &
                        " WHERE (STCC LIKE '" & mSTCC & "%')"
            End If
        End If

        'Display the blank progress Form
        'ProgressBar_Form.Text = "Fetching Records from SQL Server.  Please Wait"
        'ProgressBar_Form.Show()

        ''Add the parameters for the progress bar control
        'ProgressBar_Form.ProgressBar.Min = 0
        'ProgressBar_Form.ProgressBar.Max = 100
        'ProgressBar_Form.ProgressBar.Value = 0

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        If mDataTable.Rows.Count > 0 Then

            If mOverwrite = True Then
                If My.Computer.FileSystem.FileExists(mOutputFilePath) = True Then
                    My.Computer.FileSystem.DeleteFile(mOutputFilePath)
                End If
            End If

            ' open the output file, appending to existing data
            outfile = My.Computer.FileSystem.OpenTextFileWriter(mOutputFilePath, True, System.Text.ASCIIEncoding.ASCII)

            'Display the blank progress Form
            'ProgressBar_Form.Text = "Exporting " & mDataTable.Rows.Count.ToString & " Records for STCC " & CStr(mSTCC)
            'ProgressBar_Form.Show()

            'loop through the recordset
            For mLooper = 0 To mDataTable.Rows.Count - 1
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

                If Rec Mod 100 = 0 Then
                    'ProgressBar_Form.Text = CStr(Math.Round((Rec / mDataTable.Rows.Count) * 100, 1)) & "% - Writing Records to File for " &
                    '    mSTCC & "..."
                    'ProgressBar_Form.ProgressBar.Value = Math.Round((Rec / mDataTable.Rows.Count) * 100, 1)
                    'ProgressBar_Form.ProgressBar.TextShow = ProgBar.ProgBarPlus.eTextShow.Percent
                    Application.DoEvents()
                End If

                ' Gather the information that will be manipulated if necessary or used by functions
                mThisSTCC = mDataTable.Rows(mLooper)("stcc") 'rst.Fields("stcc").Value
                mThisSTCC_W49 = mDataTable.Rows(mLooper)("stcc_w49") 'rst.Fields("stcc_w49").Value

                If mUnMasked = True Then
                    mRate_Flg = mDataTable.Rows(mLooper)("rate_flg")
                    mTotal_Rev = mDataTable.Rows(mLooper)("total_unmasked_rev")
                    mORR_Rev = mDataTable.Rows(mLooper)("orr_unmasked_rev")
                    mJRR1_Rev = mDataTable.Rows(mLooper)("jrr1_unmasked_rev")
                    mJRR2_Rev = mDataTable.Rows(mLooper)("jrr2_unmasked_rev")
                    mJRR3_Rev = mDataTable.Rows(mLooper)("jrr3_unmasked_rev")
                    mJRR4_Rev = mDataTable.Rows(mLooper)("jrr4_unmasked_rev")
                    mJRR5_Rev = mDataTable.Rows(mLooper)("jrr5_unmasked_rev")
                    mJRR6_Rev = mDataTable.Rows(mLooper)("jrr6_unmasked_rev")
                    mTRR_Rev = mDataTable.Rows(mLooper)("trr_unmasked_rev")
                    mU_Rev = mDataTable.Rows(mLooper)("u_rev_unmasked")
                Else
                    mRate_Flg = 0
                    mTotal_Rev = mDataTable.Rows(mLooper)("total_rev")
                    mORR_Rev = mDataTable.Rows(mLooper)("orr_rev")
                    mJRR1_Rev = mDataTable.Rows(mLooper)("jrr1_rev")
                    mJRR2_Rev = mDataTable.Rows(mLooper)("jrr2_rev")
                    mJRR3_Rev = mDataTable.Rows(mLooper)("jrr3_rev")
                    mJRR4_Rev = mDataTable.Rows(mLooper)("jrr4_rev")
                    mJRR5_Rev = mDataTable.Rows(mLooper)("jrr5_rev")
                    mJRR6_Rev = mDataTable.Rows(mLooper)("jrr6_rev")
                    mTRR_Rev = mDataTable.Rows(mLooper)("trr_rev")
                    mU_Rev = mDataTable.Rows(mLooper)("u_rev")
                End If

                If mUnMasked_STCC = True Then
                    'leave the STCC fields as they are
                Else
                    If Left(Trim(mThisSTCC), 2) = "19" Then
                        mSTCC = "1900000"
                        If Left(Trim(mThisSTCC_W49), 2) = "19" Then
                            mThisSTCC_W49 = "1900000"
                        End If
                        If Left(Trim(mThisSTCC_W49), 2) = "49" Then
                            mThisSTCC_W49 = "4900000"
                        End If
                    End If
                End If

                'initialize the waybill class
                mWaybill = New Class_Waybill

                'load the masked/unmasked data variables to the class
                mWaybill.STCC = mThisSTCC
                mWaybill.STCC_W49 = mThisSTCC_W49
                mWaybill.Rate_Flg = mRate_Flg
                mWaybill.Total_Rev = mTotal_Rev
                mWaybill.ORR_Rev = mORR_Rev
                mWaybill.JRR1_Rev = mJRR1_Rev
                mWaybill.JRR2_Rev = mJRR2_Rev
                mWaybill.JRR3_Rev = mJRR3_Rev
                mWaybill.JRR4_Rev = mJRR4_Rev
                mWaybill.JRR5_Rev = mJRR5_Rev
                mWaybill.JRR6_Rev = mJRR6_Rev
                mWaybill.TRR_Rev = mTRR_Rev
                mWaybill.U_Rev = mU_Rev

                'Pass the serial_no variable to allow the class loader to locate
                'the current resultset record
                LoadWaybillClass(mDataTable.Rows(mLooper)("serial_no"), mWaybill, mDataTable, mFormat)

                'build the outfile record string
                mStrOutLine = Build913String(mWaybill)

                'the line is built - write it to the file
                outfile.WriteLine(mStrOutLine)

                'unload the waybill class
                mWaybill = Nothing

                Rec = Rec + 1

            Next mLooper

            ' Close the Progress Bar form
            'ProgressBar_Form.Close()

            'close the output file
            outfile.Close()

        End If

        mDataTable = Nothing

    End Sub

    Public Function Build570String(ByVal mWaybill As Class_Waybill) As String
        Dim This570String As New StringBuilder(900)
        Dim mWorkDate As Date

        With mWaybill
            This570String = New StringBuilder

            'cc 1-6
            This570String.Append(.Serial_No)
            'cc 7-12
            This570String.Append(Field_Right(.WB_Num, 6))
            'cc 13-18
            This570String.Append(Right(.WB_Date.Year, 2))
            This570String.Append(Field_Right(.WB_Date.Month, 2))
            This570String.Append(Field_Right(.WB_Date.Day, 2))
            'cc 19-22
            This570String.Append(Field_Right(Left(.Acct_Period, 2), 2))
            This570String.Append(Field_Right(Right(Trim(.Acct_Period), 2), 2))
            'cc 23-26
            This570String.Append(Field_Right(.U_Cars, 4))
            'cc 27-30
            This570String.Append(Field_Left(.U_Car_Init, 4))
            'cc 31-36
            This570String.Append(Field_Right(.U_Car_Num, 6))
            'cc 37-38
            This570String.Append(Field_Left(Left(.TOFC_Serv_Code, 2), 2))
            'cc 39-42
            This570String.Append(Field_Right(.U_TC_Units, 4))
            'cc 43-46
            This570String.Append(Field_Right(.U_TC_Init, 4))
            'cc 47-52
            This570String.Append(Field_Right(.U_TC_Num, 6))
            'cc 53-59
            This570String.Append(Field_Right(.STCC_W49, 7))
            'cc 60-68
            This570String.Append(Field_Right(.Bill_Wght, 9))
            'cc 69-77
            This570String.Append(Field_Right(.Act_Wght, 9))
            'cc 78-86
            This570String.Append(Field_Right(.U_Rev, 9))
            'cc 87-95
            This570String.Append(Field_Right(.Tran_Chrg, 9))
            'cc 96-104
            This570String.Append(Field_Right(.Misc_Chrg, 9))
            'cc 105
            This570String.Append(Field_Right(.Intra_State_Code, 1))
            'cc 106
            This570String.Append(Field_Right(.Transit_Code, 1))
            'cc 107
            This570String.Append(Field_Right(.All_Rail_Code, 1))
            'cc 108
            This570String.Append(Field_Right(.Type_Move, 1))
            'cc 109
            This570String.Append(Field_Right(.Move_Via_Water, 1))
            'cc 110
            This570String.Append(Field_Right(.Truck_For_Rail, 1))
            'cc 111-114
            This570String.Append(Field_Right(.Shortline_Miles, 4))
            'cc 115
            This570String.Append(Field_Right(.Rebill, 1))
            'cc 116
            This570String.Append(Field_Right(.Stratum, 1))
            'cc 117
            This570String.Append(Field_Right(.Subsample, 1))
            'cc 118
            This570String.Append(Field_Right(.Int_Eq_Flg, 1))
            'cc 119
            This570String.Append(Field_Right(.Rate_Flg, 1))
            'cc 120-144
            This570String.Append(Field_Left(.Wb_Id, 25))
            'cc 145-147
            This570String.Append(Field_Right(.Report_RR, 3))
            'cc 148-152
            This570String.Append(Field_Right(.O_FSAC, 5))
            'cc 153-155
            This570String.Append(Field_Right(.ORR, 3))
            'cc 156-160
            This570String.Append(Field_Left(.JCT1, 5))
            'cc 161-163
            This570String.Append(Field_Right(.JRR1, 3))
            'cc 164-168
            This570String.Append(Field_Left(.JCT2, 5))
            'cc 169-171
            This570String.Append(Field_Right(.JRR2, 3))
            'cc 172-176
            This570String.Append(Field_Left(.JCT3, 5))
            'cc 177-179
            This570String.Append(Field_Right(.JRR3, 3))
            'cc 180-184
            This570String.Append(Field_Left(.JCT4, 5))
            'cc 185-187
            This570String.Append(Field_Right(.JRR4, 3))
            'cc 188-192
            This570String.Append(Field_Left(.JCT5, 5))
            'cc 193-195
            This570String.Append(Field_Right(.JRR5, 3))
            'cc 196-200
            This570String.Append(Field_Left(.JCT6, 5))
            'cc 201-203
            This570String.Append(Field_Right(.JRR6, 3))
            'cc 204-208
            This570String.Append(Field_Left(.JCT7, 5))
            'cc 209-211
            This570String.Append(Field_Right(.TRR, 3))
            'cc 212-216
            This570String.Append(Field_Left(.T_FSAC, 5))
            'cc 217-224
            This570String.Append(Field_Right(.Pop_Cnt, 8))
            'cc 225-230
            This570String.Append(Field_Right(.Stratum_Cnt, 6))
            'cc 231-234
            This570String.Append(Field_Left(.Car_Own_Mark, 4))
            'cc 235-238
            This570String.Append(Field_Left(.Car_Lessee_Mark, 4))
            'cc 239-241
            This570String.Append(Field_Right(.Nom_Car_Cap, 3))
            'cc 242
            This570String.Append(Field_Right(.Art_Units, 1))
            'cc 243-246
            This570String.Append(Field_Left(.Car_Typ, 4))
            'cc 247-250
            This570String.Append(Field_Left(.Mech, 4))
            'cc 251-252
            This570String.Append(Field_Left(.Lic_St, 2))
            'cc 253-258
            This570String.Append(Field_Right(.O_SPLC, 6))
            'cc 259-264
            This570String.Append(Field_Right(.T_SPLC, 6))
            'cc 265-271
            This570String.Append(Field_Right(.STCC, 7))
            'cc 272
            This570String.Append(Field_Right(.JF, 1))
            'cc 273-275
            This570String.Append(Field_Right(.Exp_Factor_Th, 3))
            'cc 276
            This570String.Append(Field_Right(.Error_Flg, 1))
            'cc 277-278
            This570String.Append(Field_Right(.STB_Car_Type, 2))
            'cc 279-280
            This570String.Append(Field_Right(.Err_Code1, 2))
            'cc 281-282
            This570String.Append(Field_Right(.Err_Code2, 2))
            'cc 283-290
            If IsDate(.Dereg_Date) Then
                mWorkDate = CDate(.Dereg_Date)
                This570String.Append(Field_Right(mWorkDate.Year, 4))
                This570String.Append(Field_Right(mWorkDate.Month, 2))
                This570String.Append(Field_Right(mWorkDate.Day, 2))
            Else
                This570String.Append(Space(8))
            End If
            'cc 291
            This570String.Append(Field_Right(.Dereg_Flg, 1))
            'cc 292-298
            This570String.Append(Field_Right(.Bill_Wght_Tons, 7))
            'cc 299-303
            This570String.Append(Field_Right(.ORR_Dist, 5))
            'cc 304-308
            This570String.Append(Field_Right(.JRR1_Dist, 5))
            'cc 309-313
            This570String.Append(Field_Right(.JRR2_Dist, 5))
            'cc 314-318
            This570String.Append(Field_Right(.JRR3_Dist, 5))
            'cc 319-323
            This570String.Append(Field_Right(.JRR4_Dist, 5))
            'cc 324-328
            This570String.Append(Field_Right(.JRR5_Dist, 5))
            'cc 329-333
            This570String.Append(Field_Right(.JRR6_Dist, 5))
            'cc 334-338
            This570String.Append(Field_Right(.TRR_Dist, 5))
            'cc 339-343
            This570String.Append(Field_Left(.Total_Dist, 5))
            'cc 344-345
            This570String.Append(Field_Left(.O_ST, 2))
            'cc 346-347
            This570String.Append(Field_Left(.T_ST, 2))
            'cc 348-350
            This570String.Append(Field_Right(.O_BEA, 3))
            'cc 351-353
            This570String.Append(Field_Right(.T_BEA, 3))
            'cc 354-358
            This570String.Append(Field_Right(.O_FIPS, 5))
            'cc 359-363
            This570String.Append(Field_Right(.T_FIPS, 5))
            'cc 364-365
            This570String.Append(Field_Right(.O_FA, 2))
            'cc 366-367
            This570String.Append(Field_Right(.T_FA, 2))
            'cc 368
            This570String.Append(Field_Right(.O_FT, 1))
            'cc 369
            This570String.Append(Field_Right(.T_FT, 1))
            'cc 370-373
            This570String.Append(Field_Right(.O_SMSA, 4))
            'cc 374-377
            This570String.Append(Field_Right(.T_SMSA, 4))
            'cc 378
            This570String.Append(Field_Right(.AL_Flg, 1))
            'cc 379
            This570String.Append(Field_Right(.AZ_Flg, 1))
            'cc 380
            This570String.Append(Field_Right(.AR_Flg, 1))
            'cc 381
            This570String.Append(Field_Right(.CA_Flg, 1))
            'cc 382
            This570String.Append(Field_Right(.CO_Flg, 1))
            'cc 383
            This570String.Append(Field_Right(.CT_Flg, 1))
            'cc 384
            This570String.Append(Field_Right(.DE_Flg, 1))
            'cc 385
            This570String.Append(Field_Right(.DC_Flg, 1))
            'cc 386
            This570String.Append(Field_Right(.FL_Flg, 1))
            'cc 387
            This570String.Append(Field_Right(.GA_Flg, 1))
            'cc 388
            This570String.Append(Field_Right(.ID_Flg, 1))
            'cc 389
            This570String.Append(Field_Right(.IL_Flg, 1))
            'cc 390
            This570String.Append(Field_Right(.IN_Flg, 1))
            'cc 391
            This570String.Append(Field_Right(.IA_Flg, 1))
            'cc 392
            This570String.Append(Field_Right(.KS_Flg, 1))
            'cc 393
            This570String.Append(Field_Right(.KY_Flg, 1))
            'cc 394
            This570String.Append(Field_Right(.LA_Flg, 1))
            'cc 395
            This570String.Append(Field_Right(.ME_Flg, 1))
            'cc 396
            This570String.Append(Field_Right(.MD_Flg, 1))
            'cc 397
            This570String.Append(Field_Right(.MA_Flg, 1))
            'cc 398
            This570String.Append(Field_Right(.MI_Flg, 1))
            'cc 399
            This570String.Append(Field_Right(.MN_Flg, 1))
            'cc 400
            This570String.Append(Field_Right(.MS_Flg, 1))
            'cc 401
            This570String.Append(Field_Right(.MO_Flg, 1))
            'cc 402
            This570String.Append(Field_Right(.MT_Flg, 1))
            'cc 403
            This570String.Append(Field_Right(.NE_Flg, 1))
            'cc 404
            This570String.Append(Field_Right(.NV_Flg, 1))
            'cc 405
            This570String.Append(Field_Right(.NH_Flg, 1))
            'cc 406
            This570String.Append(Field_Right(.NJ_Flg, 1))
            'cc 407
            This570String.Append(Field_Right(.NM_Flg, 1))
            'cc 408
            This570String.Append(Field_Right(.NY_Flg, 1))
            'cc 409
            This570String.Append(Field_Right(.NC_Flg, 1))
            'cc 410
            This570String.Append(Field_Right(.ND_Flg, 1))
            'cc 411
            This570String.Append(Field_Right(.OH_Flg, 1))
            'cc 412
            This570String.Append(Field_Right(.OK_Flg, 1))
            'cc 413
            This570String.Append(Field_Right(.OR_Flg, 1))
            'cc 414
            This570String.Append(Field_Right(.PA_Flg, 1))
            'cc 415
            This570String.Append(Field_Right(.RI_Flg, 1))
            'cc 416
            This570String.Append(Field_Right(.SC_Flg, 1))
            'cc 417
            This570String.Append(Field_Right(.SD_Flg, 1))
            'cc 418
            This570String.Append(Field_Right(.TN_Flg, 1))
            'cc 419
            This570String.Append(Field_Right(.TX_Flg, 1))
            'cc 420
            This570String.Append(Field_Right(.UT_Flg, 1))
            'cc 421
            This570String.Append(Field_Right(.VT_Flg, 1))
            'cc 422
            This570String.Append(Field_Right(.VA_Flg, 1))
            'cc 423
            This570String.Append(Field_Right(.WA_Flg, 1))
            'cc 424
            This570String.Append(Field_Right(.WV_Flg, 1))
            'cc 425
            This570String.Append(Field_Right(.WI_Flg, 1))
            'cc 426
            This570String.Append(Field_Right(.WY_Flg, 1))
            'cc 427
            This570String.Append(Field_Right(.Othr_St_Flg, 1))
            'cc 428-432
            This570String.Append(Field_Right("0", 5))

            '433 is CD_Flg in database, but has not been used since 2003.
            'Build570String = Build570String & Field_Right(.CD_Flg, 1)  'MRS 03/2010

            '434 is MX_Flg in database, but has not been used since 2003.
            'Build570String = Build570String & Field_Right(.MX_Flg, 1)  'MRS 03/2010

            'cc 433-434
            This570String.Append(Field_Right("0", 2))
            'cc 435
            This570String.Append(Field_Left(.Car_Own, 1))
            'cc 436-439
            This570String.Append(Field_Left(.O_Census_Reg, 4))
            'cc 440-443
            This570String.Append(Field_Left(.T_Census_Reg, 4))
            'cc 444-450
            This570String.Append(Field_Right(.Exp_Factor, 7))
            'cc 451-458
            This570String.Append(Field_Right(.Total_VC, 8))

            '459 is Transborder_Flg in database that has not been used since 2003.
            ' Build570String = Build570String & Field_Right(.Transborder_Flg, 1)   ' MRS 03/2010
            This570String.Append(Space(8))

            'cc 459-466
            'Build570String = Build570String & Space(8)
            'cc 467-468
            This570String.Append(Field_Left(.CS_54, 2))
            'cc 469-476
            This570String.Append(Field_Right(.RR1_VC, 8))
            'cc 477-484
            This570String.Append(Field_Right(.RR2_VC, 8))
            'cc 485-492
            This570String.Append(Field_Right(.RR3_VC, 8))
            'cc 493-499
            This570String.Append(Field_Right(.RR4_VC, 8))
            'cc 500-506
            This570String.Append(Field_Right(.RR5_VC, 8))
            'cc 507-513
            This570String.Append(Field_Right(.RR6_VC, 8))
            'cc 514-520
            This570String.Append(Field_Right(.RR7_VC, 8))
            'cc 521-527
            This570String.Append(Field_Right(.RR8_VC, 8))
            'cc 528-539
            This570String.Append(Field_Left(.Int_Harm_Code, 12))
            'cc 540-543
            This570String.Append(Field_Left(.Indus_Class, 4))
            'cc 544-547
            This570String.Append(Field_Left(.Inter_Sic, 4))
            'cc 548-550
            This570String.Append(Field_Left(.Dom_Canada, 3))
            'cc 551-554
            This570String.Append(Field_Left(.O_FS_Type, 4))
            'cc 555-558
            This570String.Append(Field_Left(.T_FS_Type, 4))
            'cc 559
            This570String.Append(Field_Left(.O_Customs_Flg, 1))
            'cc 560
            This570String.Append(Field_Left(.T_Customs_Flg, 1))
            'cc 561
            This570String.Append(Field_Left(.O_Grain_Flg, 1))
            'cc 562
            This570String.Append(Field_Left(.T_Grain_Flg, 1))
            'cc 563
            This570String.Append(Field_Left(.O_Ramp_Code, 1))
            'cc 564
            This570String.Append(Field_Left(.T_Ramp_Code, 1))
            'cc 565
            This570String.Append(Field_Left(.O_IM_Flg, 1))
            'cc 566
            This570String.Append(Field_Left(.T_IM_Flg, 1))
            'cc 567-570
            This570String.Append(Field_Left(.TOFC_Unit_Type, 4))
        End With

        Build570String = This570String.ToString

    End Function

    Public Function Build913String(ByVal mWaybill As Class_Waybill) As String
        Dim This913String As New StringBuilder(913)

        'Get the values from the mWaybill

        Dim mWorkDate As Date

        With mWaybill
            This913String.Append(Field_Right(.Serial_No, 6))
            This913String.Append(Field_Right(.WB_Num, 6))
            This913String.Append(Field_Right(.WB_Date.Month.ToString, 2))
            This913String.Append(Field_Right(.WB_Date.Day.ToString, 2))
            This913String.Append(Field_Right(.WB_Date.Year.ToString, 4))
            This913String.Append(Field_Right(.Acct_Period, 6))
            This913String.Append(Field_Right(.U_Cars, 4))
            This913String.Append(Field_Left(.U_Car_Init, 4))
            This913String.Append(Field_Right(.U_Car_Num, 6))
            This913String.Append(Field_Left(.TOFC_Serv_Code, 3))
            This913String.Append(Field_Right(.U_TC_Units, 4))
            This913String.Append(Field_Left(.U_TC_Init, 4))
            This913String.Append(Field_Right(.U_TC_Num, 6))
            This913String.Append(Field_Right(.STCC_W49, 7))
            This913String.Append(Field_Right(.Bill_Wght, 9))
            This913String.Append(Field_Right(.Act_Wght, 9))
            This913String.Append(Field_Right(.U_Rev, 9))
            This913String.Append(Field_Right(.Tran_Chrg, 9))
            This913String.Append(Field_Right(.Misc_Chrg, 9))
            This913String.Append(Field_Right(.Intra_State_Code, 1))
            This913String.Append(Field_Right(.Transit_Code, 1))
            This913String.Append(Field_Right(.All_Rail_Code, 1))
            This913String.Append(Field_Right(.Type_Move, 1))
            This913String.Append(Field_Right(.Move_Via_Water, 1))
            This913String.Append(Field_Right(.Truck_For_Rail, 1))
            This913String.Append(Field_Right(.Shortline_Miles, 4))
            This913String.Append(Field_Right(.Rebill, 1))
            This913String.Append(Field_Right(.Stratum, 1))
            This913String.Append(Field_Right(.Subsample, 1))
            This913String.Append(Field_Right(.Int_Eq_Flg, 1))
            This913String.Append(Field_Right(.Rate_Flg, 1))
            This913String.Append(Field_Left(.Wb_Id, 25))
            This913String.Append(Field_Right(.Report_RR, 3))
            This913String.Append(Field_Right(.O_FSAC, 5))
            This913String.Append(Field_Right(.ORR, 3))
            This913String.Append(Field_Left(.JCT1, 5))
            This913String.Append(Field_Right(.JRR1, 3))
            This913String.Append(Field_Left(.JCT2, 5))
            This913String.Append(Field_Right(.JRR2, 3))
            This913String.Append(Field_Left(.JCT3, 5))
            This913String.Append(Field_Right(.JRR3, 3))
            This913String.Append(Field_Left(.JCT4, 5))
            This913String.Append(Field_Right(.JRR4, 3))
            This913String.Append(Field_Left(.JCT5, 5))
            This913String.Append(Field_Right(.JRR5, 3))
            This913String.Append(Field_Left(.JCT6, 5))
            This913String.Append(Field_Right(.JRR6, 3))
            This913String.Append(Field_Left(.JCT7, 5))
            This913String.Append(Field_Right(.TRR, 3))
            This913String.Append(Field_Right(.T_FSAC, 5))
            This913String.Append(Field_Right(.Pop_Cnt, 8))
            This913String.Append(Field_Right(.Stratum_Cnt, 6))
            This913String.Append(Field_Right(.Report_Period, 1))
            This913String.Append(Field_Left(.Car_Own_Mark, 4))
            This913String.Append(Field_Left(.Car_Lessee_Mark, 4))
            This913String.Append(Field_Right(.Car_Cap, 5))
            This913String.Append(Field_Right(.Nom_Car_Cap, 3))
            This913String.Append(Field_Right(.Tare, 4))
            This913String.Append(Field_Right(.Outside_L, 5))
            This913String.Append(Field_Right(.Outside_W, 4))
            This913String.Append(Field_Right(.Outside_H, 4))
            This913String.Append(Field_Right(.Ex_Outside_H, 4))
            This913String.Append(Field_Right(.Type_Wheel, 1))
            This913String.Append(Field_Right(.No_Axles, 1))
            This913String.Append(Field_Right(.Draft_Gear, 2))
            This913String.Append(Field_Right(.Art_Units, 1))
            This913String.Append(Field_Right(.Pool_Code, 7))
            This913String.Append(Field_Left(.Car_Typ, 4))
            This913String.Append(Field_Left(.Mech, 4))
            This913String.Append(Field_Left(.Lic_St, 2))
            This913String.Append(Field_Right(.Mx_Wght_Rail, 3))
            This913String.Append(Field_Right(.O_SPLC, 6))
            This913String.Append(Field_Right(.T_SPLC, 6))
            This913String.Append(Field_Right(.STCC, 7))
            This913String.Append(Field_Left(.ORR_Alpha, 4))
            This913String.Append(Field_Left(.JRR1_Alpha, 4))
            This913String.Append(Field_Left(.JRR2_Alpha, 4))
            This913String.Append(Field_Left(.JRR3_Alpha, 4))
            This913String.Append(Field_Left(.JRR4_Alpha, 4))
            This913String.Append(Field_Left(.JRR5_Alpha, 4))
            This913String.Append(Field_Left(.JRR6_Alpha, 4))
            This913String.Append(Field_Left(.TRR_Alpha, 4))
            This913String.Append(Field_Right(.JF, 1))
            This913String.Append(Field_Right(.Exp_Factor_Th, 3))
            This913String.Append(Field_Right(.Error_Flg, 1))
            This913String.Append(Field_Right(.STB_Car_Type, 2))
            This913String.Append(Field_Right(.Err_Code1, 2))
            This913String.Append(Field_Right(.Err_Code2, 2))
            This913String.Append(Field_Right(.Err_Code3, 2))
            This913String.Append(Field_Right(.Car_Own, 1))
            This913String.Append(Field_Left(.TOFC_Unit_Type, 4))
            If IsDate(.Dereg_Date) Then
                mWorkDate = .Dereg_Date
                This913String.Append(Field_Right(mWorkDate.Year.ToString, 4))
                This913String.Append(Field_Right(mWorkDate.Month.ToString, 2))
                This913String.Append(Field_Right(mWorkDate.Day.ToString, 2))
            Else
                This913String.Append(Space(8))
            End If
            This913String.Append(Field_Right(.Dereg_Flg, 1))
            This913String.Append(Field_Right(.Service_Type, 1))
            This913String.Append(Field_Right(.Cars, 6))
            This913String.Append(Field_Right(.Bill_Wght_Tons, 7))
            This913String.Append(Field_Right(.Tons, 8))
            This913String.Append(Field_Right(.TC_Units, 6))
            This913String.Append(Field_Right(.Total_Rev, 10))
            This913String.Append(Field_Right(.ORR_Rev, 10))
            This913String.Append(Field_Right(.JRR1_Rev, 10))
            This913String.Append(Field_Right(.JRR2_Rev, 10))
            This913String.Append(Field_Right(.JRR3_Rev, 10))
            This913String.Append(Field_Right(.JRR4_Rev, 10))
            This913String.Append(Field_Right(.JRR5_Rev, 10))
            This913String.Append(Field_Right(.JRR6_Rev, 10))
            This913String.Append(Field_Right(.TRR_Rev, 10))
            This913String.Append(Field_Right(.ORR_Dist, 5))
            This913String.Append(Field_Right(.JRR1_Dist, 5))
            This913String.Append(Field_Right(.JRR2_Dist, 5))
            This913String.Append(Field_Right(.JRR3_Dist, 5))
            This913String.Append(Field_Right(.JRR4_Dist, 5))
            This913String.Append(Field_Right(.JRR5_Dist, 5))
            This913String.Append(Field_Right(.JRR6_Dist, 5))
            This913String.Append(Field_Right(.TRR_Dist, 5))
            This913String.Append(Field_Right(.Total_Dist, 5))
            This913String.Append(Field_Left(.O_ST, 2))
            This913String.Append(Field_Left(.JCT1_ST, 2))
            This913String.Append(Field_Left(.JCT2_ST, 2))
            This913String.Append(Field_Left(.JCT3_ST, 2))
            This913String.Append(Field_Left(.JCT4_ST, 2))
            This913String.Append(Field_Left(.JCT5_ST, 2))
            This913String.Append(Field_Left(.JCT6_ST, 2))
            This913String.Append(Field_Left(.JCT7_ST, 2))
            This913String.Append(Field_Left(.T_ST, 2))
            This913String.Append(Field_Right(.O_BEA, 3))
            This913String.Append(Field_Right(.T_BEA, 3))
            This913String.Append(Field_Right(.O_FIPS, 5))
            This913String.Append(Field_Right(.T_FIPS, 5))
            This913String.Append(Field_Right(.O_FA, 2))
            This913String.Append(Field_Right(.T_FA, 2))
            This913String.Append(Field_Right(.O_FT, 1))
            This913String.Append(Field_Right(.T_FT, 1))
            This913String.Append(Field_Right(.O_SMSA, 4))
            This913String.Append(Field_Right(.T_SMSA, 4))
            This913String.Append(Field_Right(.ONET, 5))
            This913String.Append(Field_Right(.NET1, 5))
            This913String.Append(Field_Right(.NET2, 5))
            This913String.Append(Field_Right(.NET3, 5))
            This913String.Append(Field_Right(.NET4, 5))
            This913String.Append(Field_Right(.NET5, 5))
            This913String.Append(Field_Right(.NET6, 5))
            This913String.Append(Field_Right(.NET7, 5))
            This913String.Append(Field_Right(.TNET, 5))
            This913String.Append(Field_Right(.AL_Flg, 1))
            This913String.Append(Field_Right(.AZ_Flg, 1))
            This913String.Append(Field_Right(.AR_Flg, 1))
            This913String.Append(Field_Right(.CA_Flg, 1))
            This913String.Append(Field_Right(.CO_Flg, 1))
            This913String.Append(Field_Right(.CT_Flg, 1))
            This913String.Append(Field_Right(.DE_Flg, 1))
            This913String.Append(Field_Right(.DC_Flg, 1))
            This913String.Append(Field_Right(.FL_Flg, 1))
            This913String.Append(Field_Right(.GA_Flg, 1))
            This913String.Append(Field_Right(.ID_Flg, 1))
            This913String.Append(Field_Right(.IL_Flg, 1))
            This913String.Append(Field_Right(.IN_Flg, 1))
            This913String.Append(Field_Right(.IA_Flg, 1))
            This913String.Append(Field_Right(.KS_Flg, 1))
            This913String.Append(Field_Right(.KY_Flg, 1))
            This913String.Append(Field_Right(.LA_Flg, 1))
            This913String.Append(Field_Right(.ME_Flg, 1))
            This913String.Append(Field_Right(.MD_Flg, 1))
            This913String.Append(Field_Right(.MA_Flg, 1))
            This913String.Append(Field_Right(.MI_Flg, 1))
            This913String.Append(Field_Right(.MN_Flg, 1))
            This913String.Append(Field_Right(.MS_Flg, 1))
            This913String.Append(Field_Right(.MO_Flg, 1))
            This913String.Append(Field_Right(.MT_Flg, 1))
            This913String.Append(Field_Right(.NE_Flg, 1))
            This913String.Append(Field_Right(.NV_Flg, 1))
            This913String.Append(Field_Right(.NH_Flg, 1))
            This913String.Append(Field_Right(.NJ_Flg, 1))
            This913String.Append(Field_Right(.NM_Flg, 1))
            This913String.Append(Field_Right(.NY_Flg, 1))
            This913String.Append(Field_Right(.NC_Flg, 1))
            This913String.Append(Field_Right(.ND_Flg, 1))
            This913String.Append(Field_Right(.OH_Flg, 1))
            This913String.Append(Field_Right(.OK_Flg, 1))
            This913String.Append(Field_Right(.OR_Flg, 1))
            This913String.Append(Field_Right(.PA_Flg, 1))
            This913String.Append(Field_Right(.RI_Flg, 1))
            This913String.Append(Field_Right(.SC_Flg, 1))
            This913String.Append(Field_Right(.SD_Flg, 1))
            This913String.Append(Field_Right(.TN_Flg, 1))
            This913String.Append(Field_Right(.TX_Flg, 1))
            This913String.Append(Field_Right(.UT_Flg, 1))
            This913String.Append(Field_Right(.VT_Flg, 1))
            This913String.Append(Field_Right(.VA_Flg, 1))
            This913String.Append(Field_Right(.WA_Flg, 1))
            This913String.Append(Field_Right(.WV_Flg, 1))
            This913String.Append(Field_Right(.WI_Flg, 1))
            This913String.Append(Field_Right(.WY_Flg, 1))
            This913String.Append(Field_Right(.CD_Flg, 1))
            This913String.Append(Field_Right(.MX_Flg, 1))
            This913String.Append(Field_Right(.Othr_St_Flg, 1))
            This913String.Append(Field_Left(.Int_Harm_Code, 12))
            This913String.Append(Field_Left(.Indus_Class, 4))
            This913String.Append(Field_Left(.Inter_Sic, 4))
            This913String.Append(Field_Left(.Dom_Canada, 3))
            This913String.Append(Field_Right(.CS_54, 2))
            This913String.Append(Field_Left(.O_FS_Type, 4))
            This913String.Append(Field_Left(.T_FS_Type, 4))
            This913String.Append(Field_Left(.O_FS_RateZip, 9))
            This913String.Append(Field_Left(.T_FS_RateZip, 9))
            This913String.Append(Field_Left(.O_Rate_SPLC, 9))
            This913String.Append(Field_Left(.T_Rate_SPLC, 9))
            This913String.Append(Field_Left(.O_SWLimit_SPLC, 9))
            This913String.Append(Field_Left(.T_SWLimit_SPLC, 9))
            This913String.Append(Field_Left(.O_Customs_Flg, 1))
            This913String.Append(Field_Left(.T_Customs_Flg, 1))
            This913String.Append(Field_Left(.O_Grain_Flg, 1))
            This913String.Append(Field_Left(.T_Grain_Flg, 1))
            This913String.Append(Field_Left(.O_Ramp_Code, 1))
            This913String.Append(Field_Left(.T_Ramp_Code, 1))
            This913String.Append(Field_Left(.O_IM_Flg, 1))
            This913String.Append(Field_Left(.T_IM_Flg, 1))
            This913String.Append(Field_Left(.Transborder_Flg, 1))
            This913String.Append(Field_Left(.ORR_Cntry, 2))
            This913String.Append(Field_Left(.JRR1_Cntry, 2))
            This913String.Append(Field_Left(.JRR2_Cntry, 2))
            This913String.Append(Field_Left(.JRR3_Cntry, 2))
            This913String.Append(Field_Left(.JRR4_Cntry, 2))
            This913String.Append(Field_Left(.JRR5_Cntry, 2))
            This913String.Append(Field_Left(.JRR6_Cntry, 2))
            This913String.Append(Field_Left(.TRR_Cntry, 2))
            This913String.Append(Field_Right(.U_Fuel_SurChrg, 9))
            This913String.Append(Space(13))    'unused space in record
            This913String.Append(Field_Left(.O_Census_Reg, 4))
            This913String.Append(Field_Left(.T_Census_Reg, 4))
            This913String.Append(Field_Right(.Exp_Factor, 7))
            This913String.Append(Field_Right(.Total_VC, 8))
            This913String.Append(Field_Right(.RR1_VC, 8))
            This913String.Append(Field_Right(.RR2_VC, 8))
            This913String.Append(Field_Right(.RR3_VC, 8))
            This913String.Append(Field_Right(.RR4_VC, 7))
            This913String.Append(Field_Right(.RR5_VC, 7))
            This913String.Append(Field_Right(.RR6_VC, 7))
            This913String.Append(Field_Right(.RR7_VC, 7))
            This913String.Append(Field_Right(.RR8_VC, 7))
            This913String.Append(Field_Right(.Tracking_No, 13))

        End With

        Build913String = This913String.ToString

    End Function

End Module
