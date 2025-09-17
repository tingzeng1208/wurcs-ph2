Imports System.Data.SqlClient
Imports System.Text
Public Class frm_Productivity

    Private Sub frm_Productivity_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim mWB_Years As DataTable
        Dim mLooper As Integer

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        ' Open/Check the SQL connection
        Gbl_Controls_Database_Name = "URCS_Controls"
        OpenADOConnection(Gbl_Controls_Database_Name)

        gbl_SQLConnection = New SqlConnection
        gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(Gbl_Controls_Database_Name)
        gbl_SQLConnection.Open()

        mWB_Years = New DataTable
        mWB_Years = Get_Waybill_Years_Table()

        For mLooper = 0 To mWB_Years.Rows.Count - 1
            cmb_URCSYear.Items.Add(mWB_Years.Rows(mLooper)("wb_year").ToString)
        Next

    End Sub

    Private Sub btn_Return_To_MainMenu_Click(sender As System.Object, e As System.EventArgs) Handles btn_Return_To_MainMenu.Click
        ' Open the post processing Menu Form
        Dim frmNew As New frm_Post_Processing_Menu
        frmNew.Show()
        ' Close this Form
        Me.Close()
    End Sub

    Private Sub cmb_URCSYear_SelectedValueChanged(sender As System.Object, e As System.EventArgs) Handles cmb_URCSYear.SelectedValueChanged
        Dim mWayRRR As DataTable
        Dim mTrans As DataTable
        Dim mStrSQL As String
        Dim mThisRegion As Integer
        Dim mEastRegions As Integer
        Dim mWestRegions As Integer
        Dim mEastFactors As Decimal
        Dim mWestFactors As Decimal
        Dim mEastRRICC_IDs() As String
        Dim mWestRRICC_IDs() As String

        ' We have to set the default trailer usage values that changed thru the years
        ' this crap is the legacy method that didn't get updated for each year and was completely
        ' forgotten about after 2002

        ' Set defaults first
        txt_Trailer_Usage_East.Text = 1.712
        txt_Trailer_Usage_West.Text = 1.828

        If CInt(cmb_URCSYear.Text) > 1994 Then
            txt_Trailer_Usage_East.Text = 1.7768
        End If

        If CInt(cmb_URCSYear.Text) > 2002 Then
            txt_Trailer_Usage_East.Text = 3.868661
            txt_Trailer_Usage_West.Text = 5.054878
        End If

        ' OK, all of that was prior to SQL, so now we're going to do it the right way
        ' open the connection to the location of the wayrrr table

        Gbl_Controls_Database_Name = "URCS_Controls"
        OpenADOConnection(Gbl_Controls_Database_Name)

        gbl_SQLConnection = New SqlConnection
        gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(Gbl_Controls_Database_Name)
        gbl_SQLConnection.Open()

        mWayRRR = New DataTable
        mWayRRR = Get_WayRRR_Table()

        ' initialize the values we're going to use
        mEastRegions = 0
        mWestRegions = 0
        mEastFactors = 0
        mWestFactors = 0
        mEastRRICC_IDs = Nothing
        mWestRRICC_IDs = Nothing

        For i = 0 To mWayRRR.Rows.Count - 1
            If mWayRRR.Rows(i)("region_id") = 4 Then
                If mEastRRICC_IDs Is Nothing Then
                    ReDim Preserve mEastRRICC_IDs(1)
                Else
                    ReDim Preserve mEastRRICC_IDs(UBound(mEastRRICC_IDs, 1) + 1)
                End If
                mEastRRICC_IDs(UBound(mEastRRICC_IDs, 1)) = mWayRRR.Rows(i)("rricc_id").ToString
            Else
                If mWestRRICC_IDs Is Nothing Then
                    ReDim Preserve mWestRRICC_IDs(1)
                Else
                    ReDim Preserve mWestRRICC_IDs(UBound(mWestRRICC_IDs, 1) + 1)
                End If
                mWestRRICC_IDs(UBound(mWestRRICC_IDs, 1)) = mWayRRR.Rows(i)("rricc_id").ToString
            End If
        Next

        ' we need to find out where the Trans database is
        Gbl_Trans_DatabaseName = Get_Database_Name_From_SQL("1", "TRANS")
        Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "TRANS")
        OpenADOConnection(Gbl_Trans_DatabaseName)

        mTrans = New DataTable

        ' Get the values for East and West railroads for the processing year
        For mThisRegion = 1 To UBound(mEastRRICC_IDs, 1)
            mStrSQL = "SELECT * FROM " & Gbl_Trans_TableName &
                " WHERE Year = " & cmb_URCSYear.Text &
                " AND RRICC = " & mEastRRICC_IDs(mThisRegion).ToString &
                " AND SCH = 37 AND Line = 134"
            mTrans = Get_SQL_DataTable(gbl_SQLConnection, mStrSQL)
            mEastFactors = mEastFactors + mTrans.Rows(0)("C1") / 100
        Next

        For mThisRegion = 1 To UBound(mWestRRICC_IDs, 1)
            mStrSQL = "SELECT * FROM " & Gbl_Trans_TableName &
                " WHERE Year = " & cmb_URCSYear.Text &
                " AND RRICC = " & mWestRRICC_IDs(mThisRegion).ToString &
                " AND SCH = 37 AND Line = 134"
            mTrans = Get_SQL_DataTable(gbl_SQLConnection, mStrSQL)
            mWestFactors = mWestFactors + mTrans.Rows(0)("C1") / 100
        Next

        ' average them and store them to the text boxes.
        txt_Trailer_Usage_East.Text = Math.Round(mEastFactors / UBound(mEastRRICC_IDs, 1), 6, MidpointRounding.AwayFromZero)
        txt_Trailer_Usage_West.Text = Math.Round(mWestFactors / UBound(mWestRRICC_IDs, 1), 6, MidpointRounding.AwayFromZero)

    End Sub

    Private Sub btn_Output_Dir_Entry_Click(sender As System.Object, e As System.EventArgs) Handles btn_Output_Dir_Entry.Click
        Dim fd As New FolderBrowserDialog

        If cmb_URCSYear.Text = "" Then
            MsgBox("You must first select a year to process!", MsgBoxStyle.OkOnly, "ERROR!")
        Else

            fd.Description = "Select the location in which you want the output file placed."

            If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txt_Output_FilePath.Text = fd.SelectedPath & "\WBSTB"
                Select Case CInt(cmb_URCSYear.Text)
                    Case Is = 2000
                        txt_Output_FilePath.Text = txt_Output_FilePath.Text & "00.DAT"
                    Case Is > 2000
                        If (CInt(cmb_URCSYear.Text) - 2000) < 10 Then
                            txt_Output_FilePath.Text = txt_Output_FilePath.Text & "0"
                        End If
                        txt_Output_FilePath.Text = txt_Output_FilePath.Text &
                            (CInt(cmb_URCSYear.Text) - 2000).ToString & ".DAT"
                    Case Else
                        txt_Output_FilePath.Text = txt_Output_FilePath.Text &
                            (CInt(cmb_URCSYear.Text) - 1900).ToString & ".DAT"
                End Select
            End If
        End If
    End Sub

    Private Sub btn_Execute_Click(sender As System.Object, e As System.EventArgs) Handles btn_Execute.Click
        Dim outfile As StreamWriter
        Dim mWaybills As DataTable, mThisWaybill As DataTable, mProductivityTable As DataTable
        Dim mStrSQL As String, mWorkstr As StringBuilder, mSQLCmd As SqlCommand
        Dim bBadWaybill As Boolean

        'Variables from Waybill
        Dim mMiscCharge As Integer
        Dim mTranCharge As Integer
        Dim mDistanceArray(10) As Integer
        Dim mRRIDArray(8) As Integer
        Dim mRRIDLooper As Integer
        Dim mJCTArray(7) As String
        Dim mRevArray(8) As Integer
        Dim mRevLooper As Integer
        Dim mJF As Integer
        Dim mTons As Integer
        Dim mExp As Integer
        Dim mCarTypeStratum As Integer
        Dim mNumberOfCars As Integer
        Dim mExpandedCars As Integer
        Dim mOwner As Integer
        Dim mLengthOfHaulStratum As Integer
        Dim mLadingStratum As Integer
        Dim mCarsStratum As Integer
        Dim mTonMiles As Long
        Dim mReport_RR As Integer
        Dim mTotal_rev As Long
        Dim mTotal_VC As Integer
        Dim mTons_Per_Car As Integer
        Dim mNumberOfWaybills As Integer
        Dim mDistLooper As Integer

        'Items for Legacy program support
        Dim CommGroup As Integer = 1        'Unused commodity group
        Dim Ownership As Integer = 1        'Unused ownership
        Dim VarCost As Double = 0           'Unused Variable Cost

        ' Items for debugging/comparisons
        Dim mOutstring As StringBuilder
        Dim sbLogfile As StreamWriter
        Dim mWriteCSV As Boolean = True
        Dim mWrite3s As Boolean = False

        'Perform error checking for the form controls
        If cmb_URCSYear.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo SkipIt
        End If

        If txt_Output_FilePath.Text = "" Then
            MsgBox("You must select a filename to use or create.", vbOKOnly)
            GoTo SkipIt
        End If

        sbLogfile = New StreamWriter(Replace(txt_Output_FilePath.Text, "DAT", "CSV"), False)

        ' open the connection to the waybill table
        Gbl_Waybill_Database_Name = Get_Database_Name_From_SQL(cmb_URCSYear.Text, "MASKED")
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(cmb_URCSYear.Text, "Masked")
        Gbl_Unmasked_Rev_TableName = Get_Table_Name_From_SQL(cmb_URCSYear.Text, "Unmasked_Rev")
        Gbl_Productivity_TableName = Get_Table_Name_From_SQL("1", "Productivity")

        'Verify that the table exists
        If TableExist("Waybills", Gbl_Productivity_TableName) = False Then
            'We have to create it
            Create_Productivity_Table()
        Else
            ' We need to delete the records for this processing year
            mSQLCmd = New SqlCommand
            mSQLCmd.Connection = gbl_SQLConnection
            mSQLCmd.CommandText = "DELETE FROM " & Gbl_Productivity_TableName & " WHERE Year = " & cmb_URCSYear.Text
            mSQLCmd.ExecuteNonQuery()
        End If

        txt_StatusBox.Text = "Fetching Waybill Data - Please wait..."
        Refresh()

        ' open the connection to the waybills database
        OpenSQLConnection(Gbl_Waybill_Database_Name)

        mWaybills = New DataTable
        mStrSQL = "SELECT " & Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_No," &
            Gbl_Unmasked_Rev_TableName & ".ORR_Unmasked_Rev, " &
            Gbl_Unmasked_Rev_TableName & ".JRR1_Unmasked_Rev, " &
            Gbl_Unmasked_Rev_TableName & ".JRR2_Unmasked_Rev, " &
            Gbl_Unmasked_Rev_TableName & ".JRR3_Unmasked_Rev, " &
            Gbl_Unmasked_Rev_TableName & ".JRR4_Unmasked_Rev, " &
            Gbl_Unmasked_Rev_TableName & ".JRR5_Unmasked_Rev, " &
            Gbl_Unmasked_Rev_TableName & ".JRR6_Unmasked_Rev, " &
            Gbl_Unmasked_Rev_TableName & ".TRR_Unmasked_Rev, " &
            Gbl_Unmasked_Rev_TableName & ".Total_Unmasked_Rev, " &
            Gbl_Masked_TableName & ".Total_VC, " &
            Gbl_Masked_TableName & ".O_SPLC, " &
            Gbl_Masked_TableName & ".Total_Dist, " &
            Gbl_Masked_TableName & ".Tons, " &
            Gbl_Masked_TableName & ".STB_Car_Typ, " &
            Gbl_Masked_TableName & ".Car_Typ, " &
            Gbl_Masked_TableName & ".Cars, " &
            Gbl_Masked_TableName & ".U_Cars, " &
            Gbl_Masked_TableName & ".ORR_Dist, JRR1_Dist, JRR2_Dist, JRR3_Dist, JRR4_Dist, JRR5_Dist, JRR6_Dist, TRR_Dist," &
            Gbl_Masked_TableName & ".Exp_Factor_Th, Report_RR, " &
            Gbl_Masked_TableName & ".Total_Dist FROM " & Gbl_Unmasked_Rev_TableName &
            " INNER JOIN " & Gbl_Masked_TableName & " ON " &
            Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_No = " & Gbl_Masked_TableName & ".Serial_No "

        If mWrite3s = True Then
            mStrSQL = mStrSQL & " where right(" & Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_No,1) = 3"
        End If

        mWaybills = Get_SQL_DataTable(gbl_SQLConnection, mStrSQL)

        If mWriteCSV = True Then
            ' Write the header
            mOutstring = New StringBuilder
            mOutstring.Append("Serial_No, Year, LengthOfHaulStratum, CarTypeStratum, LadingStratum, CarsStratum, Total_rev, TonMiles, Total_VC, O_SPLC, STB_Car_Typ, NumWaybills, Rejected")
            sbLogfile.WriteLine(mOutstring.ToString)
        End If

        With mWaybills
            For mLooper = 0 To .Rows.Count - 1

                ' Ensure the environment is initialized
                mMiscCharge = 0
                mTranCharge = 0
                For mDistLooper = 1 To 10
                    mDistanceArray(mDistLooper) = 0
                Next
                mJF = 0
                For mRRIDLooper = 1 To 8
                    mRRIDArray(mRRIDLooper) = 0
                Next
                For mJCTLooper = 1 To 7
                    mJCTArray(mJCTLooper) = ""
                Next
                For mRevLooper = 1 To 8
                    mRevArray(mRevLooper) = 0
                Next
                mTons = 0
                mNumberOfCars = 0
                mOwner = 0
                mCarTypeStratum = 0
                mExpandedCars = 0
                mLengthOfHaulStratum = 0
                mLadingStratum = 0
                mCarsStratum = 0
                mTonMiles = 0
                mReport_RR = 0
                mTotal_rev = 0
                mTotal_VC = 0
                mTons_Per_Car = 0

                If mLooper Mod 100 = 0 Then
                    txt_StatusBox.Text = "Examining/Processing record " & mLooper.ToString & " of " & (.Rows.Count + 1).ToString & "..."
                    Refresh()
                    Application.DoEvents()
                End If

                bBadWaybill = False

                ' If the STB_Car_Typ is zero, then bypass this record
                If .Rows(mLooper)("STB_Car_Typ") = "00" Then
                    bBadWaybill = True
                End If

                ' Get unmasked expanded revenue for each segment

                ' First we load the ORR_Unmasked_Rev to the Revenue array position 1
                mRevArray(0) = .Rows(mLooper)("ORR_Unmasked_Rev")
                ' Followed by the rest, compacted
                For mRevLooper = 1 To 6
                    mRevArray(mRevLooper) = .Rows(mLooper)("JRR" & mRevLooper.ToString & "_Unmasked_Rev")
                Next
                ' Find where the first occurance of 0 is in the array and place the TRR_Unmasked_Rev there
                mRevArray(Array1DFindFirst(mRevArray, 0)) = .Rows(mLooper)("TRR_Unmasked_Rev")

                ' Get distance for each segment.  The are in tenths of miles, and need to be converted to whole miles
                ' This data is compacted in the array as well
                ' Total distance gets loaded into position 0
                mDistanceArray(0) = Math.Round(.Rows(mLooper)("Total_Dist") / 10, 0, MidpointRounding.AwayFromZero)
                ' Orr Distance is loaded into position 1
                mDistanceArray(1) = Math.Round(.Rows(mLooper)("ORR_Dist") / 10, 0, MidpointRounding.AwayFromZero)
                For mDistLooper = 1 To 6
                    mDistanceArray(mDistLooper + 1) =
                Math.Round(.Rows(mLooper)("JRR" & mDistLooper.ToString & "_Dist") / 10, 0, MidpointRounding.AwayFromZero)
                Next
                mDistanceArray(8) = Math.Round(.Rows(mLooper)("TRR_Dist") / 10, 0, MidpointRounding.AwayFromZero)

                ' Get expanded tons
                If Not IsDBNull(.Rows(mLooper)("Tons")) Then
                    mTons = .Rows(mLooper)("Tons")
                Else
                    mTons = 0
                End If

                ' Get the Exp_Factor_Th value
                mExp = .Rows(mLooper)("Exp_Factor_Th")

                ' Get the CarTypeStratum
                mCarTypeStratum = Get_FtCar(.Rows(mLooper)("STB_Car_typ"), .Rows(mLooper)("Car_typ"))

                ' Get the number of cars
                mNumberOfCars = .Rows(mLooper)("U_Cars")

                ' Get the Length of Haul stratum
                mLengthOfHaulStratum = LengthOfHaul(mDistanceArray(0))

                ' Get the Lading stratum
                mLadingStratum = Lading(mTons, .Rows(mLooper)("Cars")) 'Expanded number of cars

                ' Calculate the TonMiles
                mTonMiles = mTons * Math.Round(.Rows(mLooper)("total_dist") / 10, 0, MidpointRounding.AwayFromZero)
                mTonMiles = Math.Ceiling(mTonMiles / 1000) * 1000

                'Get the reporting RR
                mReport_RR = .Rows(mLooper)("Report_RR")

                'get total_rev
                mTotal_rev = .Rows(mLooper)("Total_Unmasked_Rev")

                'Get Cars Stratum
                mCarsStratum = Get_Car_Stratum(.Rows(mLooper)("U_Cars")) 'Unexpanded number of cars

                mTotal_VC = .Rows(mLooper)("total_vc")

                'Check to see if we have what we need from this waybill to process
                bBadWaybill = False
                If mLengthOfHaulStratum = 0 Then bBadWaybill = True
                If mCarsStratum = 0 Then bBadWaybill = True
                If mLadingStratum = 0 Then bBadWaybill = True
                If mCarTypeStratum = 0 Then bBadWaybill = True
                If mTotal_VC = 0 Then bBadWaybill = True            ' Was this waybill was costed?

                ' do not use if this move originated in canada or mexico
                Select Case .Rows(mLooper)("o_splc")
                    Case < 110000
                        bBadWaybill = True
                    Case > 910000
                        bBadWaybill = True
                End Select

                If bBadWaybill = False Then
                    ' Determine if record exists in SQL
                    If Count_Productivity_Records(cmb_URCSYear.Text,
                            mLengthOfHaulStratum,
                            mCarTypeStratum,
                            mLadingStratum,
                            mCarsStratum) = 0 Then
                        'Create it
                        Insert_Productivity_Record(cmb_URCSYear.Text,
                            mLengthOfHaulStratum,
                            mCarTypeStratum,
                            mLadingStratum,
                            mCarsStratum,
                            mTotal_rev,
                            mTonMiles)
                        mNumberOfWaybills = 1
                    Else
                        'Get the values already in the record
                        mThisWaybill = New DataTable
                        mThisWaybill = Get_Productivity_Data_Record(cmb_URCSYear.Text,
                                                            mLengthOfHaulStratum,
                                                            mCarTypeStratum,
                                                            mLadingStratum,
                                                            mCarsStratum)
                        mTotal_rev = mTotal_rev + mThisWaybill.Rows(0)("Revenue")
                        mTonMiles = mTonMiles + mThisWaybill.Rows(0)("Ton_Miles")
                        mNumberOfWaybills = mThisWaybill.Rows(0)("Waybill_Records") + 1
                        'Update them
                        Update_Productivity_Record(cmb_URCSYear.Text,
                                                mLengthOfHaulStratum,
                                                mCarTypeStratum,
                                                mLadingStratum,
                                                mCarsStratum,
                                                mTotal_rev,
                                                mTonMiles,
                                                mNumberOfWaybills)
                    End If

                    If mWriteCSV = True Then
                        mOutstring = New StringBuilder
                        mOutstring.Append(mWaybills.Rows(mLooper)("Unmasked_Serial_No").ToString & ",")
                        mOutstring.Append(cmb_URCSYear.Text & ",")
                        mOutstring.Append(mLengthOfHaulStratum.ToString & ",")
                        mOutstring.Append(mCarTypeStratum.ToString & ",")
                        mOutstring.Append(mLadingStratum.ToString & ",")
                        mOutstring.Append(mCarsStratum.ToString & ",")
                        mOutstring.Append(mWaybills.Rows(mLooper)("Total_Unmasked_Rev").ToString & ",")
                        mOutstring.Append(mWaybills.Rows(mLooper)("Tons").ToString & ",")
                        mOutstring.Append(mWaybills.Rows(mLooper)("Total_VC").ToString & ",")
                        mOutstring.Append(mWaybills.Rows(mLooper)("O_SPLC").ToString & ",")
                        mOutstring.Append(mWaybills.Rows(mLooper)("STB_Car_Typ").ToString & ",")
                        mOutstring.Append(mNumberOfWaybills.ToString & ",")
                        mOutstring.Append(bBadWaybill.ToString)
                        sbLogfile.WriteLine(mOutstring.ToString)
                    End If
                End If
            Next mLooper
        End With

        ' Get the data records we just wrote back from SQL
        mProductivityTable = New DataTable
        mProductivityTable = Get_Productivity_DataTable(cmb_URCSYear.Text)

        ' Now we need to write out the values to the legacy formatted DAT file
        outfile = My.Computer.FileSystem.OpenTextFileWriter(txt_Output_FilePath.Text, False, Encoding.ASCII)

        With mProductivityTable
            For mlooper = 0 To .Rows.Count - 1
                mWorkstr = New StringBuilder
                mWorkstr.Append("01")  'Unused Commodity Stratum
                mWorkstr.Append(Format(.Rows(mlooper)("Length_Of_Haul_Stratum"), "00"))
                mWorkstr.Append(Format(.Rows(mlooper)("Car_Type_Stratum"), "00"))
                mWorkstr.Append(Format(.Rows(mlooper)("Lading_Weight_Stratum"), "00"))
                mWorkstr.Append("1")  'Unused CarOwnership
                mWorkstr.Append(Format(.Rows(mlooper)("Cars_Stratum"), "0"))
                mWorkstr.Append(Format(.Rows(mlooper)("Waybill_Records"), "000000"))
                mWorkstr.Append(" " & CDec(.Rows(mlooper)("Revenue")).ToString("E"))
                mWorkstr.Append(Format(0, " .00000000E+00"))  ' Unused Variable Cost
                mWorkstr.Append(" " & CDec(.Rows(mlooper)("Ton_Miles")).ToString("E"))
                If CInt(cmb_URCSYear.Text) >= 2000 Then
                    mWorkstr.Append(CInt(cmb_URCSYear.Text) - 2000.ToString)
                Else
                    mWorkstr.Append(CInt(cmb_URCSYear.Text) - 1900.ToString)
                End If

                outfile.WriteLine(mWorkstr.ToString)
            Next
        End With

        outfile.Close()
        outfile = Nothing

        If mWriteCSV = True Then
            sbLogfile.Flush()
            sbLogfile.Close()
            sbLogfile = Nothing
        End If

        txt_StatusBox.Text = "Done!"
        Refresh()

SkipIt:

    End Sub

    Function Get_Owner_Flg(ByVal mCarInit As String) As Integer

        Get_Owner_Flg = 0

        Select Case mCarInit
            Case "ABOX", "Rbox", "CSX", "CSXT", "GONX"
                mCarInit = 0
            Case Else
                If Mid(mCarInit, Len(mCarInit), 1) = "X" Then
                    Get_Owner_Flg = 1
                End If
        End Select

    End Function
    Function Get_FtCar(ByVal mSTB_Car As Integer, mCar_Typ As String) As Integer

        'Emulate the car assignments as performed in the WBRecord DLL first
        Select Case mSTB_Car
            Case 36
                Get_FtCar = 1 'unequipped box cars
            Case 37
                Get_FtCar = 2 '50 ft box cars
            Case 38
                Get_FtCar = 3 'equipped box cars
            Case 39
                Get_FtCar = 4 'unequipped general service gondola
            Case 40
                Get_FtCar = 5 'equipped general service gondola
            Case 41
                Get_FtCar = 6 'covered hopper
            Case 42
                Get_FtCar = 7 'general service covered hopper
            Case 43
                Get_FtCar = 8 'open, special service hopper
            Case 44
                Get_FtCar = 9 'mechanical refridgerator
            Case 45
                Get_FtCar = 10 'non-mechanical refridgerator
            Case 46
                Get_FtCar = 11 'TOFC flat
            Case 47
                Get_FtCar = 12 'multi-level flat
            Case 48
                Get_FtCar = 13 'general service flat
            Case 49
                Get_FtCar = 14 'other flat
            Case 50
                Get_FtCar = 15 'tank, less than 22,000 gallons
            Case 51
                Get_FtCar = 16 'tank, more than 22,000 gallons
            Case 52
                Get_FtCar = 18 'all other freight cars
            Case 54
                Get_FtCar = 18 'average car - used to be cabooses
        End Select

        'Lump the "all others" and "average cars" in with unequipped box cars
        If Get_FtCar = 18 Then
            Get_FtCar = 1
        End If

        'Apply the shadow code that Legacy used
        Select Case mSTB_Car
            Case 37
                Get_FtCar = 1
            Case 38
                Get_FtCar = 2
            Case 41
                Get_FtCar = 8
            Case 42
                Get_FtCar = 6
            Case 43
                Get_FtCar = 7
            Case 44
                Get_FtCar = 11
            Case 45
                Get_FtCar = 13
            Case 46
                Get_FtCar = 17
            Case 47, 48, 49
                Get_FtCar = 10
            Case 52, 54
                Get_FtCar = 18
        End Select

        If Get_FtCar = 18 Then
            Get_FtCar = 1
        End If

        'Adding AAR Freight Car L car designations
        If mCar_Typ.First = "L" Then
            If Val(Mid(mCar_Typ, 2, 3)) < 10 Then  'This accomodates a legacy error where 0 should be assigned as FtCar = 10
                Select Case Mid(mCar_Typ, 3, 1)
                    Case "0"
                        Get_FtCar = 1
                    Case "2", "3", "9"
                        Get_FtCar = 10
                    Case "1", "4"
                        Get_FtCar = 5
                    Case "6"
                        Get_FtCar = 7
                    Case "7"
                        Get_FtCar = 3
                End Select
            Else
                Select Case mCar_Typ.Last
                    Case "0", "2", "3", "9"
                        Get_FtCar = 10
                    Case "1", "4"
                        Get_FtCar = 5
                    Case "6"
                        Get_FtCar = 7
                    Case "7"
                        Get_FtCar = 3
                End Select
            End If
        End If

        Return Get_FtCar

    End Function

    Function LengthOfHaul(ByVal mMiles As Integer) As Integer
        LengthOfHaul = 0

        If mMiles > 1 Then LengthOfHaul = 1
        If mMiles >= 250 Then LengthOfHaul = 2
        If mMiles >= 500 Then LengthOfHaul = 3
        If mMiles >= 750 Then LengthOfHaul = 4
        If mMiles >= 1000 Then LengthOfHaul = 5
        If mMiles >= 1250 Then LengthOfHaul = 6
        If mMiles >= 1500 Then LengthOfHaul = 7
        If mMiles >= 1750 Then LengthOfHaul = 8
        If mMiles >= 2000 Then LengthOfHaul = 9
        If mMiles >= 2250 Then LengthOfHaul = 10
        If mMiles >= 2500 Then LengthOfHaul = 11

    End Function

    Function Lading(ByVal mTons As Integer, mCars As Integer) As Integer
        Dim mLW As Integer = 0

        Lading = 0

        If mCars > 0 Then Lading = mTons / mCars

        If (Lading > 0 And Lading <= 20) Then mLW = 1
        If (Lading > 20 And Lading <= 25) Then mLW = 2
        If (Lading > 25 And Lading <= 30) Then mLW = 3
        If (Lading > 30 And Lading <= 35) Then mLW = 4
        If (Lading > 35 And Lading <= 40) Then mLW = 5
        If (Lading > 40 And Lading <= 45) Then mLW = 6
        If (Lading > 45 And Lading <= 50) Then mLW = 7
        If (Lading > 50 And Lading <= 55) Then mLW = 8
        If (Lading > 55 And Lading <= 60) Then mLW = 9
        If (Lading > 60 And Lading <= 65) Then mLW = 10
        If (Lading > 65 And Lading <= 70) Then mLW = 11
        If (Lading > 70 And Lading <= 75) Then mLW = 12
        If (Lading > 75 And Lading <= 80) Then mLW = 13
        If (Lading > 80 And Lading <= 85) Then mLW = 14
        If (Lading > 85 And Lading <= 90) Then mLW = 15
        If (Lading > 90 And Lading <= 95) Then mLW = 16
        If (Lading > 95 And Lading <= 100) Then mLW = 17
        If (Lading > 100 And Lading <= 105) Then mLW = 18
        If (Lading > 105 And Lading <= 110) Then mLW = 19
        If (Lading > 110) Then mLW = 20

        Lading = mLW

    End Function

    Function Get_Car_Stratum(ByVal mCars As Integer) As Integer

        Get_Car_Stratum = 0

        If (mCars >= 1 And mCars <= 2) Then Get_Car_Stratum = 1
        If (mCars >= 3 And mCars <= 5) Then Get_Car_Stratum = 2
        If (mCars >= 6 And mCars <= 15) Then Get_Car_Stratum = 3
        If (mCars >= 16 And mCars <= 49) Then Get_Car_Stratum = 4
        If (mCars >= 50 And mCars <= 60) Then Get_Car_Stratum = 5
        If (mCars >= 61 And mCars <= 100) Then Get_Car_Stratum = 6
        If (mCars > 100) Then Get_Car_Stratum = 7

    End Function
End Class