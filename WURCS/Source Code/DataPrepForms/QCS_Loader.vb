Imports SpreadsheetGear
Imports System.Data.SqlClient
Public Class QCS_Loader
    Private Const ForWriting = 8

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click

        ' Open the Data Prep Menu Form 
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()

        ' Close this Form
        Me.Close()

    End Sub

    Private Sub QCS_Loader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        Dim xclcnn As New ADODB.Connection

        Dim mDataTable As DataTable   ' for SQL work
        Dim cmdCommand As SqlCommand

        Dim mSheets(10) As String       ' For worksheet names
        Dim mWorksheet As Integer = 0

        ' Variables for SpreadsheetGear
        Dim mWorkbookSet As IWorkbookSet
        Dim mWorkbook As IWorkbook
        Dim mExcelSheet As IWorksheet
        Dim mRange As IRange
        Dim mSourceSheetName As String

        Dim mRailroad As Integer, mThisRR As String
        Dim mJunctions As Integer
        Dim mLooper As Integer, mLooper2 As Integer
        Dim mWorkVal As Decimal, mWorkStr As String
        Dim mC1 As Decimal, mC2 As Decimal, mC3 As Decimal, mC4 As Decimal
        Dim mC5 As Decimal, mC6 As Decimal, mC7 As Decimal, mC8 As Decimal
        Dim mC9 As Decimal, mC10 As Decimal, mC11 As Decimal, mC12 As Decimal
        Dim mC13 As Decimal, mC14 As Decimal
        Dim mURCSCodes(7) As Long
        Dim mAARIDs(7) As Integer
        Dim mComma As String
        Dim mStrSQL As String
        Dim mMaxRecs As Decimal
        Dim mLine As Integer, mTgtLine As Integer
        Dim mSch As Integer, mSheetCounter As Integer

        'FCS Data from URCS_FCS table
        Dim mLocal_Cars(7) As Decimal
        Dim mForwarded_Cars(7) As Decimal
        Dim mReceived_Cars(7) As Decimal
        Dim mBridged_Cars(7) As Decimal

        'Derived data from Waybills
        Dim mExpanded_Cars As Decimal
        Dim mShipment_Type As Integer
        Dim mCarType As Integer
        Dim mMoveType As Integer
        Dim mWaybill_RR(8) As Integer
        Dim mMoveType_Local(7, 19) As Decimal
        Dim mMoveType_Forwarded(7, 19) As Decimal
        Dim mMoveType_Received(7, 19) As Decimal
        Dim mMoveType_Bridged(7, 19) As Decimal
        Dim mMoveType_PctLocal(7, 19) As Decimal
        Dim mMoveType_PctForwarded(7, 19) As Decimal
        Dim mMoveType_PctReceived(7, 19) As Decimal
        Dim mMoveType_PctBridged(7, 19) As Decimal
        Dim mMoveType_LocalTotal(7) As Decimal
        Dim mMoveType_ForwardedTotal(7) As Decimal
        Dim mMoveType_ReceivedTotal(7) As Decimal
        Dim mMoveType_BridgedTotal(7) As Decimal

        ' Perform Error checking for form controls
        If Me.cmb_URCSYear.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo EndIt
        End If

        If Me.txt_Input_FilePath.Text = "" Then
            MsgBox("You must select an input file.", vbOKOnly)
            GoTo EndIt
        End If

        If Me.txt_Report_FilePath.Text = "" Then
            MsgBox("You must select an output file.", vbOKOnly)
            GoTo EndIt
        End If

        ' Delete the old file, if it exists
        If System.IO.File.Exists(txt_Report_FilePath.Text) Then
            If MsgBox("Are you sure you want to overwrite the existing file?", vbYesNo, "Warning!") = vbYes Then
                System.IO.File.Delete(txt_Report_FilePath.Text)
            Else
                GoTo EndIt
            End If
        End If


        ' Check to make sure that the Waybill data has been loaded.
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(cmb_URCSYear.Text, "MASKED")

        mStrSQL = Build_Simple_Count_Records_SQL_Statement(Gbl_Masked_TableName)

        OpenSQLConnection(My.Settings.Waybills_DB)

        cmdCommand = New SqlCommand
        cmdCommand.CommandType = CommandType.Text
        cmdCommand.Connection = gbl_SQLConnection
        cmdCommand.CommandText = Build_Simple_Count_Records_SQL_Statement(Gbl_Masked_TableName)

        If cmdCommand.ExecuteScalar < 3000 Then
            MsgBox("You must load the Waybill data before processing the FCS data.", vbOKOnly)
            GoTo EndIt
        End If

        mMaxRecs = cmdCommand.ExecuteScalar

        mComma = ","

        'Load the Arrays so that we can use the Class1Railroads and
        'URCSCodes arrays
        LoadArrayData()

        'Initialize the work arrays
        mCarType = 0
        For mRailroad = 1 To 7
            mWaybill_RR(mRailroad) = 0
        Next mRailroad

        'Dev Note - Get/set the count of railroads from table
        For mRailroad = 1 To 7
            For mMoveType = 1 To 19    '18 cartypes with additional cell for total
                mMoveType_Local(mRailroad, mCarType) = 0
                mMoveType_Forwarded(mRailroad, mCarType) = 0
                mMoveType_Received(mRailroad, mCarType) = 0
                mMoveType_Bridged(mRailroad, mCarType) = 0
                mMoveType_PctLocal(mRailroad, mCarType) = 0
                mMoveType_PctForwarded(mRailroad, mCarType) = 0
                mMoveType_PctReceived(mRailroad, mCarType) = 0
                mMoveType_PctBridged(mRailroad, mCarType) = 0
            Next mMoveType

            mMoveType_LocalTotal(mRailroad) = 0
            mMoveType_ForwardedTotal(mRailroad) = 0
            mMoveType_ReceivedTotal(mRailroad) = 0
            mMoveType_BridgedTotal(mRailroad) = 0

        Next mRailroad

        txt_StatusBox.Text = "Loading Data from input file..."
        Refresh()

        ' Open the Workbook
        mWorkbookSet = Factory.GetWorkbookSet(System.Globalization.CultureInfo.CurrentCulture)
        mWorkbook = mWorkbookSet.Workbooks.Open(txt_Input_FilePath.Text)

        ' Verify that we have all the sheets named correctly
        mSheetCounter = 0
        For Each worksheet In mWorkbook.Worksheets
            mExcelSheet = worksheet
            Select Case mExcelSheet.Name
                Case "BNSF", "CN", "CP", "CSXT", "KCS", "NS", "UP"
                    mSheetCounter = mSheetCounter + 1
            End Select
        Next
        If mSheetCounter <> 7 Then
            MsgBox("Cannot find sheets for each road.  Check the sheetnames in the Excel file.", vbOKOnly, "Error!")
            GoTo EndIt
        End If

        ' Step thru the Worksheets in the workbook
        For Each worksheet In mWorkbook.Worksheets
            mExcelSheet = worksheet
            Select Case mExcelSheet.Name.ToString
                Case "BNSF", "CN", "CP", "CSXT", "KCS", "NS", "UP"  ' We intentionally skip the US_Total sheet
                    mWorksheet = mWorksheet + 1
                    mSheets(mWorksheet) = mExcelSheet.Name.ToString
                    txt_StatusBox.Text = "Loading totals for " & mSheets(mWorksheet) & "..."

                    Select Case mExcelSheet.Name
                        Case "BNSF"
                            mWorkVal = 777
                        Case "CN"
                            mWorkVal = 103
                        Case "CP"
                            mWorkVal = 105
                        Case "CSXT"
                            mWorkVal = 712
                        Case "KCS"
                            mWorkVal = 400
                        Case "NS"
                            mWorkVal = 555
                        Case "UP"
                            mWorkVal = 802
                    End Select
                    mRailroad = Array1DFindFirst(Class1Railroads, mWorkVal)

                    ' Get the values from this sheet
                    ' Note that the AAR changed the format in 2014, so we have to accomodate
                    Select Case CInt(cmb_URCSYear.Text)
                        Case 2014
                            mLocal_Cars(mRailroad) = CDec(mExcelSheet.Cells("I514").Value)
                            mForwarded_Cars(mRailroad) = CDec(mExcelSheet.Cells("K514").Value)
                            mReceived_Cars(mRailroad) = CDec(mExcelSheet.Cells("M514").Value)
                            mBridged_Cars(mRailroad) = CDec(mExcelSheet.Cells("O514").Value)
                        Case Else '2015-2021 have been on line 509.  I don't expect this to change
                            mLocal_Cars(mRailroad) = CDec(mExcelSheet.Cells("I509").Value)
                            mForwarded_Cars(mRailroad) = CDec(mExcelSheet.Cells("K509").Value)
                            mReceived_Cars(mRailroad) = CDec(mExcelSheet.Cells("M509").Value)
                            mBridged_Cars(mRailroad) = CDec(mExcelSheet.Cells("O509").Value)
                            'Case Else
                            '    mLocal_Cars(mRailroad) = CDec(mExcelSheet.Cells("I512").Value)
                            '    mForwarded_Cars(mRailroad) = CDec(mExcelSheet.Cells("K512").Value)
                            '    mReceived_Cars(mRailroad) = CDec(mExcelSheet.Cells("M512").Value)
                            '    mBridged_Cars(mRailroad) = CDec(mExcelSheet.Cells("O512").Value)
                    End Select

                    ' Load the table name for the URCS_FCS table from the table locator table
                    Gbl_URCS_FCS_TableName = Get_Table_Name_From_SQL("1", "URCS_FCS")
                    gbl_Database_Name = Get_Database_Name_From_SQL("1", "URCS_FCS")

                    OpenSQLConnection(gbl_Database_Name)

                    cmdCommand = New SqlCommand
                    cmdCommand.CommandType = CommandType.Text
                    cmdCommand.Connection = gbl_SQLConnection
                    cmdCommand.CommandText = "SELECT Count(*) as mycount FROM " & Gbl_URCS_FCS_TableName & " WHERE Year = " &
                            cmb_URCSYear.Text & " AND AARID = " & CStr(mWorkVal)

                    If cmdCommand.ExecuteScalar > 0 Then
                        'A record exists for this railroad for the year
                        mStrSQL = Build_Update_FCS_SQL(cmb_URCSYear.Text,
                            mWorkVal,
                            mLocal_Cars(mRailroad),
                            mForwarded_Cars(mRailroad),
                            mReceived_Cars(mRailroad),
                            mBridged_Cars(mRailroad))
                    Else
                        mStrSQL = Build_Insert_FCS_SQL(cmb_URCSYear.Text,
                           mWorkVal,
                           mLocal_Cars(mRailroad),
                           mForwarded_Cars(mRailroad),
                           mReceived_Cars(mRailroad),
                           mBridged_Cars(mRailroad))
                    End If

                    cmdCommand = New SqlCommand
                    cmdCommand.Connection = gbl_SQLConnection
                    cmdCommand.CommandText = mStrSQL.ToString
                    cmdCommand.ExecuteNonQuery()

            End Select
        Next worksheet

        ' Let the user know we're getting the Waybill info
        Me.Cursor = Cursors.WaitCursor
        Me.txt_StatusBox.Text = "Fetching Waybills..."
        Me.Refresh()

        ' Check/Open SQL connection to the Waybill database
        gbl_Table_Name = Get_Table_Name_From_SQL(cmb_URCSYear.Text, "MASKED")
        gbl_Database_Name = Get_Database_Name_From_SQL(cmb_URCSYear.Text, "MASKED")

        OpenSQLConnection(gbl_Database_Name)

        mDataTable = New DataTable

        mStrSQL = "SELECT orr, jrr1, jrr2, jrr3, jrr4, jrr5, jrr6, trr, " &
                "jf, u_cars, exp_factor_th, stb_car_typ FROM " &
                gbl_Table_Name

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        For mLooper = 0 To mDataTable.Rows.Count - 1
            'Keep the user advised of where we are
            If mLooper Mod 10000 = 0 Then
                Me.txt_StatusBox.Text = "Processing record " &
                        CStr(mLooper) & " of " & CStr(mDataTable.Rows.Count)
                Text = "Progress - " &
                        CStr(Math.Round((mLooper / mDataTable.Rows.Count) * 100, MidpointRounding.AwayFromZero)) & "%"
                Me.Refresh()
            End If

            'Initialize the work array
            For mLooper2 = 1 To 8
                mWaybill_RR(mLooper2) = 0
            Next

            'Load this Waybill_RRs to the array
            Select Case mDataTable.Rows(mLooper)("jf")
                Case 0    'ORR/TRR only (They should be one in the same)
                    mWaybill_RR(1) = mDataTable.Rows(mLooper)("orr")
                Case 1    'ORR and TRR only (They should not be the same)
                    mWaybill_RR(1) = mDataTable.Rows(mLooper)("orr")
                    mWaybill_RR(2) = mDataTable.Rows(mLooper)("trr")
                Case 2    'And now for multicarrier moves w/more than 3 railroads
                    mWaybill_RR(1) = mDataTable.Rows(mLooper)("orr")
                    mWaybill_RR(2) = mDataTable.Rows(mLooper)("jrr1")
                    mWaybill_RR(3) = mDataTable.Rows(mLooper)("trr")
                Case 3
                    mWaybill_RR(1) = mDataTable.Rows(mLooper)("orr")
                    mWaybill_RR(2) = mDataTable.Rows(mLooper)("jrr1")
                    mWaybill_RR(3) = mDataTable.Rows(mLooper)("jrr2")
                    mWaybill_RR(4) = mDataTable.Rows(mLooper)("trr")
                Case 4
                    mWaybill_RR(1) = mDataTable.Rows(mLooper)("orr")
                    mWaybill_RR(2) = mDataTable.Rows(mLooper)("jrr1")
                    mWaybill_RR(3) = mDataTable.Rows(mLooper)("jrr2")
                    mWaybill_RR(4) = mDataTable.Rows(mLooper)("jrr3")
                    mWaybill_RR(5) = mDataTable.Rows(mLooper)("trr")
                Case 5
                    mWaybill_RR(1) = mDataTable.Rows(mLooper)("orr")
                    mWaybill_RR(2) = mDataTable.Rows(mLooper)("jrr1")
                    mWaybill_RR(3) = mDataTable.Rows(mLooper)("jrr2")
                    mWaybill_RR(4) = mDataTable.Rows(mLooper)("jrr3")
                    mWaybill_RR(5) = mDataTable.Rows(mLooper)("jrr4")
                    mWaybill_RR(6) = mDataTable.Rows(mLooper)("trr")
                Case 6
                    mWaybill_RR(1) = mDataTable.Rows(mLooper)("orr")
                    mWaybill_RR(2) = mDataTable.Rows(mLooper)("jrr1")
                    mWaybill_RR(3) = mDataTable.Rows(mLooper)("jrr2")
                    mWaybill_RR(4) = mDataTable.Rows(mLooper)("jrr3")
                    mWaybill_RR(5) = mDataTable.Rows(mLooper)("jrr4")
                    mWaybill_RR(6) = mDataTable.Rows(mLooper)("jrr5")
                    mWaybill_RR(7) = mDataTable.Rows(mLooper)("trr")
                Case 7
                    mWaybill_RR(1) = mDataTable.Rows(mLooper)("orr")
                    mWaybill_RR(2) = mDataTable.Rows(mLooper)("jrr1")
                    mWaybill_RR(3) = mDataTable.Rows(mLooper)("jrr2")
                    mWaybill_RR(4) = mDataTable.Rows(mLooper)("jrr3")
                    mWaybill_RR(5) = mDataTable.Rows(mLooper)("jrr4")
                    mWaybill_RR(6) = mDataTable.Rows(mLooper)("jrr5")
                    mWaybill_RR(7) = mDataTable.Rows(mLooper)("jrr6")
                    mWaybill_RR(8) = mDataTable.Rows(mLooper)("trr")
            End Select

            'Get the number of junctions
            mJunctions = mDataTable.Rows(mLooper)("jf") + 1

            'Get the expanded number of cars
            mExpanded_Cars = mDataTable.Rows(mLooper)("u_cars") * mDataTable.Rows(mLooper)("exp_factor_th")

            For mLooper2 = 1 To mJunctions
                'Process only the Class 1 Railroads only
                mRailroad = Array1DFindFirst(Class1Railroads, mWaybill_RR(mLooper2))
                If mRailroad > 0 Then

                    mCarType = STB_Car_Type(mDataTable.Rows(mLooper)("STB_Car_Typ"))

                    If mCarType > 14 Then
                        mCarType = 18      'Tank Cars and Generic Cars
                    End If

                    mShipment_Type = Shipment_Type(mLooper2, mJunctions)

                    Select Case mShipment_Type
                        Case 1  'Local
                            'add the value to the running amount for the cartype
                            mMoveType_Local(mRailroad, mCarType) =
                                    mMoveType_Local(mRailroad, mCarType) + mExpanded_Cars
                            mMoveType_LocalTotal(mRailroad) =
                                mMoveType_LocalTotal(mRailroad) + mExpanded_Cars
                        Case 2  'Forwarded
                            mMoveType_Forwarded(mRailroad, mCarType) =
                                    mMoveType_Forwarded(mRailroad, mCarType) + mExpanded_Cars
                            mMoveType_ForwardedTotal(mRailroad) =
                                mMoveType_ForwardedTotal(mRailroad) + mExpanded_Cars
                        Case 3  'Received
                            mMoveType_Received(mRailroad, mCarType) =
                                    mMoveType_Received(mRailroad, mCarType) + mExpanded_Cars
                            mMoveType_ReceivedTotal(mRailroad) =
                                mMoveType_ReceivedTotal(mRailroad) + mExpanded_Cars
                        Case 4  'Bridged
                            mMoveType_Bridged(mRailroad, mCarType) =
                                    mMoveType_Bridged(mRailroad, mCarType) + mExpanded_Cars
                            mMoveType_BridgedTotal(mRailroad) =
                                mMoveType_BridgedTotal(mRailroad) + mExpanded_Cars
                    End Select

                End If
            Next mLooper2
        Next

        'Produce the Excel file with the results for all railroads

        ' Open the Excel file
        txt_StatusBox.Text = "Creating & Loading Excel Output File..."
        Refresh()

        mWorkbook = Factory.GetWorkbook(Globalization.CultureInfo.CurrentCulture)

        mThisRR = ""
        For mRailroad = 1 To 7

            ' Get the Name for the sheet
            mStrSQL = Build_Select_RRInfo_SQL(Class1Railroads(mRailroad))
            If mStrSQL = "" Then
                'We didn't get any info for the railroad!
            Else
                ' Check/Open SQL connection
                OpenSQLConnection(Gbl_Controls_Database_Name)

                mDataTable = New DataTable

                Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                    daAdapter.Fill(mDataTable)
                End Using

                mThisRR = Trim(mDataTable.Rows(0)("rr_alpha"))

                ' Create sheet for this road
                If mWorkbook.ActiveSheet.Name <> "Sheet1" Then
                    mWorkbook.Worksheets.Add()
                End If

                mWorkbook.ActiveSheet.Name = mThisRR

            End If

            mExcelSheet = mWorkbook.ActiveSheet
            mExcelSheet.Cells().Font.Size = 12 ' Sets default font size
            mExcelSheet.Cells("A1").Value = "URCS QCS Load Processing Report - " & cmb_URCSYear.Text.ToString
            mExcelSheet.Cells("A2").Value = "Railroad: " & mThisRR

            mExcelSheet.Cells("A1:A2").Font.Bold = True
            mExcelSheet.Cells("A3").Value = "Date/Time: " & CStr(Now).ToString

            mExcelSheet.Cells("A5").Value = "Interface with Waybill Sample /QCS - Car Switching By Railroad and Car Type)"
            mExcelSheet.Cells("A6").Value = Me.cmb_URCSYear.Text & " FCS Processing"
            mExcelSheet.Cells("A7").Value = "Use: Worktable A1 - Part 6"

            mExcelSheet.Cells("A9").Value = "URCS" & vbCrLf & "Car Type"
            mExcelSheet.Cells("B9").Value = "% Local"
            mExcelSheet.Cells("C9").Value = "Local" & vbCrLf & "QCS" & vbCrLf & "Spread"
            mExcelSheet.Cells("D9").Value = "% Forwarded"
            mExcelSheet.Cells("E9").Value = "Forwarded" & vbCrLf & "QCS" & vbCrLf & "Spread"
            mExcelSheet.Cells("F9").Value = "% Received"
            mExcelSheet.Cells("G9").Value = "Received" & vbCrLf & "QCS" & vbCrLf & "Spread"
            mExcelSheet.Cells("H9").Value = "% Bridged"
            mExcelSheet.Cells("I9").Value = "Bridged" & vbCrLf & "QCS" & vbCrLf & "Spread"

            'formatting
            mExcelSheet.Cells("A1:A7").HorizontalAlignment = SpreadsheetGear.HAlign.Left
            mExcelSheet.Cells("A1:A7").WrapText = False
            mExcelSheet.Cells("A1").ColumnWidth = 10
            mExcelSheet.Cells("A9:I9").WrapText = True
            mExcelSheet.Cells("A9:I9").Font.Bold = True
            mExcelSheet.Cells("A9:I9").Font.Underline = True
            mExcelSheet.Cells("B9:I24").ColumnWidth = 14
            mExcelSheet.Cells("B9:I25").HorizontalAlignment = SpreadsheetGear.HAlign.Right
            mExcelSheet.Cells("A9:I25").Rows.AutoFit()

            mExcelSheet.Cells("B10:B25").NumberFormat = "###.000"
            mExcelSheet.Cells("C10:C25").NumberFormat = "#,##0"
            mExcelSheet.Cells("D10:D25").NumberFormat = "###.000"
            mExcelSheet.Cells("E10:E25").NumberFormat = "#,##0"
            mExcelSheet.Cells("F10:F25").NumberFormat = "###.000"
            mExcelSheet.Cells("G10:G25").NumberFormat = "#,##0"
            mExcelSheet.Cells("H10:H25").NumberFormat = "###.000"
            mExcelSheet.Cells("I10:I25").NumberFormat = "#,##0"

            'For each cartype
            For mCarType = 1 To 18
                mC1 = 0
                mC2 = 0
                mC3 = 0
                mC4 = 0
                Select Case mCarType
                    Case 1 To 14
                        mLine = mCarType + 9
                        mExcelSheet.Cells("A" & mLine.ToString).Value = CStr(mCarType) & " - " & CStr(540 + mCarType)

                        If mMoveType_Local(mRailroad, mCarType) > 0 Then
                            mExcelSheet.Cells("B" & mLine.ToString).Formula = "=(" &
                                mMoveType_Local(mRailroad, mCarType).ToString & "/" & mMoveType_LocalTotal(mRailroad).ToString & ")*100"
                            mExcelSheet.Cells("C" & mLine.ToString).Formula = "=(" &
                                mLocal_Cars(mRailroad).ToString & "*" & mExcelSheet.Cells("B" & mLine.ToString).Value.ToString & ")/100"
                            mC1 = mExcelSheet.Cells("C" & mLine.ToString).Value
                        Else
                            mExcelSheet.Cells("B" & mLine.ToString).Formula = "=0"
                            mExcelSheet.Cells("C" & mLine.ToString).Formula = "=0"
                        End If

                        If mMoveType_Forwarded(mRailroad, mCarType) > 0 Then
                            mExcelSheet.Cells("D" & mLine.ToString).Formula = "=(" &
                                mMoveType_Forwarded(mRailroad, mCarType).ToString & "/" & mMoveType_ForwardedTotal(mRailroad).ToString & ")*100"
                            mExcelSheet.Cells("E" & mLine.ToString).Formula = "=(" &
                                mForwarded_Cars(mRailroad).ToString & "*" & mExcelSheet.Cells("D" & mLine.ToString).Value.ToString & ")/100"
                            mC2 = mExcelSheet.Cells("E" & mLine.ToString).Value
                        Else
                            mExcelSheet.Cells("D" & mLine.ToString).Formula = "=0"
                            mExcelSheet.Cells("E" & mLine.ToString).Formula = "=0"
                        End If

                        If mMoveType_Received(mRailroad, mCarType) > 0 Then
                            mExcelSheet.Cells("F" & mLine.ToString).Formula = "=(" &
                                mMoveType_Received(mRailroad, mCarType).ToString & "/" & mMoveType_ReceivedTotal(mRailroad).ToString & ")*100"
                            mExcelSheet.Cells("G" & mLine.ToString).Formula = "=(" &
                                mReceived_Cars(mRailroad).ToString & "*" & mExcelSheet.Cells("F" & mLine.ToString).Value.ToString & ")/100"
                            mC3 = mExcelSheet.Cells("G" & mLine.ToString).Value
                        Else
                            mExcelSheet.Cells("F" & mLine.ToString).Formula = "=0"
                            mExcelSheet.Cells("G" & mLine.ToString).Formula = "=0"
                        End If

                        If mMoveType_Bridged(mRailroad, mCarType) > 0 Then
                            mExcelSheet.Cells("H" & mLine.ToString).Formula = "=(" &
                                mMoveType_Bridged(mRailroad, mCarType).ToString & "/" & mMoveType_BridgedTotal(mRailroad).ToString & ")*100"
                            mExcelSheet.Cells("I" & mLine.ToString).Formula = "=(" &
                                mBridged_Cars(mRailroad).ToString & "*" & mExcelSheet.Cells("H" & mLine.ToString).Value.ToString & ")/100"
                            mC4 = mExcelSheet.Cells("I" & mLine.ToString).Value
                        Else
                            mExcelSheet.Cells("H" & mLine.ToString).Formula = "=0"
                            mExcelSheet.Cells("I" & mLine.ToString).Formula = "=0"
                        End If
                    Case 15 To 17
                        'Skip these lines - they are tank cars that were added to type 18
                    Case 18
                        mLine = 24
                        mExcelSheet.Cells("A" & mLine.ToString).Value = CStr(mCarType) & " - " & CStr(540 + mCarType)

                        If mMoveType_Local(mRailroad, mCarType) > 0 Then
                            mExcelSheet.Cells("B" & mLine.ToString).Formula = "=(" &
                                mMoveType_Local(mRailroad, mCarType).ToString & "/" & mMoveType_LocalTotal(mRailroad).ToString & ")*100"
                            mExcelSheet.Cells("C" & mLine.ToString).Formula = "=(" &
                                mLocal_Cars(mRailroad).ToString & "*" & mExcelSheet.Cells("B" & mLine.ToString).Value.ToString & ")/100"
                            mC1 = mExcelSheet.Cells("C" & mLine.ToString).Value
                        Else
                            mExcelSheet.Cells("B" & mLine.ToString).Formula = "=0"
                            mExcelSheet.Cells("C" & mLine.ToString).Formula = "=0"
                        End If

                        If mMoveType_Forwarded(mRailroad, mCarType) > 0 Then
                            mExcelSheet.Cells("D" & mLine.ToString).Formula = "=(" &
                                mMoveType_Forwarded(mRailroad, mCarType).ToString & "/" & mMoveType_ForwardedTotal(mRailroad).ToString & ")*100"
                            mExcelSheet.Cells("E" & mLine.ToString).Formula = "=(" &
                                mForwarded_Cars(mRailroad).ToString & "*" & mExcelSheet.Cells("D" & mLine.ToString).Value.ToString & ")/100"
                            mC2 = mExcelSheet.Cells("E" & mLine.ToString).Value
                        Else
                            mExcelSheet.Cells("D" & mLine.ToString).Formula = "=0"
                            mExcelSheet.Cells("E" & mLine.ToString).Formula = "=0"
                        End If

                        If mMoveType_Received(mRailroad, mCarType) > 0 Then
                            mExcelSheet.Cells("F" & mLine.ToString).Formula = "=(" &
                                mMoveType_Received(mRailroad, mCarType).ToString & "/" & mMoveType_ReceivedTotal(mRailroad).ToString & ")*100"
                            mExcelSheet.Cells("G" & mLine.ToString).Formula = "=(" &
                               mReceived_Cars(mRailroad).ToString & "*" & mExcelSheet.Cells("F" & mLine.ToString).Value.ToString & ")/100"
                            mC3 = mExcelSheet.Cells("G" & mLine.ToString).Value
                        Else
                            mExcelSheet.Cells("F" & mLine.ToString).Formula = "=0"
                            mExcelSheet.Cells("G" & mLine.ToString).Formula = "=0"
                        End If

                        If mMoveType_Bridged(mRailroad, mCarType) > 0 Then
                            mExcelSheet.Cells("H" & mLine.ToString).Formula = "=(" &
                                mMoveType_Bridged(mRailroad, mCarType).ToString & "/" & mMoveType_BridgedTotal(mRailroad).ToString & ")*100"
                            mExcelSheet.Cells("I" & mLine.ToString).Formula = "=(" &
                                mBridged_Cars(mRailroad).ToString & "*" & mExcelSheet.Cells("H" & mLine.ToString).Value.ToString & ")/100"
                            mC4 = mExcelSheet.Cells("I" & mLine.ToString).Value
                        Else
                            mExcelSheet.Cells("H" & mLine.ToString).Formula = "=0"
                            mExcelSheet.Cells("I" & mLine.ToString).Formula = "=0"
                        End If
                End Select

                'set the line number into mWorkVal
                If mCarType = 18 Then
                    mWorkVal = 555
                Else
                    mWorkVal = 540 + mCarType
                End If

                If mC1 + mC2 + mC3 + mC4 > 0 Then
                    'Find out if we have already have a record in the trans table

                    ' Check/Open SQL connection
                    OpenSQLConnection(My.Settings.Controls_DB)

                    'build the SQL select statement for the line, Sch 51
                    cmdCommand = New SqlCommand
                    cmdCommand.CommandType = CommandType.Text
                    cmdCommand.Connection = gbl_SQLConnection
                    cmdCommand.CommandText = Build_Count_Trans_SQL_Statement(cmb_URCSYear.Text,
                        Class1RRICC(mRailroad),
                        51,
                        CInt(mWorkVal))

                    If cmdCommand.ExecuteScalar > 0 Then
                        'there is already a record there and we need to update it
                        mStrSQL = Build_Update_Trans_SQL_Record_Statement(Me.cmb_URCSYear.Text,
                            Class1RRICC(mRailroad),
                            51,
                            CInt(mWorkVal),
                            Math.Round(mC1, 0, MidpointRounding.AwayFromZero),
                            Math.Round(mC2, 0, MidpointRounding.AwayFromZero),
                            Math.Round(mC3, 0, MidpointRounding.AwayFromZero),
                            Math.Round(mC4, 0, MidpointRounding.AwayFromZero)).ToString
                    Else
                        'No record exists - we need to insert it
                        mStrSQL = Build_Insert_Trans_SQL_Record_Statement(Me.cmb_URCSYear.Text,
                            Class1RRICC(mRailroad),
                            51,
                            CInt(mWorkVal),
                            Math.Round(mC1, 0, MidpointRounding.AwayFromZero),
                            Math.Round(mC2, 0, MidpointRounding.AwayFromZero),
                            Math.Round(mC3, 0, MidpointRounding.AwayFromZero),
                            Math.Round(mC4, 0, MidpointRounding.AwayFromZero)).ToString
                    End If

                    cmdCommand = New SqlCommand
                    cmdCommand.Connection = gbl_SQLConnection
                    cmdCommand.CommandText = mStrSQL.ToString
                    cmdCommand.ExecuteNonQuery()

                End If

            Next mCarType

            'Now we write the totals
            mExcelSheet.Cells("A25").Value = "19 - 556"
            mExcelSheet.Cells("B25").Value = 100
            mExcelSheet.Cells("C25").Value = mLocal_Cars(mRailroad)
            mExcelSheet.Cells("D25").Value = 100
            mExcelSheet.Cells("E25").Value = mForwarded_Cars(mRailroad)
            mExcelSheet.Cells("F25").Value = 100
            mExcelSheet.Cells("G25").Value = mReceived_Cars(mRailroad)
            mExcelSheet.Cells("H25").Value = 100
            mExcelSheet.Cells("I25").Value = mBridged_Cars(mRailroad)

            'Process Line 556, Sch 51
            OpenSQLConnection(My.Settings.Controls_DB)
            cmdCommand = New SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.Connection = gbl_SQLConnection
            cmdCommand.CommandText = Build_Count_Trans_SQL_Statement(cmb_URCSYear.Text,
                Class1RRICC(mRailroad),
                51,
                556)

            If cmdCommand.ExecuteScalar > 0 Then
                'Found the total line in the trans table for this railroad
                mStrSQL = Build_Update_Trans_SQL_Record_Statement(Me.cmb_URCSYear.Text,
                    Class1RRICC(mRailroad),
                    51,
                    556,
                    mLocal_Cars(mRailroad),
                    mForwarded_Cars(mRailroad),
                    mReceived_Cars(mRailroad),
                    mBridged_Cars(mRailroad)).ToString
            Else
                'We need to insert it
                mStrSQL = Build_Insert_Trans_SQL_Record_Statement(Me.cmb_URCSYear.Text,
                    Class1RRICC(mRailroad),
                    51,
                    556,
                    mLocal_Cars(mRailroad),
                    mForwarded_Cars(mRailroad),
                    mReceived_Cars(mRailroad),
                    mBridged_Cars(mRailroad)).ToString
            End If

            cmdCommand = New SqlCommand
            cmdCommand.Connection = gbl_SQLConnection
            cmdCommand.CommandText = mStrSQL.ToString
            cmdCommand.ExecuteNonQuery()

            ' Write the footer info
            mExcelSheet.Cells("A27").Value = "Input File Location: " & txt_Input_FilePath.Text
            mExcelSheet.Cells("A28").Value = "Local Traffic Cars Value: " & Format(mLocal_Cars(mRailroad), "##,##0")
            mExcelSheet.Cells("A29").Value = "Forwarded Cars Value: " & Format(mForwarded_Cars(mRailroad), "##,##0")
            mExcelSheet.Cells("A30").Value = "Received Cars Value: " & Format(mReceived_Cars(mRailroad), "##,##0")
            mExcelSheet.Cells("A31").Value = "Bridged Cars Value: " & Format(mBridged_Cars(mRailroad), "##,##0")

            ' Print formatting and switch to landscape
            mExcelSheet.PageSetup.PrintArea = mExcelSheet.Cells("A1:I31").ToString
            mExcelSheet.PageSetup.Orientation = PageOrientation.Landscape
            mExcelSheet.PageSetup.FitToPages = True

            'Save this sheets name for referencing back to it from Variance sheet
            mSourceSheetName = mExcelSheet.Name

            mWorkbook.Worksheets.Add()
            mWorkbook.ActiveSheet.Name = mExcelSheet.Name & " Variance"
            mExcelSheet = mWorkbook.ActiveSheet

            mExcelSheet.Cells().Font.Size = 12
            mExcelSheet.Cells("A1").Value = "URCS QCS Variance Report - " & cmb_URCSYear.Text.ToString & " v. " &
                CStr(CInt(cmb_URCSYear.Text) - 1)
            mExcelSheet.Cells("A2").Value = "Railroad: " & mThisRR
            mExcelSheet.Cells("A1").Font.Size = 18
            mExcelSheet.Cells("A1:A2").Font.Bold = True
            mExcelSheet.Cells("A3").Value = "Railroad: " & mThisRR
            mExcelSheet.Cells("A3").Value = "Date/Time: " & CStr(Now).ToString
            mExcelSheet.Cells("A1:A3").HorizontalAlignment = HorizontalAlignment.Left

            mExcelSheet.Cells("A5").Value = "URCS" & vbCrLf & "Car Type"
            mExcelSheet.Cells("B5").Value = "Traffic Type"
            mExcelSheet.Cells("C5").Value = cmb_URCSYear.Text & " Value"
            mExcelSheet.Cells("D5").Value = CStr(CInt(cmb_URCSYear.Text) - 1) & " Value"
            mExcelSheet.Cells("E5").Value = "% Variance"
            mExcelSheet.Cells("F5").Value = "Impact"
            mExcelSheet.Cells("G5").Value = "Action"
            mExcelSheet.Cells("A5:G5").Font.Underline = True
            mExcelSheet.Cells("A5:G5").Font.Bold = True
            mExcelSheet.Cells("A9:I9").WrapText = True
            mExcelSheet.Cells("C5:E5").HorizontalAlignment = SpreadsheetGear.HAlign.Right

            'Format data cells
            mExcelSheet.Cells("B6").ColumnWidth = 14
            mExcelSheet.Cells("C6:E70").ColumnWidth = 12
            mExcelSheet.Cells("C6:E70").HorizontalAlignment = SpreadsheetGear.HAlign.Right
            mExcelSheet.Cells("C6:C70").NumberFormat = "#,##0"
            mExcelSheet.Cells("D6:D70").NumberFormat = "#,##0"
            mExcelSheet.Cells("E6:E70").NumberFormat = "#0.00"

            mTgtLine = 7

            For mLine = 10 To 25

                'Write the car type value
                mExcelSheet.Cells("A" & mTgtLine.ToString).Value = mWorkbook.Worksheets(mSourceSheetName).Cells("A" & mLine.ToString).Value

                For mLooper = 1 To 4
                    Select Case mLooper
                        Case 1
                            'write the traffic type
                            mExcelSheet.Cells("B" & mTgtLine.ToString).Value = "Local"

                            ' Write this years value
                            mExcelSheet.Cells("C" & mTgtLine.ToString).Value = mWorkbook.Worksheets(mSourceSheetName).Cells("C" & mLine.ToString).Value

                            ' Get and write last year's value
                            mWorkStr = mWorkbook.Worksheets(mSourceSheetName).Cells("A" & mLine.ToString).Value
                            mExcelSheet.Cells("D" & mTgtLine.ToString).Value = Get_Trans_Value(
                                (cmb_URCSYear.Text - 1),
                                Get_RRICC_By_Short_Name(mSourceSheetName),
                                51,
                                mWorkStr.Substring(Len(mWorkStr) - 3),
                                "C1")

                        Case 2
                            'write the traffic type
                            mExcelSheet.Cells("B" & mTgtLine.ToString).Value = "Forwarded"

                            ' Write this years value
                            mExcelSheet.Cells("C" & mTgtLine.ToString).Value = mWorkbook.Worksheets(mSourceSheetName).Cells("E" & mLine.ToString).Value

                            ' Get and write last year's value
                            mWorkStr = mWorkbook.Worksheets(mSourceSheetName).Cells("A" & mLine.ToString).Value
                            mExcelSheet.Cells("D" & mTgtLine.ToString).Value = Get_Trans_Value(
                                (cmb_URCSYear.Text - 1),
                                Get_RRICC_By_Short_Name(mSourceSheetName),
                                51,
                                mWorkStr.Substring(Len(mWorkStr) - 3),
                                "C2")

                        Case 3
                            'write the traffic type
                            mExcelSheet.Cells("B" & mTgtLine.ToString).Value = "Received"

                            ' Write this years value
                            mExcelSheet.Cells("C" & mTgtLine.ToString).Value = mWorkbook.Worksheets(mSourceSheetName).Cells("G" & mLine.ToString).Value

                            ' Get and write last year's value
                            mWorkStr = mWorkbook.Worksheets(mSourceSheetName).Cells("A" & mLine.ToString).Value
                            mExcelSheet.Cells("D" & mTgtLine.ToString).Value = Get_Trans_Value(
                                (cmb_URCSYear.Text - 1),
                                Get_RRICC_By_Short_Name(mSourceSheetName),
                                51,
                                mWorkStr.Substring(Len(mWorkStr) - 3),
                                "C3")

                        Case 4
                            'write the traffic type
                            mExcelSheet.Cells("B" & mTgtLine.ToString).Value = "Bridged"

                            ' Write this years value
                            mExcelSheet.Cells("C" & mTgtLine.ToString).Value = mWorkbook.Worksheets(mSourceSheetName).Cells("I" & mLine.ToString).Value

                            ' Get and write last year's value
                            mWorkStr = mWorkbook.Worksheets(mSourceSheetName).Cells("A" & mLine.ToString).Value
                            mExcelSheet.Cells("D" & mTgtLine.ToString).Value = Get_Trans_Value(
                                (cmb_URCSYear.Text - 1),
                                Get_RRICC_By_Short_Name(mSourceSheetName),
                                51,
                                mWorkStr.Substring(Len(mWorkStr) - 3),
                                "C4")
                    End Select

                    'write the variance value
                    mExcelSheet.Cells("E" & mTgtLine.ToString).Value =
                        Calculate_Variance(mExcelSheet.Cells("D" & mTgtLine.ToString).Value,
                                           mExcelSheet.Cells("C" & mTgtLine.ToString).Value)

                    mTgtLine = mTgtLine + 1

                Next mLooper
            Next mLine
        Next mRailroad

        mWorkbook.Worksheets(0).Select()
        mWorkbook.SaveAs(txt_Report_FilePath.Text, FileFormat.OpenXMLWorkbook)

        ' Lastly, we need to extract the FCS data from the spreadsheet
        ' total lines for each of the commodities and save them to the
        ' trans table

        ' Read the excel spreadsheet file that the AAR sends us and write
        ' the data to the Trans table
        mWorkbook = Factory.GetWorkbook(txt_Input_FilePath.Text, Globalization.CultureInfo.CurrentCulture)

        For Each sheet As IWorksheet In mWorkbook.Worksheets
            Select Case sheet.Name
                Case "BNSF", "CN", "CP", "CSXT", "KCS", "NS", "UP"
                    ' We intentionally skip the US_Total sheet

                    ' Tell the user what we're doing
                    Text = "Processing FCS for " & sheet.Name & "..."
                    txt_StatusBox.Text = "Processing FCS for " & sheet.Name & "..."
                    Refresh()

                    mRailroad = Get_RRICC_By_Short_Name(sheet.Name)

                    mRange = sheet.Cells

                    Select Case cmb_URCSYear.Text
                        Case 2015
                            mLine = 509
                        Case 2014, 2016
                            mLine = 514
                        Case Else
                            mLine = 512
                    End Select

                    For mLooper = 9 To mLine
                        mSch = Set_Commodity_Schedule(mRange.Cells("A" & mLooper.ToString).Value)
                        mLine = 0
                        If mSch > 0 Then
                            ' we should have a total line at this point
                            Select Case mSch
                                Case 146 To 147
                                    ' we should have a total line at this point
                                    ' line set to 900
                                    mLine = 900
                            End Select

                            'fetch the values from the cols in the Excel file
                            'default to zero if a null is encountered
                            mC1 = ReturnValidNumber(mRange.Cells("I" & mLooper.ToString).Value)
                            mC2 = ReturnValidNumber(mRange.Cells("J" & mLooper.ToString).Value)
                            mC3 = ReturnValidNumber(mRange.Cells("K" & mLooper.ToString).Value)
                            mC4 = ReturnValidNumber(mRange.Cells("L" & mLooper.ToString).Value)
                            mC5 = ReturnValidNumber(mRange.Cells("M" & mLooper.ToString).Value)
                            mC6 = ReturnValidNumber(mRange.Cells("N" & mLooper.ToString).Value)
                            mC7 = ReturnValidNumber(mRange.Cells("O" & mLooper.ToString).Value)
                            mC8 = ReturnValidNumber(mRange.Cells("P" & mLooper.ToString).Value)
                            mC9 = ReturnValidNumber(mRange.Cells("G" & mLooper.ToString).Value)
                            mC10 = ReturnValidNumber(mRange.Cells("H" & mLooper.ToString).Value)
                            mC11 = ReturnValidNumber(mRange.Cells("Q" & mLooper.ToString).Value)

                            'Compute Total Carloads Originated and Terminated (CLOT)
                            mC12 = (mC1 * 2) + mC3 + mC5

                            'Compute Total Carloads Handled (CLOR)
                            mC13 = mC1 + mC3 + mC5 + mC7

                            'Compute Total Carloads Interchanged (CLRF)
                            mC14 = mC3 + mC5 + (mC7 * 2)

                            ' Check/Open SQL connection to the database
                            OpenSQLConnection(My.Settings.Controls_DB)

                            cmdCommand = New SqlCommand
                            cmdCommand.CommandType = CommandType.Text
                            cmdCommand.Connection = gbl_SQLConnection
                            cmdCommand.CommandText = Build_Count_Trans_SQL_Statement(cmb_URCSYear.Text,
                                mRailroad,
                                mSch,
                                mLine)

                            If cmdCommand.ExecuteScalar > 0 Then
                                'there is already a record there and we need
                                'to update it
                                mStrSQL = Build_Update_Trans_SQL_Record_Statement(
                                    Me.cmb_URCSYear.Text,
                                    mRailroad,
                                    mSch,
                                    mLine,
                                    mC1,
                                    mC2,
                                    mC3,
                                    mC4,
                                    mC5,
                                    mC6,
                                    mC7,
                                    mC8,
                                    mC9,
                                    mC10,
                                    mC11,
                                    mC12,
                                    mC13,
                                    mC14).ToString
                            Else
                                'No record exists - we need to insert it
                                mStrSQL = Build_Insert_Trans_SQL_Record_Statement(
                                    Me.cmb_URCSYear.Text,
                                    mRailroad,
                                    mSch,
                                    mLine,
                                    mC1,
                                    mC2,
                                    mC3,
                                    mC4,
                                    mC5,
                                    mC6,
                                    mC7,
                                    mC8,
                                    mC9,
                                    mC10,
                                    mC11,
                                    mC12,
                                    mC13,
                                    mC14).ToString
                            End If

                            cmdCommand = New SqlCommand
                            cmdCommand.Connection = gbl_SQLConnection
                            cmdCommand.CommandText = mStrSQL.ToString
                            cmdCommand.ExecuteNonQuery()

                        End If
                    Next
            End Select
        Next sheet

        Text = "Done!"
        txt_StatusBox.Text = "Done!"
        Refresh()

EndIt:
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub btn_Input_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub btn_Report_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Report_File_Entry.Click
        Dim fd As New FolderBrowserDialog

        If cmb_URCSYear.Text = "" Then
            MsgBox("You must select a year first.", vbOKOnly, "Error!")
        Else
            fd.Description = "Select the location in which you want the output report placed."

            If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txt_Report_FilePath.Text = fd.SelectedPath.ToString & "\WB" & cmb_URCSYear.Text & " QCS Load Report.xlsx"
            End If
        End If
    End Sub
End Class