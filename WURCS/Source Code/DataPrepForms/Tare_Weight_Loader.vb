Imports System.Data.SqlClient
Imports SpreadsheetGear
Public Class Tare_Weight_Loader

    '**********************************************************************
    ' Title:        Tare Weight Loader Form
    ' Author:       Michael Sanders
    ' Purpose:      This form handles the entry and updating of Tare Weight information for URCS.
    ' Revisions:    Conversion from Access database/VBA - 14 Mar 2013
    '               Conversion to Compartmentalized Databases - 23 Jul 2015
    '               Change to add default report name - 8/30/2019
    ' 
    ' This program is US Government Property - For Official Use Only
    '**********************************************************************

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        Dim mCommand As SqlCommand
        Dim TempTable As DataTable, mDatatable As DataTable

        ' Variables for SpreadsheetGear
        Dim mWorkbookSet As IWorkbookSet
        Dim mWorkbook As IWorkbook
        Dim mExcelSheet As IWorksheet

        Dim bolWrite As Integer, mLooper As Integer
        Dim mSumOfCars(29) As Decimal
        Dim mSumOfTare(29) As Decimal
        Dim mAvgTareWeight(29) As Decimal

        Dim mLine_no As String
        Dim mDesc As String
        Dim idx As Integer
        Dim mStrSQL As String
        Dim mWorkVal As Single, mStartline As Integer, mThisYearVal As Integer, mLastYearVal As Single
        Dim mColB As Single, mColC As Integer

        'Ensure we have the trans table and database loaded
        Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "Trans")

        ' Check to make sure that the user has selected a year, an input file and an output file
        If IsNothing(Me.cmb_URCSYear.Text) Then
            MsgBox("You must select a year value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        If txt_Input_FilePath.TextLength = 0 Then
            MsgBox("You must select an Input File value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        If txt_Report_FilePath.TextLength = 0 Then
            MsgBox("You must select an Report File value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        ' Delete the old file, if it exists
        If System.IO.File.Exists(txt_Report_FilePath.Text) Then
            If MsgBox("Are you sure you want to overwrite the existing file?", vbYesNo, "Warning!") = vbYes Then
                File.Delete(txt_Report_FilePath.Text)
            Else
                GoTo EndIt
            End If
        End If

        mLine_no = " "
        mDesc = " "

        bolWrite = MsgBox("Are you sure you want to Write this data?", vbYesNo)

        If bolWrite = vbYes Then

            'Initialize the arrays
            For idx = 1 To 29
                mSumOfCars(idx) = 0
                mSumOfTare(idx) = 0
                mAvgTareWeight(idx) = 0
            Next

            'Create the xclTable that we'll use for storing a computing values from Excel file
            TempTable = New DataTable
            ' Add the columns
            TempTable.Columns.Add("Car_Type_Code", Type.GetType("System.String"))
            TempTable.Columns.Add("Average_Tare_Weight", Type.GetType("System.Double"))
            TempTable.Columns.Add("No_Tare_Cars", Type.GetType("System.Double"))
            TempTable.Columns.Add("Tare_Cars", Type.GetType("System.Double"))
            TempTable.Columns.Add("Total_Tare_Weight", Type.GetType("System.Double"))

            ' Open the Workbook
            mWorkbookSet = Factory.GetWorkbookSet(System.Globalization.CultureInfo.CurrentCulture)
            mWorkbook = mWorkbookSet.Workbooks.Open(txt_Input_FilePath.Text)
            mExcelSheet = mWorkbook.Sheets(0)

            'Find the A100 Line
            For mStartline = 1 To 15
                If IsDBNull(mExcelSheet.Cells("A" & mStartline).Value) = False Then
                    If mExcelSheet.Cells("A" & mStartline).Value = "" Then
                        ' Skip it
                    Else
                        If mExcelSheet.Cells("A" & mStartline).Value = "A100" Then
                            Exit For
                        End If
                    End If
                End If
            Next

            'Read the spreadsheet, adding the data to the temp table
            For mLooper = mStartline To (mExcelSheet.UsedRange.RowCount - mStartline)
                mColB = 0
                mColC = 0
                If IsNumeric(mExcelSheet.Cells("B" & mLooper).Value) Then
                    mColB = mExcelSheet.Cells("B" & mLooper).Value
                End If
                If IsNumeric(mExcelSheet.Cells("C" & mLooper).Value) Then
                    mColC = mExcelSheet.Cells("C" & mLooper).Value
                End If
                mWorkVal = mColB * mColC
                ' Add the data from Excel to the TempTable
                TempTable.Rows.Add(mExcelSheet.Cells("A" & mLooper).Value,
                                   Math.Round(mColB, 0, MidpointRounding.AwayFromZero),
                                   CDbl(0),
                                   mColC,
                                   Math.Round(mWorkVal, 0, MidpointRounding.AwayFromZero))
                'Sort the table on the Car Type Code
                TempTable.DefaultView.Sort = "Car_Type_Code ASC"
            Next

            'OK, we're done loading the source
            'Begin computations from the data stored in the TempTable
            For Each xclRow As DataRow In TempTable.Rows
                Select Case Microsoft.VisualBasic.Left(xclRow(0), 1)
                    Case "A"
                        If Val(Microsoft.VisualBasic.Mid(xclRow(0), 3, 1)) = 5 Then
                            mSumOfCars(21) = mSumOfCars(21) + xclRow(3)
                            mSumOfTare(21) = mSumOfTare(21) + xclRow(4)
                        Else
                            mSumOfCars(3) = mSumOfCars(3) + xclRow(3)
                            mSumOfTare(3) = mSumOfTare(3) + xclRow(4)
                        End If
                    Case "B"
                        Select Case Microsoft.VisualBasic.Left(xclRow(0), 2)
                            Case "B1", "B2"
                                mSumOfCars(1) = mSumOfCars(1) + xclRow(3)
                                mSumOfTare(1) = mSumOfTare(1) + xclRow(4)
                            Case "B3", "B4"
                                Select Case Val(Microsoft.VisualBasic.Right(xclRow(0), 1))
                                    Case 0 To 7
                                        mSumOfCars(2) = mSumOfCars(2) + xclRow(3)
                                        mSumOfTare(2) = mSumOfTare(2) + xclRow(4)
                                End Select
                            Case "B5", "B6", "B7", "B8"
                                mSumOfCars(2) = mSumOfCars(2) + xclRow(3)
                                mSumOfTare(2) = mSumOfTare(2) + xclRow(4)
                        End Select
                    Case "C"
                        Select Case Val(Microsoft.VisualBasic.Right(xclRow(0), 1))
                            Case 1 To 4
                                mSumOfCars(6) = mSumOfCars(6) + xclRow(3)
                                mSumOfTare(6) = mSumOfTare(6) + xclRow(4)
                        End Select
                    Case "E"
                        mSumOfCars(5) = mSumOfCars(5) + xclRow(3)
                        mSumOfTare(5) = mSumOfTare(5) + xclRow(4)
                    Case "F"
                        Select Case Val(Mid(xclRow(0), 2, 2))
                            Case 10, 20, 30
                                mSumOfCars(17) = mSumOfCars(17) + xclRow(3)
                                mSumOfTare(17) = mSumOfTare(17) + xclRow(4)
                            Case 40
                                mSumOfCars(18) = mSumOfCars(18) + xclRow(3)
                                mSumOfTare(18) = mSumOfTare(18) + xclRow(4)
                        End Select
                        Select Case Val(Mid(xclRow(0), 3, 1))
                            Case 1 To 6, 8
                                mSumOfCars(18) = mSumOfCars(18) + xclRow(3)
                                mSumOfTare(18) = mSumOfTare(18) + xclRow(4)
                            Case 7
                                mSumOfCars(21) = mSumOfCars(21) + xclRow(3)
                                mSumOfTare(21) = mSumOfTare(21) + xclRow(4)
                        End Select
                    Case "G"
                        mSumOfCars(4) = mSumOfCars(4) + xclRow(3)
                        mSumOfTare(4) = mSumOfTare(4) + xclRow(4)
                    Case "H"
                        mSumOfCars(7) = mSumOfCars(7) + xclRow(3)
                        mSumOfTare(7) = mSumOfTare(7) + xclRow(4)
                    Case "J"
                        Select Case Val(Microsoft.VisualBasic.Right(xclRow(0), 1))
                            Case 0
                                mSumOfCars(8) = mSumOfCars(8) + xclRow(3)
                                mSumOfTare(8) = mSumOfTare(8) + xclRow(4)
                            Case 1 To 4
                                mSumOfCars(4) = mSumOfCars(4) + xclRow(3)
                                mSumOfTare(4) = mSumOfTare(4) + xclRow(4)
                        End Select
                    Case "K"
                        mSumOfCars(8) = mSumOfCars(8) + xclRow(3)
                        mSumOfTare(8) = mSumOfTare(8) + xclRow(4)
                    Case "L"
                        mSumOfCars(21) = mSumOfCars(21) + xclRow(3)
                        mSumOfTare(21) = mSumOfTare(21) + xclRow(4)
                    Case "M"
                        If Val(Mid(xclRow(0), 2, 3)) = 930 Then
                            mSumOfCars(23) = mSumOfCars(23) + xclRow(3)
                            mSumOfTare(23) = mSumOfTare(23) + xclRow(4)
                        End If
                    Case "P"
                        mSumOfCars(11) = mSumOfCars(11) + xclRow(3)
                        mSumOfTare(11) = mSumOfTare(11) + xclRow(4)
                        mSumOfCars(15) = mSumOfCars(15) + xclRow(3)
                        mSumOfTare(15) = mSumOfTare(15) + xclRow(4)
                    Case "Q"
                        If Val(Mid(xclRow(0), 2, 1)) = 8 Then
                            mSumOfCars(13) = mSumOfCars(13) + xclRow(3)
                            mSumOfTare(13) = mSumOfTare(13) + xclRow(4)
                        Else
                            mSumOfCars(12) = mSumOfCars(12) + xclRow(3)
                            mSumOfTare(12) = mSumOfTare(12) + xclRow(4)
                            mSumOfCars(15) = mSumOfCars(15) + xclRow(3)
                            mSumOfTare(15) = mSumOfTare(15) + xclRow(4)
                        End If
                    Case "R"
                        Select Case Val(Mid(xclRow(0), 3, 1))
                            Case 0 To 2
                                mSumOfCars(10) = mSumOfCars(10) + xclRow(3)
                                mSumOfTare(10) = mSumOfTare(10) + xclRow(4)
                            Case 5 To 9
                                mSumOfCars(9) = mSumOfCars(9) + xclRow(3)
                                mSumOfTare(9) = mSumOfTare(9) + xclRow(4)
                        End Select
                    Case "S"
                        mSumOfCars(14) = mSumOfCars(14) + xclRow(3)
                        mSumOfTare(14) = mSumOfTare(14) + xclRow(4)
                        mSumOfCars(15) = mSumOfCars(15) + xclRow(3)
                        mSumOfTare(15) = mSumOfTare(15) + xclRow(4)
                    Case "T"
                        If Val(Mid(xclRow(0), 2, 3)) > 0 Then
                            Select Case Val(Microsoft.VisualBasic.Right(xclRow(0), 1))
                                Case 0 To 5
                                    mSumOfCars(19) = mSumOfCars(19) + xclRow(3)
                                    mSumOfTare(19) = mSumOfTare(19) + xclRow(4)
                                Case 6 To 9
                                    mSumOfCars(20) = mSumOfCars(20) + xclRow(3)
                                    mSumOfTare(20) = mSumOfTare(20) + xclRow(4)
                            End Select
                        End If
                    Case "U"
                        If Val(Mid(xclRow(0), 2, 1)) = 5 Then
                            mSumOfCars(26) = mSumOfCars(26) + xclRow(3)
                            mSumOfTare(26) = mSumOfTare(26) + xclRow(4)
                        Else
                            mSumOfCars(27) = mSumOfCars(27) + xclRow(3)
                            mSumOfTare(27) = mSumOfTare(27) + xclRow(4)
                        End If
                    Case "V"
                        mSumOfCars(16) = mSumOfCars(16) + xclRow(3)
                        mSumOfTare(16) = mSumOfTare(16) + xclRow(4)
                    Case "Z"
                        If Val(Mid(xclRow(0), 2, 1)) = 5 Then
                            mSumOfCars(24) = mSumOfCars(24) + xclRow(3)
                            mSumOfTare(24) = mSumOfTare(24) + xclRow(4)
                        Else
                            mSumOfCars(25) = mSumOfCars(25) + xclRow(3)
                            mSumOfTare(25) = mSumOfTare(25) + xclRow(4)
                        End If
                End Select

            Next

            ' Add the tank cars to arrays pos 21
            mSumOfCars(21) = mSumOfCars(21) + mSumOfCars(19) + mSumOfCars(20)
            mSumOfTare(21) = mSumOfTare(21) + mSumOfTare(19) + mSumOfTare(20)

            'Compute the subtotal and total lines
            mSumOfCars(22) = 0
            mSumOfTare(22) = 0

            For idx = 1 To 10
                mSumOfCars(22) = mSumOfCars(22) + mSumOfCars(idx)
                mSumOfTare(22) = mSumOfTare(22) + mSumOfTare(idx)
            Next
            For idx = 15 To 18
                mSumOfCars(22) = mSumOfCars(22) + mSumOfCars(idx)
                mSumOfTare(22) = mSumOfTare(22) + mSumOfTare(idx)
            Next

            mSumOfCars(22) = mSumOfCars(21) + mSumOfCars(22)
            mSumOfTare(22) = mSumOfTare(21) + mSumOfTare(22)

            mSumOfCars(28) = mSumOfCars(24) + mSumOfCars(26)
            mSumOfTare(28) = mSumOfTare(24) + mSumOfTare(26)
            mSumOfCars(29) = mSumOfCars(25) + mSumOfCars(27)
            mSumOfTare(29) = mSumOfTare(25) + mSumOfTare(27)

            'Compute the Average Tare Weight
            For idx = 1 To 29
                If Val(cmb_URCSYear.Text) < 2011 Then
                    ' Tare in thousands of lbs (prior to 2011)
                    mAvgTareWeight(idx) = ((mSumOfTare(idx) * 1000) / mSumOfCars(idx)) / 2000
                Else
                    ' Tare is in tons (2011 & later)
                    If mSumOfCars(idx) > 0 Then
                        mAvgTareWeight(idx) = mSumOfTare(idx) / mSumOfCars(idx)
                    Else
                        mAvgTareWeight(idx) = 0
                    End If
                End If
            Next

            Me.txt_StatusBox.Text = "Creating & Loading Excel File..."
            Me.Refresh()

            ' Open the Workbook
            mWorkbookSet = Factory.GetWorkbookSet(System.Globalization.CultureInfo.CurrentCulture)
            mWorkbook = mWorkbookSet.Workbooks.Open(txt_Input_FilePath.Text)

            'Generate the report
            ' Set the active sheet's name
            If mWorkbook.ActiveSheet.Name <> "Sheet1" Then
                mWorkbook.Worksheets.Add()
            End If

            mWorkbook.ActiveSheet.Name = "Summary"

            ' Set the data to the Summary page
            mExcelSheet = mWorkbook.ActiveSheet
            mExcelSheet.Cells().Font.Size = 12 ' Sets default font size

            mExcelSheet.Cells("A1").Value = "URCS Tare Weight Load Process Report - " & cmb_URCSYear.Text.ToString
            mExcelSheet.Cells("A2").Value = "Input File: " & Me.txt_Input_FilePath.Text
            mExcelSheet.Cells("A3").Value = "Date/Time: " & CStr(Now).ToString
            mExcelSheet.Cells("A1:A2").Font.Bold = True

            mExcelSheet.Cells("A5").Value = "Item" & vbCrLf & "No."
            mExcelSheet.Cells("B5").Value = "URCS" & vbCrLf & "Line No."
            mExcelSheet.Cells("C5").Value = "URCS Car Type"
            mExcelSheet.Cells("D5").Value = "Sum of" & vbCrLf & "Cars"
            mExcelSheet.Cells("E5").Value = "Sum of" & vbCrLf & "Tare"
            mExcelSheet.Cells("F5").Value = "Avg Tare" & vbCrLf & "Weight"
            mExcelSheet.Cells("G5").Value = (CInt(cmb_URCSYear.Text) - 1).ToString & vbCrLf & "Avg" & vbCrLf & "From" & vbCrLf & "Trans"
            mExcelSheet.Cells("H5").Value = "Variance"

            mExcelSheet.Cells("A6:B6").HorizontalAlignment = SpreadsheetGear.HAlign.Left
            mExcelSheet.Cells("D6:G6").HorizontalAlignment = SpreadsheetGear.HAlign.Left
            mExcelSheet.Cells("H6:H40").HorizontalAlignment = SpreadsheetGear.HAlign.Right

            mExcelSheet.Cells("A6:H6").Font.Bold = True
            mExcelSheet.Cells("A6:H6").Font.Underline = True

            mExcelSheet.Cells("A8:A40").ColumnWidth = 8
            mExcelSheet.Cells("C8:C40").ColumnWidth = 40
            mExcelSheet.Cells("D8:D40").ColumnWidth = 11
            mExcelSheet.Cells("E8:E40").ColumnWidth = 14
            mExcelSheet.Cells("F8:F40").ColumnWidth = 8
            mExcelSheet.Cells("G8:G40").ColumnWidth = 8
            mExcelSheet.Cells("H8:H40").ColumnWidth = 10
            mExcelSheet.Cells("F8:H40").NumberFormat = "#0.0"


            For idx = 1 To 29
                Select Case idx
                    Case 1
                        mLine_no = "501"
                        mDesc = "Plain Box - 40"
                    Case 2
                        mLine_no = "502"
                        mDesc = "Plain Box - 50"
                    Case 3
                        mLine_no = "503"
                        mDesc = "Equipped Box"
                    Case 4
                        mLine_no = "504"
                        mDesc = "Plain Gondola"
                    Case 5
                        mLine_no = "505"
                        mDesc = "Equipped Gondola"
                    Case 6
                        mLine_no = "506"
                        mDesc = "Covered Hopper"
                    Case 7
                        mLine_no = "507"
                        mDesc = "Open Top Hopper - General"
                    Case 8
                        mLine_no = "508"
                        mDesc = "Open Top Hopper - Special"
                    Case 9
                        mLine_no = "509"
                        mDesc = "Refrigerated - Mechanical"
                    Case 10
                        mLine_no = "510"
                        mDesc = "Refrigerated - Non-Mechanical"
                    Case 11
                        mLine_no = " "
                        mDesc = "Conventional Intermodal"
                    Case 12
                        mLine_no = " "
                        mDesc = "Light, Low Profile Intermodal"
                    Case 13
                        mLine_no = " "
                        mDesc = "Road Railers"
                    Case 14
                        mLine_no = " "
                        mDesc = "Stack Cars"
                    Case 15
                        mLine_no = "511"
                        mDesc = "Flat - TOFC/COFC"
                    Case 16
                        mLine_no = "512"
                        mDesc = "Flat - Multi-Level"
                    Case 17
                        mLine_no = "513"
                        mDesc = "Flat - General Service"
                    Case 18
                        mLine_no = "514"
                        mDesc = "Flat - Other"
                    Case 19
                        mLine_no = "517"
                        mDesc = "Tank < 22,000 Gallons"
                    Case 20
                        mLine_no = "518"
                        mDesc = "Tank > 22,000 Gallons"
                    Case 21
                        mLine_no = "515"
                        mDesc = "All Other"
                    Case 22
                        mLine_no = "516"
                        mDesc = "Total"
                    Case 23
                        mLine_no = " "
                        mDesc = "Cabooses"
                    Case 24
                        mLine_no = " "
                        mDesc = "Trailers - Refrigerated"
                    Case 25
                        mLine_no = " "
                        mDesc = "Trailers - Non-Refrigerated"
                    Case 26
                        mLine_no = " "
                        mDesc = "Containers - Refrigerated"
                    Case 27
                        mLine_no = " "
                        mDesc = "Containers - Non-Refrigerated"
                    Case 28
                        mLine_no = "584"
                        mDesc = "Refrigerated - Trailers and Containers"
                    Case 29
                        mLine_no = "585"
                        mDesc = "Non-Refrigerated - Trailers and Containers"
                End Select

                mExcelSheet.Cells("A" & (idx + 7).ToString).Value = idx
                mExcelSheet.Cells("B" & (idx + 7).ToString).Value = mLine_no
                mExcelSheet.Cells("C" & (idx + 7).ToString).Value = mDesc
                mExcelSheet.Cells("D" & (idx + 7).ToString).Value = Format(mSumOfCars(idx), "###,###,###")
                mExcelSheet.Cells("E" & (idx + 7).ToString).Value = Format(mSumOfTare(idx), "###,###,###")
                mExcelSheet.Cells("F" & (idx + 7).ToString).Value = Math.Round(mAvgTareWeight(idx), 1, MidpointRounding.AwayFromZero)

                mLastYearVal = 0
                mThisYearVal = 0
                'get the value from the previous year
                If Trim(mLine_no) <> "" Then
                    mWorkVal = CInt(cmb_URCSYear.Text) - 1
                    Select Case mLine_no
                        Case 584, 585
                            mLastYearVal = Get_Trans_Value(mWorkVal, 900099, 42, mLine_no, "C1")
                        Case Else
                            mLastYearVal = Get_Trans_Value(mWorkVal, 900004, 41, mLine_no, "C1")
                    End Select

                    mLastYearVal = Math.Round(mLastYearVal / 1000, 1, MidpointRounding.AwayFromZero)

                    mThisYearVal = Math.Round(mAvgTareWeight(idx), 1, MidpointRounding.AwayFromZero)

                    mExcelSheet.Cells("G" & (idx + 7).ToString).Value = Format(mLastYearVal, "#.#")
                    mExcelSheet.Cells("H" & (idx + 7).ToString).Value =
                        mExcelSheet.Cells("F" & (idx + 7).ToString).Value - mExcelSheet.Cells("G" & (idx + 7).ToString).Value
                Else
                    mExcelSheet.Cells("H" & (idx + 7).ToString).Value = "N/A"
                End If

                mCommand = New SqlCommand
                mCommand.Connection = gbl_SQLConnection
                mCommand.CommandType = CommandType.Text


                ' Update the values in the Trans Data table
                If mLine_no <> " " Then
                    Select Case Val(mLine_no)
                        Case 501 To 518

                            mStrSQL = Build_Count_Trans_SQL_Statement(cmb_URCSYear.Text, 900004, 41, mLine_no)

                            ' Load the Year combobox from the SQL database
                            OpenSQLConnection(Gbl_Controls_Database_Name)

                            mDatatable = New DataTable

                            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                                daAdapter.Fill(mDatatable)
                            End Using

                            If mDatatable.Rows(0)(0) = 0 Then
                                mStrSQL = Build_Insert_Trans_SQL_Field_Statement(cmb_URCSYear.Text, 900004, 41, mLine_no, 1,
                                          (Math.Round(mAvgTareWeight(idx), 1, MidpointRounding.AwayFromZero) * 1000).ToString)
                            Else
                                mStrSQL = Build_Update_Trans_SQL_Field_Statement(cmb_URCSYear.Text, 900004, 41, mLine_no, 1,
                                          (Math.Round(mAvgTareWeight(idx), 1, MidpointRounding.AwayFromZero) * 1000).ToString)
                            End If

                            ' Write the value to the East data if the line exists

                            mCommand = New SqlCommand
                            OpenSQLConnection(My.Settings.Controls_DB)
                            mCommand.Connection = gbl_SQLConnection
                            mCommand.CommandType = CommandType.Text
                            mCommand.CommandText = mStrSQL
                            mCommand.ExecuteNonQuery()

                            mStrSQL = Build_Count_Trans_SQL_Statement(cmb_URCSYear.Text, 900007, 41, mLine_no)

                            ' Load the Year combobox from the SQL database
                            OpenSQLConnection(Gbl_Controls_Database_Name)

                            mDatatable = New DataTable

                            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                                daAdapter.Fill(mDatatable)
                            End Using

                            If mDatatable.Rows(0)(0) = 0 Then
                                mStrSQL = Build_Insert_Trans_SQL_Field_Statement(cmb_URCSYear.Text, 900007, 41, mLine_no, 1,
                                          (Math.Round(mAvgTareWeight(idx), 1, MidpointRounding.AwayFromZero) * 1000).ToString)
                            Else
                                mStrSQL = Build_Update_Trans_SQL_Field_Statement(cmb_URCSYear.Text, 900007, 41, mLine_no, 1,
                                          (Math.Round(mAvgTareWeight(idx), 1, MidpointRounding.AwayFromZero) * 1000).ToString)
                            End If

                            ' Write the value to the West data
                            mCommand = New SqlCommand
                            OpenSQLConnection(My.Settings.Controls_DB)
                            mCommand.Connection = gbl_SQLConnection
                            mCommand.CommandType = CommandType.Text
                            mCommand.CommandText = mStrSQL
                            mCommand.ExecuteNonQuery()

                        Case 584, 585

                            mStrSQL = Build_Count_Trans_SQL_Statement(cmb_URCSYear.Text, 900099, 42, mLine_no)

                            ' Load the Year combobox from the SQL database
                            OpenSQLConnection(Gbl_Controls_Database_Name)

                            mDatatable = New DataTable

                            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                                daAdapter.Fill(mDatatable)
                            End Using

                            If mDatatable.Rows(0)(0) = 0 Then
                                mStrSQL = Build_Insert_Trans_SQL_Field_Statement(cmb_URCSYear.Text, 900099, 42, mLine_no, 1,
                                          (Math.Round(mAvgTareWeight(idx), 1, MidpointRounding.AwayFromZero) * 1000).ToString)
                            Else
                                mStrSQL = Build_Update_Trans_SQL_Field_Statement(cmb_URCSYear.Text, 900099, 42, mLine_no, 1,
                                          (Math.Round(mAvgTareWeight(idx), 1, MidpointRounding.AwayFromZero) * 1000).ToString)
                            End If

                            ' Write the data to the National data

                            mCommand = New SqlCommand
                            OpenSQLConnection(My.Settings.Controls_DB)
                            mCommand.Connection = gbl_SQLConnection
                            mCommand.CommandType = CommandType.Text
                            mCommand.CommandText = mStrSQL
                            mCommand.ExecuteNonQuery()

                    End Select

                End If

            Next

            mExcelSheet.PageSetup.PrintArea = mExcelSheet.Cells("A1:H36").ToString
            mExcelSheet.PageSetup.Orientation = PageOrientation.Landscape
            mExcelSheet.PageSetup.FitToPages = True

            mWorkbook.SaveAs(txt_Report_FilePath.Text, FileFormat.OpenXMLWorkbook)

        End If
EndIt:

        Me.txt_StatusBox.Text = "Done!"

    End Sub

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close the Tare Weight Loader Form
        Me.Close()
    End Sub

    Private Sub frm_Tare_Weight_Loader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        CenterToScreen()

        ' Load the Year combobox from the SQL database
        OpenSQLConnection(Gbl_Controls_Database_Name)

        mDataTable = Get_URCS_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            mDataTable.Rows(mLooper)("urcs_year") = Trim(mDataTable.Rows(mLooper)("urcs_year"))
            cmb_URCSYear.Items.Add(mDataTable.Rows(mLooper)("urcs_year"))
        Next

        mDataTable = Nothing

    End Sub


    Private Sub btn_Input_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog
        Dim mLocation As Integer

        mLocation = 0
        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_FilePath.Text = fd.FileName
            mLocation = txt_Input_FilePath.Text.LastIndexOf("\")
            txt_Report_FilePath.Text = Mid(txt_Input_FilePath.Text, 1, mLocation) & "\WB" & cmb_URCSYear.Text & " Tare Weight Load Report.xlsx"
        End If

    End Sub

    Private Sub btn_Report_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Report_File_Entry.Click
        Dim fd As New FolderBrowserDialog

        If cmb_URCSYear.Text = "" Then
            MsgBox("You must select a year first.", vbOKOnly, "Error!")
        Else
            fd.Description = "Select the location in which you want the output report placed."

            If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txt_Report_FilePath.Text = fd.SelectedPath.ToString & "\WB" & cmb_URCSYear.Text & " Tare Weight Load Report.xlsx"
            End If
        End If

    End Sub
End Class