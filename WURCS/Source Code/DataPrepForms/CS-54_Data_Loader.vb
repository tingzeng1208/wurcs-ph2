'Imports Excel = Microsoft.Office.Interop.Excel - No longer used.  10/27/20 M.Sanders
'Imports System.Runtime.InteropServices - No longer used.  10/27/20 M.Sanders
Imports System.Data.SqlClient
Imports SpreadsheetGear

Public Class CS54_Data_Loader

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button input file entry click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Input_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_FilePath.Text = fd.FileName
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button return to data prep menu click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close this form
        Me.Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Create structs structure 54 data loader load. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub CS54_Data_Loader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        mDataTable = Get_URCS_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            cmb_URCS_Year.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
        Next

        mDataTable = Nothing

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button execute click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        ' Variables for writing to using SpreadsheetGear
        Dim mWorkbookSet As IWorkbookSet
        Dim mWorkbook As IWorkbook
        Dim mExcelSheet As IWorksheet
        Dim mRange As IRange

        Dim mTable As DataTable
        Dim cmdCommand As SqlCommand
        Dim xclTable As DataTable

        Dim mRailroad As Long, mLine As Integer, mWorksheet As Integer
        Dim mStrSQL As String, mWorkStr As String
        Dim mWorkVal As Integer, mLooper As Integer, mSheetNames As Integer
        Dim TotalRailroad As Decimal, TotalPrivate As Decimal, Total As Decimal
        Dim TotalLoading As Decimal, TotalTermination As Decimal
        Dim RRCarsLoaded As Decimal, PrivateCarsLoaded As Decimal
        Dim RRCarsTerminated As Decimal, PrivateCarsTerminated As Decimal

        ' Declare the arrays we'll need for handling the 36 lines of data
        ' for each of the 7 railroads
        Dim mDataline(7, 36) As Long
        Dim mSheets(10) As String
        Dim mLine30(6) As Decimal
        Dim mLine31(6) As Decimal
        Dim mLine35(6) As Decimal
        Dim mLine36(6) As Decimal
        Dim mValues(9) As Decimal
        Dim mCarTypes(36) As String

        Dim mSQLCmd As String

        ' Perform Error checking for form controls
        If Me.cmb_URCS_Year.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo EndIt
        End If

        If Me.txt_Input_FilePath.Text = "" Then
            MsgBox("You must select an input file.", vbOKOnly)
            GoTo EndIt
        End If

        If Me.txt_Output_FilePath.Text = "" Then
            MsgBox("You must select an output file.", vbOKOnly)
            GoTo EndIt
        End If

        ' If the output file exists verify that the user wants to overwrite it.  If not, end this.
        If My.Computer.FileSystem.FileExists(txt_Output_FilePath.Text) Then
            If MsgBox("File already exists!  Overwrite?", vbYesNo, "Warning!") = vbYes Then
                My.Computer.FileSystem.DeleteFile(txt_Output_FilePath.Text)
            Else
                GoTo EndIt
            End If
        End If

        ' Let the user know it is working
        txt_StatusBox.Text = "Working..."
        Refresh()

        'Load the Arrays so that we can use the Class1Railroads and
        'URCSCodes arrays
        LoadArrayData()

        ' Make sure we know where the Trans table is located
        Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "TRANS")

        'Initialize the Data line array just in case
        For mRailroad = 1 To 7
            For mLine = 1 To 36
                mDataline(mRailroad, mLine) = 0
                mCarTypes(mLine) = ""
            Next mLine
        Next mRailroad

        For mWorksheet = 1 To 10
            mSheets(mWorksheet) = ""
        Next mWorksheet

        For mLooper = 1 To 6
            mLine30(mLooper) = 0
            mLine31(mLooper) = 0
            mLine35(mLooper) = 0
            mLine36(mLooper) = 0
        Next mLooper

        For mLooper = 1 To 9
            mValues(mLooper) = 0
        Next mLooper

        'establish the connection to the spreadsheet file

        mWorkbookSet = SpreadsheetGear.Factory.GetWorkbookSet
        mWorkbook = mWorkbookSet.Workbooks.Open(txt_Input_FilePath.Text)

        For mSheetNames = 0 To mWorkbook.Worksheets.Count - 1
            ' Check to see the name of the sheet.  If it is Summary, skip it
            If mWorkbook.Worksheets(mSheetNames).Name <> "Summary" Then

                ' Set the work environment up for this sheet.
                For mLooper = 1 To 6
                    mLine30(mLooper) = 0
                    mLine31(mLooper) = 0
                    mLine35(mLooper) = 0
                    mLine36(mLooper) = 0
                Next mLooper
                For mLooper = 1 To 9
                    mValues(mLooper) = 0
                Next mLooper

                mWorkStr = Mid(mWorkbook.Worksheets(mSheetNames).Name, 1, 4)
                mWorkStr = Replace(mWorkStr, " ", "")
                mWorkStr = Replace(mWorkStr, "(", "")
                mWorkStr = Replace(mWorkStr, "-", "")

                If mWorkStr = "GTC" Then
                    mWorkStr = "CN"
                End If

                mWorkVal = Get_RRICC_By_Short_Name(mWorkStr)

                ' Let the user know it is working
                txt_StatusBox.Text = "Storing " & mWorkStr & " data to SQL..."
                Refresh()

                mRange = mWorkbook.Worksheets(mSheetNames).Cells("A16:J45")
                xclTable = New DataTable
                xclTable = mRange.GetDataTable(Data.GetDataFlags.NoColumnHeaders)

                For Each xclRow As DataRow In xclTable.Rows
                    'Get the line number
                    mLine = xclRow.Item(0)

                    'Store the data from the line
                    mCarTypes(mLine) = xclRow.Item(1)
                    RRCarsLoaded = ReturnValidNumber(xclRow.Item(2))
                    PrivateCarsLoaded = ReturnValidNumber(xclRow.Item(3))
                    RRCarsTerminated = ReturnValidNumber(xclRow.Item(7))
                    PrivateCarsTerminated = ReturnValidNumber(xclRow.Item(8))

                    'Calculate the totals
                    TotalRailroad = RRCarsLoaded + RRCarsTerminated
                    TotalPrivate = PrivateCarsLoaded + PrivateCarsTerminated
                    TotalLoading = RRCarsLoaded + PrivateCarsLoaded
                    TotalTermination = RRCarsTerminated + PrivateCarsTerminated
                    Total = TotalRailroad + TotalPrivate

                    'Check to see if a record exists in the Trans table

                    OpenSQLConnection(My.Settings.Controls_DB)

                    cmdCommand = New SqlCommand
                    cmdCommand.CommandType = CommandType.Text
                    cmdCommand.Connection = gbl_SQLConnection
                    cmdCommand.CommandText = Build_Count_Trans_SQL_Statement(
                        cmb_URCS_Year.Text,
                        mWorkVal,
                        54,
                        mLine)

                    If cmdCommand.ExecuteScalar > 0 Then
                        mStrSQL = Build_Update_Trans_SQL_Record_Statement(
                            cmb_URCS_Year.Text,
                            mWorkVal,
                            54,
                            mLine,
                            RRCarsLoaded,
                            PrivateCarsLoaded,
                            TotalLoading,
                            RRCarsTerminated,
                            PrivateCarsTerminated,
                            TotalTermination,
                            TotalRailroad,
                            TotalPrivate,
                            Total).ToString

                    Else
                        'The record needs to be inserted
                        mStrSQL = Build_Insert_Trans_SQL_Record_Statement(
                            cmb_URCS_Year.Text,
                            mWorkVal,
                            54,
                            mLine,
                            RRCarsLoaded,
                            PrivateCarsLoaded,
                            TotalLoading,
                            RRCarsTerminated,
                            PrivateCarsTerminated,
                            TotalTermination,
                            TotalRailroad,
                            TotalPrivate,
                            Total).ToString
                    End If

                    cmdCommand = New SqlCommand
                    cmdCommand.Connection = gbl_SQLConnection
                    cmdCommand.CommandText = mStrSQL.ToString
                    cmdCommand.ExecuteNonQuery()

                    Select Case mLine
                        Case 2, 3 ' Box Cars
                            mLine31(1) = mLine31(1) + RRCarsLoaded
                            mLine31(2) = mLine31(2) + PrivateCarsLoaded
                            mLine31(3) = mLine31(3) + TotalLoading
                            mLine31(4) = mLine31(4) + RRCarsTerminated
                            mLine31(5) = mLine31(5) + PrivateCarsTerminated
                            mLine31(6) = mLine31(6) + TotalTermination
                        Case 6, 11 ' Equipped Box Cars
                            mLine36(1) = mLine36(1) + RRCarsLoaded
                            mLine36(2) = mLine36(2) + PrivateCarsLoaded
                            mLine36(3) = mLine36(3) + TotalLoading
                            mLine36(4) = mLine36(4) + RRCarsTerminated
                            mLine36(5) = mLine36(5) + PrivateCarsTerminated
                            mLine36(6) = mLine36(6) + TotalTermination
                        Case 7, 10, 14, 19, 22, 27 To 29 ' All Others
                            mLine30(1) = mLine30(1) + RRCarsLoaded
                            mLine30(2) = mLine30(2) + PrivateCarsLoaded
                            mLine30(3) = mLine30(3) + TotalLoading
                            mLine30(4) = mLine30(4) + RRCarsTerminated
                            mLine30(5) = mLine30(5) + PrivateCarsTerminated
                            mLine30(6) = mLine30(6) + TotalTermination
                        Case 15, 17 ' Gondolas
                            mLine35(1) = mLine35(1) + RRCarsLoaded
                            mLine35(2) = mLine35(2) + PrivateCarsLoaded
                            mLine35(3) = mLine35(3) + TotalLoading
                            mLine35(4) = mLine35(4) + RRCarsTerminated
                            mLine35(5) = mLine35(5) + PrivateCarsTerminated
                            mLine35(6) = mLine35(6) + TotalTermination
                    End Select

                Next xclRow

                ' Write out lines 30, 31, 35 and 36 to the trans file, checking for existing recs
                For mLooper = 1 To 4
                    Select Case mLooper
                        Case 1 ' All Others
                            mLine = 30
                            mValues(1) = mLine30(1)
                            mValues(2) = mLine30(2)
                            mValues(3) = mLine30(3)
                            mValues(4) = mLine30(4)
                            mValues(5) = mLine30(5)
                            mValues(6) = mLine30(6)
                            mValues(7) = mLine30(1) + mLine30(4)
                            mValues(8) = mLine30(2) + mLine30(5)
                            mValues(9) = mValues(7) + mValues(8)
                        Case 2 ' Box Cars
                            mLine = 31
                            mValues(1) = mLine31(1)
                            mValues(2) = mLine31(2)
                            mValues(3) = mLine31(3)
                            mValues(4) = mLine31(4)
                            mValues(5) = mLine31(5)
                            mValues(6) = mLine31(6)
                            mValues(7) = mLine31(1) + mLine31(4)
                            mValues(8) = mLine31(2) + mLine31(5)
                            mValues(9) = mValues(7) + mValues(8)
                        Case 3 ' Gondolas
                            mLine = 35
                            mValues(1) = mLine35(1)
                            mValues(2) = mLine35(2)
                            mValues(3) = mLine35(3)
                            mValues(4) = mLine35(4)
                            mValues(5) = mLine35(5)
                            mValues(6) = mLine35(6)
                            mValues(7) = mLine35(1) + mLine35(4)
                            mValues(8) = mLine35(2) + mLine35(5)
                            mValues(9) = mValues(7) + mValues(8)
                        Case 4 ' Equipped Box Cars
                            mLine = 36
                            mValues(1) = mLine36(1)
                            mValues(2) = mLine36(2)
                            mValues(3) = mLine36(3)
                            mValues(4) = mLine36(4)
                            mValues(5) = mLine36(5)
                            mValues(6) = mLine36(6)
                            mValues(7) = mLine36(1) + mLine36(4)
                            mValues(8) = mLine36(2) + mLine36(5)
                            mValues(9) = mValues(7) + mValues(8)

                    End Select

                    cmdCommand = New SqlCommand
                    cmdCommand.CommandType = CommandType.Text
                    cmdCommand.Connection = gbl_SQLConnection
                    cmdCommand.CommandText = Build_Count_Trans_SQL_Statement(
                        cmb_URCS_Year.Text,
                        mWorkVal,
                        54,
                        mLine)

                    If cmdCommand.ExecuteScalar > 0 Then
                        'A record exists for this railroad, schedule and line for the year
                        mStrSQL = Build_Update_Trans_SQL_Record_Statement(
                            cmb_URCS_Year.Text,
                            mWorkVal,
                            54,
                            mLine,
                            mValues(1),
                            mValues(2),
                            mValues(3),
                            mValues(4),
                            mValues(5),
                            mValues(6),
                            mValues(7),
                            mValues(8),
                            mValues(9)).ToString
                    Else
                        'we need to insert one
                        mStrSQL = Build_Insert_Trans_SQL_Record_Statement(
                            cmb_URCS_Year.Text,
                            mWorkVal,
                            54,
                            mLine,
                            mValues(1),
                            mValues(2),
                            mValues(3),
                            mValues(4),
                            mValues(5),
                            mValues(6),
                            mValues(7),
                            mValues(8),
                            mValues(9)).ToString
                    End If

                    cmdCommand = New SqlCommand
                    cmdCommand.Connection = gbl_SQLConnection
                    cmdCommand.CommandText = mStrSQL.ToString
                    cmdCommand.ExecuteNonQuery()

                Next

            End If

            mCarTypes(30) = "Total All Others"
            mCarTypes(31) = "Total Box Cars"
            mCarTypes(35) = "Total Gondolas"
            mCarTypes(36) = "Total Equipped Box Cars"

        Next

        mWorkbookSet.Workbooks.Close()
        mWorkbookSet = Nothing

        ' Now that we've stored all the information to SQL, we'll read it back to create the report

        ' Get the data, sorted by railroad and line

        ' Store the SQL statement
        mSQLCmd = "SELECT * FROM " & Get_Table_Name_From_SQL("1", "TRANS") & " " &
            "WHERE Year = " & cmb_URCS_Year.Text & " AND sch = 54 ORDER BY RRICC, Line"

        ' Get the records from SQL Server
        mTable = New DataTable
        Using daAdapter As New SqlDataAdapter(mSQLCmd, gbl_SQLConnection)
            daAdapter.Fill(mTable)
        End Using

        mRailroad = 1

        txt_StatusBox.Text = "Creating & Loading Excel File..."
        Refresh()

        ' Open the Excel file
        mWorkbookSet = Factory.GetWorkbookSet(Globalization.CultureInfo.CurrentCulture)
        mWorkbook = mWorkbookSet.Workbooks.Add

        ' Create sheet for this road
        If mWorkbook.ActiveSheet.Name <> "Sheet1" Then
            mWorkbook.Worksheets.Add()
        End If

        ' Set the active sheet's name
        mExcelSheet = mWorkbook.ActiveSheet
        mExcelSheet.Name = "Summary"

        ' Set the data to the Summary page
        mRange = mExcelSheet.Cells

        mRange("A1").Value = "URCS CS-54 Load Process Report"
        mRange("A3").Value = "Input File:"
        mRange("A4").Value = "Date/Time:"

        mRange("B3").Value = txt_Input_FilePath.Text
        mRange("B4").Value = CStr(Now)

        mRange("A1").Font.Bold = True
        mRange("A1").Font.Size = 14

        mRange.Columns("A:B").ColumnWidth = 20

        ' Now we create/write the sheet for each road/region
        For mLooper = 0 To mTable.Rows.Count - 1

            If mRailroad <> mTable.Rows(mLooper)("rricc") Then
                ' Either this is the first road or we've switched railroads
                ' and need to create a new sheet
                mRailroad = mTable.Rows(mLooper)("rricc")
                mWorkbook.Worksheets.Add()
                mExcelSheet = mWorkbook.ActiveSheet
                mExcelSheet.Name = Get_Short_Name_By_RRICC(mRailroad)
                mRange = mExcelSheet.Cells

                mRange("A1").Value = "URCS STB-54 Load Process Report"
                mRange("A3").Value = "Railroad:"
                mRange("A4").Value = "Date/Time: " & CStr(Now)

                mRange("B3").Value = mExcelSheet.Name.ToString

                mRange("A1").Font.Bold = True
                mRange("A1").Font.Size = 14

                mRange.Columns("A:B").ColumnWidth = 20

                mRange("A6").Value = "Car Types"
                mRange("B6").Value = "Line"
                mRange("C5").Value = "RR Cars"
                mRange("C6").Value = "Loaded"
                mRange("D5").Value = "Private Cars"
                mRange("D6").Value = "Loaded"
                mRange("E6").Value = "Total Loaded"
                mRange("F5").Value = "RR Cars"
                mRange("F6").Value = "Terminated"
                mRange("G5").Value = "Private Cars"
                mRange("G6").Value = "Terminated"
                mRange("H5").Value = "Total"
                mRange("H6").Value = "Terminated"
                mRange("I5").Value = "Total RR"
                mRange("I6").Value = "Cars"
                mRange("J5").Value = "Total"
                mRange("J6").Value = "Private Cars"
                mRange("K6").Value = "Total Cars"

                mRange("A5:K6").Font.Bold = True
                mRange("A6:K6").Font.Underline = True
                mRange("B5:K6").HorizontalAlignment = HAlign.Right

            End If

            If mTable.Rows(mLooper)("line") > 31 Then
                mWorkVal = mTable.Rows(mLooper)("line")
            Else
                mWorkVal = mTable.Rows(mLooper)("line") + 5
            End If

            mRange(mWorkVal, 0).Value = mCarTypes(mTable.Rows(mLooper)("line"))
            mRange(mWorkVal, 1).Value = mTable.Rows(mLooper)("line")
            mRange(mWorkVal, 2).Value = mTable.Rows(mLooper)("C1")
            mRange(mWorkVal, 3).Value = mTable.Rows(mLooper)("C2")
            mRange(mWorkVal, 4).Value = mTable.Rows(mLooper)("C3")
            mRange(mWorkVal, 5).Value = mTable.Rows(mLooper)("C4")
            mRange(mWorkVal, 6).Value = mTable.Rows(mLooper)("C5")
            mRange(mWorkVal, 7).Value = mTable.Rows(mLooper)("C6")
            mRange(mWorkVal, 8).Value = mTable.Rows(mLooper)("C7")
            mRange(mWorkVal, 9).Value = mTable.Rows(mLooper)("C8")
            mRange(mWorkVal, 10).Value = mTable.Rows(mLooper)("C9")

            ' This cleans up the columnwidth v. autofit problem
            mRange("A:K").Columns.AutoFit()
            For col As Integer = 0 To mExcelSheet.UsedRange.ColumnCount - 1
                mExcelSheet.Cells(1, col).ColumnWidth =
                    Math.Round(mExcelSheet.Cells(1, col).ColumnWidth, 0, MidpointRounding.AwayFromZero) + 2.5
            Next

            mExcelSheet.PageSetup.Orientation = PageOrientation.Landscape
            mExcelSheet.PageSetup.FitToPages = True

        Next

        ' Set the summary sheet as default when the file opens in Excel for the first time
        mWorkbook.Worksheets(0).Select()

        ' We're done!  Save it and close
        mWorkbook.SaveAs(txt_Output_FilePath.Text, FileFormat.OpenXMLWorkbook)
        mExcelSheet = Nothing
        mWorkbook = Nothing
        mWorkbookSet = Nothing

EndIt:

        Me.txt_StatusBox.Text = "Done!"
        Me.Refresh()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button output file entry click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Output_File_Entry_Click(sender As System.Object, e As System.EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New FolderBrowserDialog

        If cmb_URCS_Year.Text = "" Then
            MsgBox("You must select a year first.", vbOKOnly, "Error!")
        Else
            fd.Description = "Select the location in which you want the output report placed."

            If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txt_Output_FilePath.Text = fd.SelectedPath.ToString & "\WB" & cmb_URCS_Year.Text & " STB-54 Load Report.xlsx"
            End If
        End If
    End Sub
End Class