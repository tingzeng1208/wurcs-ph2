Imports System.Data.SqlClient
Imports SpreadsheetGear

Public Class R1_Data_Load

    'Last changed:  5/7/2013
    'Purpose:       Loads R-1 Data from Excel Spreadsheet file provided by AAR
    'Author:        Michael R. Sanders
    'Tested:        Results confirmed as good by Jackie Bienko on 5/7/2013

    'Last changed:  8/22/2019
    'Purpose:       Added Load Report Creation
    'Author:        Michael R. Sanders

    'Last changed:  3/5/2021
    'Purpose:       Added TransModification logic from Phase 2 with corrections
    'Author:        Michael R. Sanders

    'Last changed:  12/23/2021
    'Purpose:       Updated for Schedule 700 correction.
    'Author:        Michael R. Sanders

    Private Sub btn_Input_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub btn_Output_File_Entry_Click(sender As System.Object, e As System.EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New FolderBrowserDialog

        If cmb_URCS_Year.Text = "" Then
            MsgBox("You must select a year first.", vbOKOnly, "Error!")
        Else
            fd.Description = "Select the location in which you want the output report placed."

            If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txt_Output_FilePath.Text = fd.SelectedPath.ToString & "\WB" & cmb_URCS_Year.Text & " R-1 Load Report.xlsx"
            End If
        End If
    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        Dim mWorkbookSet As SpreadsheetGear.IWorkbookSet
        Dim mWorkbook As SpreadsheetGear.IWorkbook
        Dim mWorksheet As SpreadsheetGear.IWorksheet
        Dim mCells As SpreadsheetGear.IRange

        Dim mSQLCmd As New SqlCommand

        Dim xclTable As DataTable
        Dim mDataTable As DataTable

        Dim bolWrite As Integer, mRec As Integer, mWorkNum As Integer, mRealRecords As Integer
        Dim StrSQL As String
        Dim mRRICC As Single, mSch As String
        Dim SchedArray() As String, SchArray() As String
        Dim mZeroValueRow As Boolean, mSkippedRecs As Integer

        mZeroValueRow = True
        mSkippedRecs = 0

        If cmb_URCS_Year.Text = "" Then
            MsgBox("You must select a Year value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        If txt_Input_FilePath.TextLength = 0 Then
            MsgBox("You must select an Input File value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        ' ask the user if he wants to load x records for this RRID
        bolWrite = MsgBox("Are you sure you want to Write this data?", vbYesNo)

        If bolWrite = vbYes Then

            'Get the table name and database name for the Trans table and the table names for the AuditLog and URCS codes
            Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "TRANS")
            Gbl_AuditTrailLog_Tablename = Get_Table_Name_From_SQL("1", "ACTIVITYAUDITLOG")
            Gbl_URCS_Codes_TableName = Get_Table_Name_From_SQL("1", "URCS_Codes")

            'Open the connection to the trans database
            OpenSQLConnection(Gbl_Controls_Database_Name)

            'Wipe the road submitted data from the Trans table, keeping the national and region data.
            StrSQL = "DELETE " & Gbl_Trans_TableName & " WHERE year = " & cmb_URCS_Year.Text & " AND RRICC < 900000"
            mSQLCmd.CommandType = CommandType.Text
            mSQLCmd.Connection = gbl_SQLConnection
            mSQLCmd.CommandText = StrSQL
            mSQLCmd.ExecuteNonQuery()

            ' redim the arrays to match the number of URCS Codes/Schedules

            OpenSQLConnection(Gbl_Controls_Database_Name)

            mSQLCmd = New SqlCommand
            mSQLCmd.CommandType = CommandType.Text
            mSQLCmd.Connection = gbl_SQLConnection
            mSQLCmd.CommandText = "SELECT Count(*) FROM " & Gbl_URCS_Codes_TableName
            mRec = mSQLCmd.ExecuteScalar

            ReDim SchedArray(mRec)
            ReDim SchArray(mRec)

            ' Load the values from the URCSCodes table
            mDataTable = New DataTable
            StrSQL = "Select * From " & Gbl_URCS_Codes_TableName

            Using daAdapter As New SqlDataAdapter(StrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            For mWorkNum = 0 To mDataTable.Rows.Count - 1
                If mDataTable.Rows(mWorkNum)("sch") Is DBNull.Value Then
                    SchArray(mWorkNum) = "0"
                Else
                    SchArray(mWorkNum) = mDataTable.Rows(mWorkNum)("sch")
                End If
                If mDataTable.Rows(mWorkNum)("schedule") Is DBNull.Value Then
                    SchedArray(mWorkNum) = "0"
                Else
                    SchedArray(mWorkNum) = mDataTable.Rows(mWorkNum)("schedule")
                End If
            Next

            'Open the connection to the trans database for this year
            OpenSQLConnection(My.Settings.Controls_DB)

            ' Open the Excel file and establish the table from Sheet1
            mWorkbook = Factory.GetWorkbook(txt_Input_FilePath.Text)
            mWorksheet = mWorkbook.Worksheets(0)
            mCells = mWorksheet.UsedRange
            xclTable = mCells.GetDataTable(Data.GetDataFlags.None)


            mRealRecords = 1
            mRec = 1

            For Each xclRow As DataRow In xclTable.Rows

                If Not IsNothing(xclRow("RRID")) And xclRow("RRID").ToString <> "" Then
                    mRealRecords += 1

                    If mRec Mod 10 = 0 Then
                        txt_StatusBox.Text = "Processing record " & CStr(mRec) & "..."
                        Refresh()
                        Application.DoEvents()
                    End If

                    ' Change RRID Column to standard RRICC code
                    Select Case xclRow("RRID")
                        Case 3050
                            mRRICC = "130500"      'BNSF
                        Case 1370
                            mRRICC = "114900"      'CNGT
                        Case 2670
                            mRRICC = "125600"      'CSX
                        Case 3410
                            mRRICC = "134500"      'KCS
                        Case 1550
                            mRRICC = "117000"      'NS
                        Case 3680
                            mRRICC = "137700"      'SOO/CP
                        Case 3740
                            mRRICC = "139300"      'UP
                        Case Else
                            mRRICC = "0"
                    End Select

                    ' Check to ensure that values are not null
                    If xclRow("ColA") Is DBNull.Value Then
                        xclRow("ColA") = 0
                    End If

                    If xclRow("ColB") Is DBNull.Value Then
                        xclRow("ColB") = 0
                    End If

                    If xclRow("ColC") Is DBNull.Value Then
                        xclRow("ColC") = 0
                    End If

                    If xclRow("ColD") Is DBNull.Value Then
                        xclRow("ColD") = 0
                    End If

                    If xclRow("ColE") Is DBNull.Value Then
                        xclRow("ColE") = 0
                    End If

                    If xclRow("ColF") Is DBNull.Value Then
                        xclRow("ColF") = 0
                    End If

                    If xclRow("ColG") Is DBNull.Value Then
                        xclRow("ColG") = 0
                    End If

                    If xclRow("ColH") Is DBNull.Value Then
                        xclRow("ColH") = 0
                    End If

                    If xclRow("ColI") Is DBNull.Value Then
                        xclRow("ColI") = 0
                    End If

                    If xclRow("ColJ") Is DBNull.Value Then
                        xclRow("ColJ") = 0
                    End If

                    If xclRow("ColK") Is DBNull.Value Then
                        xclRow("ColK") = 0
                    End If

                    If xclRow("ColL") Is DBNull.Value Then
                        xclRow("ColL") = 0
                    End If

                    If xclRow("ColM") Is DBNull.Value Then
                        xclRow("ColM") = 0
                    End If

                    If xclRow("ColN") Is DBNull.Value Then
                        xclRow("ColN") = 0
                    End If

                    ' If the record contains all zeros, skip this record
                    If xclRow("ColA") + xclRow("ColB") + xclRow("ColC") + xclRow("ColD") + xclRow("ColE") + xclRow("ColF") +
                        xclRow("ColG") + xclRow("ColH") + xclRow("ColI") + xclRow("ColJ") + xclRow("ColK") + xclRow("ColL") +
                        xclRow("ColM") + xclRow("ColN") = 0 Then
                        mSkippedRecs = mSkippedRecs + 1
                        GoTo SkipIt
                    End If


                    ' Adjust Schedule number, if necessary
                    Select Case xclRow("Sched")
                        Case "0352"
                            mSch = "352A"
                        Case "0353"
                            mSch = "352B"
                        Case Else
                            mSch = xclRow("Sched").ToString
                    End Select

                    ' Find/Convert the Schedule to Trans Schedule number
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.CommandText = "SELECT SCH FROM " & Gbl_URCS_Codes_TableName & " WHERE convert(varchar,Schedule) = '" & mSch & "'"
                    mSch = mSQLCmd.ExecuteScalar

                    If mSch > 0 Then

                        mSQLCmd = New SqlCommand
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.CommandText = Build_Count_Trans_SQL_Statement(xclRow("Year"),
                            mRRICC,
                            Val(mSch),
                            xclRow("Line"))

                        If mSQLCmd.ExecuteScalar = 0 Then
                            ' We must accomodate some inconsistencies in the AAR Excel file
                            ' C1 in the database is actually ColB in the file except for Sch 700 (starts in ColC) and Sch 725 (Starts in ColA)

                            StrSQL = ""

                            Select Case mSch
                                Case 33         'Schedule 700 in input file.
                                    StrSQL = Build_Insert_Trans_SQL_Record_Statement(xclRow("Year"),
                                        mRRICC,
                                        Val(mSch),
                                        xclRow("Line"),
                                        xclRow("ColC"),
                                        xclRow("ColD"),
                                        xclRow("ColE"),
                                        xclRow("ColF"),
                                        xclRow("ColG"),
                                        xclRow("ColH"),
                                        xclRow("ColI"),
                                        xclRow("ColJ"),
                                        xclRow("ColK"),
                                        xclRow("ColL"),
                                        xclRow("ColM"),
                                        xclRow("ColN")).ToString
                                Case 725
                                    ' A record for this row does not exist
                                    ' Add the current Row values to the SQL table
                                    StrSQL = Build_Insert_Trans_SQL_Record_Statement(xclRow("Year"),
                                        mRRICC,
                                        Val(mSch),
                                        xclRow("Line"),
                                        xclRow("ColA"),
                                        xclRow("ColB"),
                                        xclRow("ColC"),
                                        xclRow("ColD"),
                                        xclRow("ColE"),
                                        xclRow("ColF"),
                                        xclRow("ColG"),
                                        xclRow("ColH"),
                                        xclRow("ColI"),
                                        xclRow("ColJ"),
                                        xclRow("ColK"),
                                        xclRow("ColL"),
                                        xclRow("ColM"),
                                        xclRow("ColN")).ToString
                                Case Else
                                    ' A record for this row does not exist
                                    ' Add the current Row values to the SQL table
                                    StrSQL = Build_Insert_Trans_SQL_Record_Statement(xclRow("Year"),
                                        mRRICC,
                                        Val(mSch),
                                        xclRow("Line"),
                                        xclRow("ColB"),
                                        xclRow("ColC"),
                                        xclRow("ColD"),
                                        xclRow("ColE"),
                                        xclRow("ColF"),
                                        xclRow("ColG"),
                                        xclRow("ColH"),
                                        xclRow("ColI"),
                                        xclRow("ColJ"),
                                        xclRow("ColK"),
                                        xclRow("ColL"),
                                        xclRow("ColM"),
                                        xclRow("ColN")).ToString
                            End Select
                        Else

                            ' We must accomodate some inconsistencies in the AAR Excel file
                            ' C1 in the database is actually ColB in the file except for Sch 700 (Starts in ColC) and Sch 725 (Starts in ColA)

                            Select Case mSch
                                Case 33     '700 in input file
                                    StrSQL = Build_Update_Trans_SQL_Record_Statement(xclRow("Year"),
                                        mRRICC,
                                        Val(mSch),
                                        xclRow("Line"),
                                        xclRow("ColC"),
                                        xclRow("ColD"),
                                        xclRow("ColE"),
                                        xclRow("ColF"),
                                        xclRow("ColG"),
                                        xclRow("ColH"),
                                        xclRow("ColI"),
                                        xclRow("ColJ"),
                                        xclRow("ColK"),
                                        xclRow("ColL"),
                                        xclRow("ColM"),
                                        xclRow("ColN")).ToString
                                Case 725
                                    'A record for this row does exist
                                    'Update the existing record
                                    StrSQL = Build_Update_Trans_SQL_Record_Statement(xclRow("Year"),
                                        mRRICC,
                                        Val(mSch),
                                        xclRow("Line"),
                                        xclRow("ColA"),
                                        xclRow("ColB"),
                                        xclRow("ColC"),
                                        xclRow("ColD"),
                                        xclRow("ColE"),
                                        xclRow("ColF"),
                                        xclRow("ColG"),
                                        xclRow("ColH"),
                                        xclRow("ColI"),
                                        xclRow("ColJ"),
                                        xclRow("ColK"),
                                        xclRow("ColL"),
                                        xclRow("ColM"),
                                        xclRow("ColN")).ToString
                                Case Else
                                    'A record for this row does exist
                                    'Update the existing record
                                    StrSQL = Build_Update_Trans_SQL_Record_Statement(xclRow("Year"),
                                        mRRICC,
                                        Val(mSch),
                                        xclRow("Line"),
                                        xclRow("ColB"),
                                        xclRow("ColC"),
                                        xclRow("ColD"),
                                        xclRow("ColE"),
                                        xclRow("ColF"),
                                        xclRow("ColG"),
                                        xclRow("ColH"),
                                        xclRow("ColI"),
                                        xclRow("ColJ"),
                                        xclRow("ColK"),
                                        xclRow("ColL"),
                                        xclRow("ColM"),
                                        xclRow("ColN")).ToString
                            End Select
                        End If

                        ' Execute the SQL command
                        mSQLCmd = New SqlCommand
                        mSQLCmd.Connection = gbl_SQLConnection
                        mSQLCmd.CommandType = CommandType.Text
                        mSQLCmd.CommandText = StrSQL
                        mSQLCmd.ExecuteNonQuery()

                        'If Sch 100 to 148 then we have to update C12, C13, and C14
                        If Val(mSch) > 99 And Val(mSch) < 149 Then

                            'Update C12
                            StrSQL = Build_Update_Trans_SQL_Field_Statement(cmb_URCS_Year.Text,
                                        mRRICC,
                                        mSch,
                                        xclRow("Line"),
                                        12,
                                        (xclRow("ColB") * 2) + xclRow("ColD") + xclRow("ColF"))

                            mSQLCmd.CommandText = StrSQL
                            mSQLCmd.ExecuteNonQuery()

                            'Update C13
                            StrSQL = Build_Update_Trans_SQL_Field_Statement(cmb_URCS_Year.Text,
                                        mRRICC,
                                        mSch,
                                        xclRow("Line"),
                                        13,
                                        xclRow("ColB") + xclRow("ColD") + xclRow("ColF") + xclRow("ColH"))

                            mSQLCmd.CommandText = StrSQL
                            mSQLCmd.ExecuteNonQuery()

                            'Update C14
                            StrSQL = Build_Update_Trans_SQL_Field_Statement(cmb_URCS_Year.Text,
                                        mRRICC,
                                        mSch,
                                        xclRow("Line"),
                                        14,
                                        xclRow("ColD") + xclRow("ColF") + (xclRow("ColH") * 2))

                            mSQLCmd.CommandText = StrSQL
                            mSQLCmd.ExecuteNonQuery()
                        End If

                        'If Sch 420 then we have to update C12
                        If Val(mSch) = 420 Then
                            StrSQL = Build_Update_Trans_SQL_Field_Statement(cmb_URCS_Year.Text,
                                        mRRICC,
                                        mSch,
                                        xclRow("Line"),
                                        12,
                                        xclRow("ColC") + xclRow("ColD"))

                            mSQLCmd.CommandText = StrSQL
                            mSQLCmd.ExecuteNonQuery()
                        End If

                        'If Sch 33 and Line 57 then we have to update C8, C9, and C10
                        If Val(mSch) = 33 And xclRow("Line") = 57 Then

                            'Update C8
                            StrSQL = Build_Update_Trans_SQL_Field_Statement(cmb_URCS_Year.Text,
                                        mRRICC,
                                        mSch,
                                        xclRow("Line"),
                                        8,
                                        xclRow("ColC") + xclRow("ColD") + xclRow("ColE") + xclRow("ColF"))

                            mSQLCmd.CommandText = StrSQL
                            mSQLCmd.ExecuteNonQuery()

                            'Update C9
                            StrSQL = Build_Update_Trans_SQL_Field_Statement(cmb_URCS_Year.Text,
                                        mRRICC,
                                        mSch,
                                        xclRow("Line"),
                                        9,
                                        xclRow("ColG") + xclRow("ColH"))

                            mSQLCmd.CommandText = StrSQL
                            mSQLCmd.ExecuteNonQuery()

                            'Update C10
                            StrSQL = Build_Update_Trans_SQL_Field_Statement(cmb_URCS_Year.Text,
                                        mRRICC,
                                        mSch,
                                        xclRow("Line"),
                                        10,
                                        xclRow("ColC") + xclRow("ColD") + xclRow("ColE") + xclRow("ColF") + xclRow("ColG"))

                            mSQLCmd.CommandText = StrSQL
                            mSQLCmd.ExecuteNonQuery()
                        End If
                    End If
                End If

SkipIt:

                mRec = mRec + 1
                Me.Refresh()

            Next

            'We need to change some possible negative values to positive in SCH 412

            Remove_Negative_Values()

        Else
            GoTo EndIt
        End If

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
        mWorksheet = mWorkbook.ActiveSheet
        mWorksheet.Name = "Summary"

        ' Set the data to the Summary page
        mCells = mWorksheet.Cells

        mCells("A1").Value = "URCS R-1 Load Process Report"
        mCells("A3").Value = "Input File:"
        mCells("A4").Value = "Date/Time:"

        mCells("B3").Value = txt_Input_FilePath.Text
        mCells("B4").Value = CStr(Now)

        mCells("A1").Font.Bold = True
        mCells("A1").Font.Size = 14

        mCells("A6").Value = "Records Processed: " & mRealRecords.ToString
        mCells("A7").Value = "Zero Value Records: " & mSkippedRecs.ToString

        mWorkbook.SaveAs(txt_Output_FilePath.Text, FileFormat.OpenXMLWorkbook)

        txt_StatusBox.Text = "Done! Processed " & mRealRecords.ToString & " Records - Skipped " & mSkippedRecs.ToString & " Zero Value Records"
        Refresh()

EndIt:

    End Sub

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close the Tare Weight Loader Form
        Me.Close()
    End Sub

    Private Sub R1_Data_Load_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Sub Remove_Negative_Values()

        Dim mSQLCmd As New SqlCommand
        Dim mTable As DataTable
        Dim mStrSQL As String
        Dim mLooper As Integer, mLooper2 As Integer
        Dim rtfColumns(4) As String

        txt_StatusBox.Text = "Converting Schedule 412 Negative Values..."
        Refresh()

        Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "TRANS")

        'Open the connection to the trans database
        OpenSQLConnection(Gbl_Controls_Database_Name)

        ' Store the SQL statement
        mStrSQL = "Select C1, C2, C3, C4, C5, C6, C7, YEAR, RRICC, LINE FROM " & Gbl_Trans_TableName &
            " WHERE SCH = 412 AND YEAR = " & cmb_URCS_Year.Text & " ORDER BY rricc, line ASC"

        ' Get the records from SQL Server
        mTable = New DataTable

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mTable)
        End Using

        For mLooper = 0 To mTable.Rows.Count - 1
            Select Case mTable.Rows(mLooper)("line")
                Case 121 To 123, 127 To 129, 133 To 135, 142 To 144,
                    208, 210, 212, 215, 216, 227, 229, 231, 234,
                    235, 312, 314, 316, 319, 320, 417, 433, 515, 525, 617
                    For mLooper2 = 1 To 7
                        If mTable.Rows(mLooper)("C" & mLooper2.ToString) < 0 Then
                            mStrSQL = "UPDATE " & Trim(Gbl_Trans_TableName) &
                                " SET " & ("C" & mLooper2.ToString) & " = " & Math.Abs(mTable.Rows(mLooper)("C" & mLooper2.ToString)) &
                                " WHERE YEAR = " & mTable.Rows(mLooper)("year") & " AND " &
                                " RRICC = " & mTable.Rows(mLooper)("RRICC") & " AND " &
                                " SCH = 412 AND LINE = " & mTable.Rows(mLooper)("line")
                            mSQLCmd = New SqlCommand
                            mSQLCmd.CommandType = CommandType.Text
                            mSQLCmd.Connection = gbl_SQLConnection
                            mSQLCmd.CommandText = mStrSQL
                            mSQLCmd.ExecuteNonQuery()

                        End If
                    Next
            End Select
        Next mLooper

    End Sub

    Private Sub ModifyTransData()

        Dim mProcessYear As Integer

        For mProcessYear = (CInt(cmb_URCS_Year.Text) - 4) To CInt(cmb_URCS_Year.Text)

        Next

    End Sub

End Class