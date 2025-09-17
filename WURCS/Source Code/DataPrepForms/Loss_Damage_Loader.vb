Imports System.Data.SqlClient
Imports SpreadsheetGear
Public Class Loss_Damage_Loader

    'Last changed:  5/7/2013
    'Purpose:       Loads Loss & Damage Data from Excel Spreadheet file provided by AAR
    'Author:        Michael R. Sanders
    'Tested:        

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close this Form
        Me.Close()
    End Sub

    Private Sub Loss_Damage_Loader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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

        ' Variables for SpreadsheetGear
        Dim mWorkbookSet As IWorkbookSet
        Dim mWorkbook As IWorkbook
        Dim mRange As IRange

        Dim xclTable As DataTable
        Dim mDataTable As New DataTable
        Dim mSQLcmd As SqlCommand

        Dim mC1 As Decimal, mC2 As Decimal
        Dim mLine482C2 As Decimal
        Dim mStrSQL As String
        Dim mLine As Integer
        Dim mSch As Integer
        Dim mRailroad As Decimal

        ' Perform Error checking for form controls
        If Len(cmb_URCSYear.Text) = 0 Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo EndIt
        End If

        If txt_Input_FilePath.TextLength = 0 Then
            MsgBox("You must select an input file.", vbOKOnly)
            GoTo EndIt
        End If

        gbl_Database_Name = Get_Database_Name_From_SQL(cmb_URCSYear.Text, "MASKED")
        gbl_Table_Name = Get_Table_Name_From_SQL(cmb_URCSYear.Text, "MASKED")

        ' Verify that the Waybill data has been previously loaded for the year selected.

        OpenSQLConnection(My.Settings.Waybills_DB)

        mSQLcmd = New SqlCommand
        mSQLcmd.CommandType = CommandType.Text
        mSQLcmd.Connection = gbl_SQLConnection
        mSQLcmd.CommandText = "SELECT COUNT(*) FROM " & gbl_Table_Name

        If mSQLcmd.ExecuteScalar < 200000 Then
            MsgBox("You must load the Waybill data before processing the L&D data.", vbOKOnly)
            mDataTable = Nothing
            GoTo EndIt
        End If

        txt_StatusBox.Text = "Processing..."
        Refresh()

        Gbl_Trans_TableName = Trim(Get_Table_Name_From_SQL("1", "Trans")
)
        ' Extract the L&D data from the spreadsheet
        ' total lines for each of the commodities and save them to the
        ' trans table

        ' establish the connection to the spreadsheet file
        ' Open the Workbook

        mWorkbookSet = SpreadsheetGear.Factory.GetWorkbookSet
        mWorkbook = mWorkbookSet.Workbooks.Open(txt_Input_FilePath.Text)

        mRange = mWorkbook.Worksheets(0).Cells("A7:F75")

        ' Get a DataTable from the range ignoring the hidden rows.
        xclTable = mRange.GetDataTable(3)

        mSch = 60
        mRailroad = 900099
        mLine = 401

        ' Open the connection to the Trans file location
        OpenSQLConnection(My.Settings.Controls_DB)

        For Each xclRow As DataRow In xclTable.Rows

            txt_StatusBox.Text = "Processing record " & CStr(mLine)
            Refresh()

            'Write the data to the Trans table

            Select Case mLine
                ' if any we are at any of these lines, the values should be set to zero
                Case 406, 426, 431, 439, 445, 449, 453, 457, 460, 465, 470, 475, 480
                    mC1 = 0
                    mC2 = 0

                    'we need to check to see if the record exists in
                    'the trans table

                    'build the SQL select statement for railroad 900099, Sch 60, mline
                    mSQLcmd = New SqlCommand
                    mSQLcmd.CommandType = CommandType.Text
                    mSQLcmd.Connection = gbl_SQLConnection
                    mSQLcmd.CommandText = Build_Count_Trans_SQL_Statement(cmb_URCSYear.Text,
                        mRailroad,
                        mSch,
                        mLine)

                    If mSQLcmd.ExecuteScalar > 0 Then
                        'there is already a record there and we need
                        'to update it
                        mStrSQL = Build_Update_Trans_SQL_Record_Statement(
                            CInt(cmb_URCSYear.Text),
                            CInt(mRailroad),
                            mSch,
                            mLine,
                            mC1,
                            mC2).ToString
                    Else
                        'No record exists - we need to insert it
                        mStrSQL = Build_Insert_Trans_SQL_Record_Statement(
                            CInt(cmb_URCSYear.Text),
                            CInt(mRailroad),
                            mSch,
                            mLine,
                            mC1,
                            mC2).ToString
                    End If

                    mSQLcmd = New SqlCommand
                    mSQLcmd.Connection = gbl_SQLConnection
                    mSQLcmd.CommandType = CommandType.Text
                    mSQLcmd.CommandText = mStrSQL
                    mSQLcmd.ExecuteNonQuery()

                    mLine = mLine + 1

            End Select

            ' set the values from the excel file
            ' Note that they are intentionally reversed
            mC1 = CDec(ReturnValidNumber(xclRow(4)))
            mC2 = CDec(ReturnValidNumber(xclRow(3)))

            Select Case mLine
                'Case 481
                '    mLine482C1 = mC1
                '    mLine482C2 = mC2
                Case 482
                    'This is the Hazmat line and we need to save the C2
                    'value for use later
                    mLine482C2 = mC2
            End Select

            'we need to check to see if the record exists in
            'the trans table

            ' Line 482 in the Trans file should contain the values from line 481, so we have to switch them
            If mLine = 481 Then
                'build the SQL select statement for railroad 900099, Sch 60, Line 482
                mStrSQL = Build_Count_Trans_SQL_Statement(cmb_URCSYear.Text,
                            mRailroad,
                            mSch,
                            482)
            Else
                'build the SQL select statement for railroad 900099, Sch 60, mline
                mStrSQL = Build_Count_Trans_SQL_Statement(cmb_URCSYear.Text,
                            mRailroad,
                            mSch,
                            mLine)
            End If

            mSQLcmd = New SqlCommand
            mSQLcmd.Connection = gbl_SQLConnection
            mSQLcmd.CommandType = CommandType.Text
            mSQLcmd.CommandText = mStrSQL

            If mSQLcmd.ExecuteScalar > 0 Then
                'there is already a record there and we need
                'to update it
                mStrSQL = Build_Update_Trans_SQL_Record_Statement(
                    CInt(cmb_URCSYear.Text),
                    CInt(mRailroad),
                    mSch,
                    mLine,
                    mC1,
                    mC2).ToString
            Else
                'No record exists - we need to insert it
                mStrSQL = Build_Insert_Trans_SQL_Record_Statement(
                    CInt(cmb_URCSYear.Text),
                    CInt(mRailroad),
                    mSch,
                    mLine,
                    mC1,
                    mC2).ToString
            End If

            mSQLcmd = New SqlCommand
            mSQLcmd.Connection = gbl_SQLConnection
            mSQLcmd.CommandType = CommandType.Text
            mSQLcmd.CommandText = mStrSQL
            mSQLcmd.ExecuteNonQuery()

            mLine = mLine + 1

        Next

        ' Now we have to get the hazmat (STCC 48) expanded tons for the
        ' data from the waybills to be stored into C1 on line 482
        ' We should only sum the tons for BNSF, CSX, KCS, NS and UP
        ' as the ORR - No CP or CN traffic should be included.

        ' Note that this is an error in STCC selection - it should be STCC 49, but is pending rulemaking.

        Me.txt_StatusBox.Text = "Getting Tons data from Waybills..."
        Me.Refresh()

        gbl_Table_Name = Get_Table_Name_From_SQL(cmb_URCSYear.Text, "MASKED")

        ' Open the connection to the waybill file location
        OpenSQLConnection(My.Settings.Waybills_DB)

        mSQLcmd = New SqlCommand
        mSQLcmd.CommandType = CommandType.Text
        mSQLcmd.Connection = gbl_SQLConnection
        mSQLcmd.CommandText = "SELECT SUM(tons) as MyTons from " & gbl_Table_Name & " " &
            "WHERE ((stcc LIKE '48%') AND " &
            "(orr = 555 OR orr = 712 OR orr = 777 OR orr = 802 OR orr = 400))"

        mLine = 482
        mC1 = mSQLcmd.ExecuteScalar
        mC2 = mLine482C2

        ' Open the connection to the Trans file location
        OpenSQLConnection(My.Settings.Controls_DB)

        mSQLcmd = New SqlCommand
        mSQLcmd.Connection = gbl_SQLConnection
        mSQLcmd.CommandType = CommandType.Text
        mSQLcmd.CommandText = Build_Update_Trans_SQL_Record_Statement(
            CInt(cmb_URCSYear.Text),
            CInt(mRailroad),
            mSch,
            mLine,
            mC1,
            mC2).ToString
        mSQLcmd.ExecuteNonQuery()

        'Clean up
        mDataTable = Nothing
        xclTable = Nothing

        mWorkbook.Close()

        'Advise the user
        Me.txt_StatusBox.Text = "Done!"
        Me.Refresh()

EndIt:

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


End Class