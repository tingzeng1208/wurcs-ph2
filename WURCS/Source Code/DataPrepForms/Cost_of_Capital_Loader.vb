Imports SpreadsheetGear
Imports System.Data.SqlClient

Public Class Cost_of_Capital_Loader

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
        ' Close the Tare Weight Loader Form
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Cost of capital loader load. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub Cost_of_Capital_Loader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        CenterToScreen()

        ' Load the Year combobox from the SQL database

        mDataTable = Get_URCS_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            cmb_URCSYear.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
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

        ' Variables for SpreadsheetGear
        Dim mWorkbook As IWorkbook
        Dim mExcelSheet As IWorksheet

        Dim mDataTable As DataTable
        Dim mSQLCommand As SqlCommand
        Dim mStrSQL As String
        Dim RRICCArray(10) As Decimal
        Dim idx As Integer, mWorkval As Decimal

        ' get these from SQL
        RRICCArray(1) = 900004
        RRICCArray(2) = 900007
        RRICCArray(3) = 900099
        RRICCArray(4) = 114900
        RRICCArray(5) = 117000
        RRICCArray(6) = 125600
        RRICCArray(7) = 130500
        RRICCArray(8) = 134500
        RRICCArray(9) = 137700
        RRICCArray(10) = 139300

        'Ensure we have the Trans table location
        Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "TRANS")

        ' Perform Error checking for form controls
        If cmb_URCSYear.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo EndIt
        ElseIf IsNumeric(cmb_URCSYear.Text) = False Then
            MsgBox("Invalid Year value entered.", vbOKOnly)
            GoTo EndIt
        End If

        ' Perform Error checking for form controls
        If txt_CostOfCapital.Text = "" Then
            MsgBox("You must enter a value for the Cost of Capital.", vbOKOnly)
            GoTo EndIt
        ElseIf IsNumeric(txt_CostOfCapital.Text) = False Then
            MsgBox("Invalid Cost of Capital value entered.", vbOKOnly)
            GoTo EndIt
        End If

        If txt_Report_FilePath.TextLength = 0 Then
            MsgBox("You must select an Report File value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        ' Load the Year combobox from the SQL database
        ' Open the SQL connection using the global variable holding the connection string
        OpenSQLConnection(Gbl_Controls_Database_Name)

        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection

        For idx = 1 To 10
            'Check to see if the record exists
            mDataTable = New DataTable
            mStrSQL = Build_Count_Trans_SQL_Statement(cmb_URCSYear.Text,
                RRICCArray(idx),
                45,
                205)

            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            ' Remove decimal point
            mWorkval = Val(Replace(txt_CostOfCapital.Text, ".", ""))

            If mDataTable.Rows(0)(0) > 0 Then
                'there is already a record there and we need
                'to update it
                mStrSQL = Build_Update_Trans_SQL_Field_Statement(cmb_URCSYear.Text,
                    RRICCArray(idx),
                    45,
                    205,
                    1,
                    Val(mWorkval)).ToString

                mSQLCommand.CommandText = mStrSQL.ToString
            Else
                'No record exists - we need to insert it
                mStrSQL = Build_Insert_Trans_SQL_Field_Statement(cmb_URCSYear.Text,
                    RRICCArray(idx),
                    45,
                    205,
                    1,
                    Val(mWorkval)).ToString
                mSQLCommand.CommandText = mStrSQL.ToString
            End If
            mSQLCommand.ExecuteNonQuery()
        Next idx

        txt_StatusBox.Text = "Creating & Loading Excel File..."
        Refresh()

        ' Open the Workbook
        mWorkbook = Factory.GetWorkbook(Globalization.CultureInfo.CurrentCulture)

        'Generate the report
        mWorkbook.ActiveSheet.Name = "Summary"

        mExcelSheet = mWorkbook.ActiveSheet
        mExcelSheet.Cells().Font.Size = 12 ' Sets default font size
        mExcelSheet.Cells("A1").Value = "URCS Cost of Capital Load Process Report - " & cmb_URCSYear.Text.ToString
        mExcelSheet.Cells("A1").Font.Bold = True
        mExcelSheet.Cells("A1").Font.Size = 14
        mExcelSheet.Cells("A2").Value = "Date/Time: " & CStr(Now).ToString

        mExcelSheet.Cells("A4").Value = "RR/Region"
        mExcelSheet.Cells("B4").Value = "Sch"
        mExcelSheet.Cells("C4").Value = "Line"
        mExcelSheet.Cells("D4").Value = "Column"
        mExcelSheet.Cells("E4").Value = "Value"

        mExcelSheet.Cells("A4:E4").HorizontalAlignment = SpreadsheetGear.HAlign.Right
        mExcelSheet.Cells("A4:E4").Font.Underline = True

        For idx = 1 To 10

            mExcelSheet.Cells("A" & (idx + 5).ToString).Value = Get_Short_Name_By_RRICC(RRICCArray(idx)).ToString
            mExcelSheet.Cells("B" & (idx + 5).ToString).Value = 45
            mExcelSheet.Cells("C" & (idx + 5).ToString).Value = 205
            mExcelSheet.Cells("D" & (idx + 5).ToString).Value = 1

            mDataTable = New DataTable

            Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "TRANS")

            mStrSQL = "SELECT C1 FROM " & Gbl_Trans_TableName & " WHERE Year = " & cmb_URCSYear.Text &
                                " AND RRICC = " & RRICCArray(idx).ToString &
                                " AND SCH = 45" &
                                " AND LINE = 205"

            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If RRICCArray(idx) = 900099 Then
                mExcelSheet.Cells("A" & (idx + 5).ToString).Value = "NATIONAL"
            End If

            mExcelSheet.Cells("E" & (idx + 5).ToString).Value = mDataTable.Rows(0)("C1")

            mDataTable = Nothing

        Next

        mWorkbook.SaveAs(txt_Report_FilePath.Text, FileFormat.OpenXMLWorkbook)

        txt_StatusBox.Text = "Done!"
        Refresh()

EndIt:

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Cmb urcs year selected index changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub cmb_URCSYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_URCSYear.SelectedIndexChanged

        Dim mStrSQL As String
        Dim mDataTable As DataTable

        ' Load the Year combobox from the SQL database
        ' Open the SQL connection using the global variable holding the connection string
        OpenSQLConnection(Gbl_Controls_Database_Name)

        'Build the SQL statement
        mStrSQL = "SELECT c1 FROM " & Get_Table_Name_From_SQL("1", "TRANS").ToString & " " &
            "WHERE year = " & cmb_URCSYear.Text & " AND " &
            "sch = 45 AND line = 205"

        mDataTable = New DataTable

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        'get the first record - they're all the same anyways
        If mDataTable.Rows.Count = 0 Then
            txt_CostOfCapital.Text = "0.0"
        Else
            txt_CostOfCapital.Text = mDataTable.Rows(0)(0)
        End If

        mDataTable = Nothing
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button report file entry click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Report_File_Entry_Click(sender As System.Object, e As System.EventArgs) Handles btn_Report_File_Entry.Click
        Dim fd As New FolderBrowserDialog

        If cmb_URCSYear.Text = "" Then
            MsgBox("You must select a year first.", vbOKOnly, "Error!")
        Else
            fd.Description = "Select the location in which you want the output report placed."

            If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txt_Report_FilePath.Text = fd.SelectedPath.ToString & "\WB" & cmb_URCSYear.Text & " Cost of Capital Load Report.xlsx"
            End If
        End If
    End Sub
End Class