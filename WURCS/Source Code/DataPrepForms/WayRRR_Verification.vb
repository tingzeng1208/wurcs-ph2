Imports System.Data.SqlClient
Public Class WayRRR_Verification

    Dim mErrorAlphas() As String
    Dim mErrorRRIDs() As Integer
    Dim mRailroadsAlphas() As String
    Dim mRailroadsRRIDs() As Integer
    Dim mErrors As Integer

    Private Sub WayRRR_Verification_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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

    Private Sub btn_Output_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New FolderBrowserDialog

        fd.Description = "Select the location in which you want the output report placed."

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Output_FilePath.Text = fd.SelectedPath.ToString & "\WB" & cmb_URCS_Year.Text & "_WayRRR_Exceptions.txt"
        End If
    End Sub

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close this Form
        Close()
    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        Dim rst As New ADODB.Recordset
        Dim mDataTable As DataTable
        Dim mStrSQL As String
        Dim mArrayPos As Integer
        Dim mMaxRecs As Single, mLooper As Single, mJCTLooper As Integer
        Dim fs, outfile

        If cmb_URCS_Year.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo EndIt
        End If

        If txt_Output_FilePath.Text = "" Then
            MsgBox("You must select a filename to use or create.", vbOKOnly)
            GoTo EndIt
        End If

        ' Ensure the error arrays are set/cleared.
        ReDim mErrorAlphas(0)
        ReDim mErrorRRIDs(0)

        ' Open the SQL connection using the global variable holding the connection string
        OpenSQLConnection(Get_Database_Name_From_SQL("1", "WAYRRR"))

        ' Get the WAYRRR data from the table
        mDataTable = New DataTable
        mStrSQL = "SELECT * From " & Get_Table_Name_From_SQL("1", "WAYRRR")

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        ' resize the arrays to fit
        ReDim mRailroadsAlphas(mDataTable.Rows.Count)
        ReDim mRailroadsRRIDs(mDataTable.Rows.Count)

        ' And load it to the arrays
        For mLooper = 0 To mDataTable.Rows.Count - 1
            mRailroadsAlphas(mLooper + 1) = mDataTable.Rows(mLooper)("rr_name")
            mRailroadsRRIDs(mLooper + 1) = mDataTable.Rows(mLooper)("aarid")
        Next

        ' Find the location and name of the Waybill table and the database in which it resides
        gbl_Database_Name = Get_Database_Name_From_SQL(cmb_URCS_Year.Text, "MASKED")
        gbl_Table_Name = Get_Table_Name_From_SQL(cmb_URCS_Year.Text, "MASKED")
        OpenSQLConnection(gbl_Database_Name)

        'Initialize the recordset and run the query to load it.

        txt_StatusBox.Text = "Processing Waybills for " & cmb_URCS_Year.Text &
            " - Please wait...  Fetching Data"
        Refresh()
        Cursor.Current = Cursors.WaitCursor

        ' Then get the waybill fields we need.  No real need to load the entire waybill.
        mDataTable = New DataTable
        mStrSQL = "SELECT serial_no, orr_alpha, jrr1_alpha, jrr2_alpha, jrr3_alpha, " &
            "jrr4_alpha, jrr5_alpha, jrr6_alpha, trr_alpha, jf, orr, jrr1, jrr2, jrr3, " &
            "jrr4, jrr5, jrr6, trr FROM " & gbl_Table_Name

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        If mDataTable.Rows.Count = 0 Then
            MsgBox("No records found for " & cmb_URCS_Year.Text, vbOKOnly, "Error!")
            GoTo EndIt
        End If

        mMaxRecs = mDataTable.Rows.Count

        If MsgBox("The number of records found is " & mMaxRecs & ".  Are you sure you want to process the data?", vbYesNo) = vbNo Then
            ' Terminate the operation.
            txt_StatusBox.Text = "Aborted."
            GoTo EndIt
        End If

        For mLooper = 0 To mDataTable.Rows.Count - 1

            If mLooper Mod 1000 = 0 Then
                txt_StatusBox.Text = "Loading/examining record " & CStr(mLooper) & " of " & CStr(mMaxRecs) & "..."
                Text = "Progress - " & CStr(Math.Round((mLooper / mMaxRecs) * 100)) & "%"
                Refresh()
                Application.DoEvents()
            End If

            ' Always check the ORR
            mArrayPos = Array1DFindFirst(mRailroadsRRIDs, CInt(mDataTable.Rows(mLooper)("orr")))
            If mArrayPos = 0 Then
                ' We've got a road that isn't known, so store/check it into the
                ' error arrays
                AddError(mDataTable.Rows(mLooper)("orr"), Trim(mDataTable.Rows(mLooper)("orr_alpha")))
            End If

            ' Always check the TRR
            mArrayPos = Array1DFindFirst(mRailroadsRRIDs, CInt(mDataTable.Rows(mLooper)("trr")))
            If mArrayPos = 0 Then
                ' We've got a road that isn't known, so store/check it into the
                ' error arrays
                AddError(mDataTable.Rows(mLooper)("trr"), Trim(mDataTable.Rows(mLooper)("trr_alpha")))
            End If

            ' Now we check the JRR's as per the JF value, starting with the last
            ' and working backwards

            For mJCTLooper = 1 To 6

                If mDataTable.Rows(mLooper)("jf") > mJCTLooper Then
                    mArrayPos = Array1DFindFirst(mRailroadsRRIDs, mDataTable.Rows(mLooper)("jrr" & mJCTLooper.ToString))
                    If mArrayPos = 0 Then
                        ' We've got a road that isn't known, so store/check it into the
                        ' error arrays
                        AddError(mDataTable.Rows(mLooper)("jrr" & mJCTLooper.ToString), mDataTable.Rows(mLooper)("jrr" & mJCTLooper.ToString & "_alpha"))
                    End If
                End If

            Next mJCTLooper


        Next mLooper

        mDataTable = Nothing

        ' Create the text file contents from the resulting error arrays

        ' Open the output file
        fs = CreateObject("Scripting.FileSystemObject")
        outfile = fs.createtextfile(txt_Output_FilePath.Text, True)

        outfile.writeline("Railroads missing from WayRRR for WB" & cmb_URCS_Year.Text &
            "_Masked")
        outfile.writeline("")

        For mLooper = 1 To mErrors
            outfile.writeline(mErrorAlphas(mLooper) & ", " & CStr(mErrorRRIDs(mLooper)))
        Next mLooper

        outfile.Close()
        outfile = Nothing
        fs = Nothing

        txt_StatusBox.Text = "Done!"
        Text = "Done!"
        Refresh()

        Cursor.Current = Cursors.Default

EndIt:

    End Sub

    Private Sub AddError(ByVal mErrorRRID As Integer, ByVal mErrorAlpha As String)

        Dim mArrayPos As Integer

        mArrayPos = Array1DFindFirst(mErrorRRIDs, mErrorRRID)

        ' Exclude CN/CNUS and CP/CPRS as errors
        If (mErrorRRID = 103 Or mErrorRRID = 105) And
            (mErrorAlpha = "CNUS" Or mErrorAlpha = "CPRS") Then
            ' Ignore it
        Else
            If mArrayPos > 0 Then
                ' This railroad RRID has already been added to the error array
            Else
                If mErrors = 0 Then
                    ' We need to redim the array to 1
                    mErrors = 1
                    ReDim mErrorAlphas(mErrors)
                    ReDim mErrorRRIDs(mErrors)
                Else
                    ' we need to redim the array(x) + 1
                    mErrors = mErrors + 1
                    ReDim Preserve mErrorAlphas(mErrors)
                    ReDim Preserve mErrorRRIDs(mErrors)
                End If
                ' Now we can store the values
                mErrorAlphas(mErrors) = mErrorAlpha
                mErrorRRIDs(mErrors) = mErrorRRID
            End If
        End If

    End Sub
End Class