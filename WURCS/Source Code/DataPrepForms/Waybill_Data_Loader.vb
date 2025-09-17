Imports System.Data.SqlClient
Public Class Waybill_Data_Loader

    Private Const ForReading = 1
    Private Const ForAppending = 8

    Private Sub Waybill_Data_Loader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        CenterToScreen()

        ' Load the Year combobox from the SQL database
        mDataTable = Get_Waybill_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            cmb_URCSYear.Items.Add(mDataTable.Rows(mLooper)("wb_year").ToString)
        Next

        mDataTable = Nothing

    End Sub

    Private Sub btn_Input_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Txt Files|*.txt|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_FilePath.Text = fd.FileName
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

        Dim mDataTable As New DataTable
        Dim mSQLCommand As New SqlCommand

        Dim strSQL As String, strInline As String
        Dim Rec As Long, mMaxRecs As Long
        Dim bolOverwrite As Object
        Dim fs, infile

        ' Perform error checks on the form controls
        If cmb_URCSYear.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo EndIt
        End If

        If txt_Input_FilePath.Text = "" Then
            MsgBox("You must select a filename to read.", vbOKOnly)
            GoTo EndIt
        End If

        ' Determine if the tables exist for this year.  Create them if they don't.

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        'Set up the environment
        LoadArrayData()

        bolOverwrite = False

        ' open the file that the user wants to load
        fs = CreateObject("Scripting.FileSystemObject")
        infile = fs.OpenTextFile(txt_Input_FilePath.Text, ForReading)
        strInline = infile.readline

        Cursor.Current = Cursors.WaitCursor

        txt_StatusBox.Text = "Examining Data Environment - " & cmb_URCSYear.Text & " Text File...  Please wait"
        Refresh()

        ' find out how many records are currently in the table
        mMaxRecs = File.ReadAllLines(txt_Input_FilePath.Text).Length

        Cursor.Current = Cursors.Default

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        Select Case Len(strInline)
            Case 445
                Gbl_Masked_TableName = "WB" & cmb_URCSYear.Text & "_Annual_Interim"
            Case 900, 913
                Gbl_Masked_TableName = "WB" & cmb_URCSYear.Text & "_Masked"
        End Select

        Cursor.Current = Cursors.WaitCursor

        If TableExist(Gbl_Waybill_Database_Name, Gbl_Masked_TableName) = True Then

            ' delete the existing table
            mSQLCommand.CommandType = CommandType.Text
            mSQLCommand.Connection = gbl_SQLConnection
            mSQLCommand.CommandText = "DROP TABLE " & Gbl_Masked_TableName
            mSQLCommand.ExecuteNonQuery()

        End If

        txt_StatusBox.Text = "Creating/Overwriting tables..."

        Create_Masked_913_Table(Gbl_Waybill_Database_Name, Gbl_Masked_TableName)

        'close and then reopen the txt file to move the pointer back to BOF
        infile.Close()
        infile = fs.OpenTextFile(txt_Input_FilePath.Text, ForReading)

        'Insert_AuditTrail_Record("URCS_Controls", "Inserted Waybill Data to " & Gbl_Masked_TableName)

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        ' cycle through the records one by one
        Rec = 1
        Do While infile.atendofstream <> True
            If Rec Mod 1000 = 0 Then
                txt_StatusBox.Text = "Loading/examining record " & CStr(Rec) & " of " & CStr(mMaxRecs) & "..."
                Text = "Progress - " & CStr(Math.Round((Rec / mMaxRecs) * 100)) & "%"
                Refresh()
                Application.DoEvents()
            End If
            strInline = infile.readline

            Select Case Len(strInline)

                Case 445
                    strSQL = Build445SQL(
                           cmb_URCSYear.Text,
                           Gbl_Masked_TableName,
                           strInline)

                Case 913
                    'Updated to address 913 waybill record.  11/16/2020
                    strSQL = Build913SQL(
                            cmb_URCSYear.Text,
                            Gbl_Masked_TableName,
                            strInline)

                Case Else
                    MsgBox("Input file format is unrecognized. Load aborted.", vbOKOnly)
                    infile.Close()
                    infile = Nothing
                    GoTo EndIt

            End Select

            mSQLCommand = New SqlCommand
            mSQLCommand.CommandType = CommandType.Text
            mSQLCommand.Connection = gbl_SQLConnection
            mSQLCommand.CommandText = strSQL
            mSQLCommand.ExecuteNonQuery()
            Rec = Rec + 1

        Loop

        txt_StatusBox.Text = "Done!"
        Refresh()

        infile.Close()
        infile = Nothing

EndIt:

    End Sub

End Class