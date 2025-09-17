Imports System.Data.SqlClient
Public Class CSM_Data_Loader

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button return to main menu click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Return_To_MainMenu_Click(sender As Object, e As EventArgs) Handles btn_Return_To_MainMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close the Tare Weight Loader Form
        Me.Close()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Csm data loader load. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub CSM_Data_Loader_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Center the form on the user's screen
        CenterToScreen()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button input file entry click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Input_File_Entry_Click(sender As Object, e As EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Text Files|*.txt|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_FilePath.Text = fd.FileName
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button execute click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Execute_Click(sender As Object, e As EventArgs) Handles btn_Execute.Click

        Dim mSr As StreamReader
        Dim mSQLCommand As New SqlCommand

        Dim bolWrite As Integer
        Dim mInString As String
        Dim mThisRec As Integer, mMaxRecs As Integer
        Dim mThisRoadMark As String, mThisMaintStamp As String, mThisFSAC As String
        Dim mSQLStr As String

        If txt_Input_FilePath.TextLength = 0 Then
            MsgBox("You must select an Input File value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        ' ask the user if he wants to load x records for this RRID
        bolWrite = MsgBox("Are you sure you want to Write this data?", vbYesNo)

        If bolWrite = vbYes Then

            'Gbl_Controls_Database_Name = My.Settings.Controls_DB
            Gbl_Controls_Database_Name = Get_Database_Name_From_SQL("1", "CSM")
            Gbl_CSM_TableName = Get_Table_Name_From_SQL("1", "CSM")

            'If the table does not exist, create it
            If VerifyTableExist(Gbl_Controls_Database_Name, Gbl_CSM_TableName) = False Then
                Create_CSM_Table()
            End If

            mThisRec = 0
            mMaxRecs = 0

            mSr = New StreamReader(txt_Input_FilePath.Text)

            txt_StatusBox.Text = "Examining input file..."
            Refresh()

            ' Count how many lines we have in the file
            Do While Not mSr.EndOfStream
                mSr.ReadLine()
                mMaxRecs = mMaxRecs + 1
            Loop

            ' Move the pointer to the top of the file
            mSr.Close()

            mSr = New StreamReader(txt_Input_FilePath.Text)

            txt_StatusBox.Text = "Working..."
            Refresh()

            Do While Not mSr.EndOfStream

                mInString = mSr.ReadLine()

                mThisRec = mThisRec + 1
                If mThisRec Mod 100 = 0 Then
                    txt_StatusBox.Text = "Working - Loaded " & mThisRec.ToString & " of " & mMaxRecs.ToString & " records..."
                    Refresh()
                    System.Windows.Forms.Application.DoEvents()
                End If

                'Get the index key values for this table
                mThisRoadMark = Trim(mInString.Substring(0, 4))
                mThisFSAC = Trim(mInString.Substring(4, 5))
                mThisMaintStamp = Trim(mInString.Substring(754, 26))

                'Check to see if this record exists
                If Count_CSM_Records(mThisRoadMark, mThisFSAC, mThisMaintStamp) = 0 Then
                    ' This record doesn't exist. We need to insert it
                    mSQLStr = Build_Insert_CSM_Record(mInString, Gbl_CSM_TableName)
                Else
                    ' A record already exists.  We need to update it
                    mSQLStr = Build_Update_CSM_Record(mInString, Gbl_CSM_TableName)
                End If

                ' Execute the command
                mSQLCommand.Connection = gbl_SQLConnection
                mSQLCommand.CommandText = mSQLStr
                mSQLCommand.ExecuteNonQuery()

            Loop

            txt_StatusBox.Text = "Done!"
            Refresh()

            'Clean up
            mSr.Close()
            mSr = Nothing

        Else
            txt_StatusBox.Text = "Operation Aborted"
        End If


EndIt:
    End Sub
End Class