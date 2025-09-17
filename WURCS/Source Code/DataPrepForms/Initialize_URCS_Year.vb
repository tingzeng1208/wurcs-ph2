Imports System.Data.SqlClient

Public Class Initialize_URCS_Year

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Form initialize urcs year load. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub frm_Initialize_URCS_Year_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Set the form so it centers on the user's screen
        Me.CenterToScreen()
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
        Dim mDataTable As New DataTable
        Dim mCommand As New SqlCommand

        Dim mSQLCmd As String
        Dim bolWrite As Integer, mThisLine As Integer, mLooper As Integer, mLineCnt As Integer

        ' Check to make sure the user entered a year
        If IsNothing(Me.txt_URCS_Year.Text) Then
            MsgBox("You must enter a year value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        'Check to make sure that the user wants to actually do this
        bolWrite = MsgBox("Are you sure you want to initialize the data for " & txt_URCS_Year.Text & "?", MsgBoxStyle.YesNo, "Are you Sure?")

        If bolWrite = vbYes Then
            '**************************************************************
            ' Creat Database on SQL Server
            '**************************************************************

            txt_StatusBox.Text = "Creating Database..."
            Refresh()

            ' Make sure we haven't left a connection open previously
            If gbl_SQLConnection Is Nothing Then
                ' we're good
            Else
                If gbl_SQLConnection.State = ConnectionState.Open Then
                    gbl_SQLConnection.Close()
                End If
            End If

            gbl_Database_Name = "Master"
            gbl_SQLConnection = New SqlConnection
            mCommand.Connection = gbl_SQLConnection
            mCommand.CommandType = CommandType.Text

            gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(gbl_Database_Name)
            gbl_SQLConnection.Open()

            mSQLCmd = "Select name from Sys.Databases WHERE name = 'URCS" & txt_URCS_Year.Text & "'"

            Using daAdapter As New SqlDataAdapter(mSQLCmd, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count = 0 Then
                ' It doesn't exist, so we can go ahead and create it
                gbl_SQLConnection.Close()
                gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(gbl_Database_Name)
                gbl_SQLConnection.Open()
                mCommand.CommandText = Create_Database_Command(txt_URCS_Year.Text)
                mCommand.ExecuteNonQuery()
            Else
                ' It does exist, so can we go ahead and drop/recreate it
                bolWrite = MsgBox("Database exists!  Overwrite?", vbYesNo, "CAUTION!")
                If bolWrite = vbNo Then
                    txt_StatusBox.Text = "Aborted."
                    GoTo EndIt
                Else
                    mCommand.CommandText = "DROP DATABASE " & "URCS" & txt_URCS_Year.Text
                    mCommand.ExecuteNonQuery()
                    mCommand.CommandText = Create_Database_Command(txt_URCS_Year.Text)
                    mCommand.ExecuteNonQuery()
                End If
            End If

            ' Change the connection to the database
            gbl_SQLConnection.Close()
            gbl_Database_Name = "URCS" & txt_URCS_Year.Text
            gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(gbl_Database_Name)
            gbl_SQLConnection.Open()

            mCommand.Connection = gbl_SQLConnection

            txt_StatusBox.Text = "Creating ActivityAuditLog Table..."
            Refresh()
            Create_AuditLog_Table(gbl_Database_Name, "ActivityAuditLog")

            Insert_AuditTrail_Record(gbl_Database_Name, "Started Initialization for URCS " & txt_URCS_Year.Text & ".")

            txt_StatusBox.Text = "Creating U_AVALUES Table..."
            Refresh()
            Create_AValues_Table(gbl_Database_Name, "U_AVALUES")

            txt_StatusBox.Text = "Creating U_CRPRES Table..."
            Refresh()
            Create_CRPRES_Tables(gbl_Database_Name, "U_CRPRES")

            txt_StatusBox.Text = "Creating U_CRPRES_Legacy Table..."
            Refresh()
            Create_CRPRES_Tables(gbl_Database_Name, "U_CRPRES_Legacy")

            txt_StatusBox.Text = "Creating U_CRPSEG Table..."
            Refresh()
            Create_CRPSEG_Tables(gbl_Database_Name, "U_CRPSEG")

            txt_StatusBox.Text = "Creating U_CRPSEG_1 Table..."
            Refresh()
            Create_CRPSEG_Tables(gbl_Database_Name, "U_CRPSEG_1")

            txt_StatusBox.Text = "Creating U_CRPSEG_2 Table..."
            Refresh()
            Create_CRPRES_Tables(gbl_Database_Name, "U_CRPSEG_2")

            txt_StatusBox.Text = "Creating U_CRPSEG_Legacy Table..."
            Refresh()
            Create_CRPSEG_Tables(gbl_Database_Name, "U_CRPSEG_Legacy")

            txt_StatusBox.Text = "Creating U_CRPSEG_Legacy_1 Table..."
            Refresh()
            Create_CRPSEG_Tables(gbl_Database_Name, "U_CRPSEG_Legacy_1")

            txt_StatusBox.Text = "Creating U_CRPSEG_Legacy_2 Table..."
            Refresh()
            Create_CRPSEG_Tables(gbl_Database_Name, "U_CRPSEG_Legacy_2")

            txt_StatusBox.Text = "Creating ERRORS Table..."
            Refresh()
            Create_Errors_Table(gbl_Database_Name, "U_ERRORS")

            txt_StatusBox.Text = "Creating U_EVALUES Table..."
            Refresh()
            Create_EValues_Table(gbl_Database_Name, "U_EVALUES")

            txt_StatusBox.Text = "Creating U_SUBSTITUTIONS Table..."
            Refresh()
            Create_Substitutions_Table(gbl_Database_Name, "U_SUBSTITUTIONS")

            txt_StatusBox.Text = "Creating UT_MAKEWHOLE_FACTORS Table..."
            Refresh()
            Create_MakeWhole_Factors_Table(gbl_Database_Name, "UT_MAKEWHOLE_FACTORS")

            Create_ufn_EValues_Function(txt_URCS_Year.Text)

            txt_StatusBox.Text = "Creating Functions and Procedures..."
            Refresh()

            CreateProcedure(txt_URCS_Year.Text, "usp_GenerateEValuesXML")
            CreateProcedure(txt_URCS_Year.Text, "usp_RunSubstitutions")
            CreateProcedure(txt_URCS_Year.Text, "usp_WriteCRPSEG_Legacy")

            ' Close the connection and open the connection to the Waybills database
            gbl_SQLConnection.Close()
            gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(My.Settings.Waybills_DB)
            gbl_SQLConnection.Open()

            mCommand.Connection = gbl_SQLConnection

            txt_StatusBox.Text = "Creating Waybill Tables..."
            Refresh()

            'Check to see if Masked table has already been created
            If TableExist("Waybills", "WB" & txt_URCS_Year.Text & "_Masked") = False Then
                ' Need to create the Masked Table here
                Create_Masked_913_Table(My.Settings.Waybills_DB, "WB" & txt_URCS_Year.Text & "_Masked")
                'Otherwise don't create it
            End If

            'Check to see if UnMasked_Rev table has already been created
            If TableExist("Waybills", "WB" & txt_URCS_Year.Text & "_UnMasked_Rev") = False Then
                ' Need to create the Unmasked_Rev Table here
                Create_Unmasked_Rev_Table(My.Settings.Waybills_DB, "WB" & txt_URCS_Year.Text & "_Unmasked_Rev")
            Else
                'Otherwise don't create it
            End If

            'Check to see if Segments table has already been created
            If TableExist("Waybills", "WB" & txt_URCS_Year.Text & "_Segments") = False Then
                ' Need to create the Segments Table here
                Create_Segments_Table(My.Settings.Waybills_DB, "WB" & txt_URCS_Year.Text & "_Segments")
            Else
                'Otherwise don't create it
            End If

            'Check to see if the UnMasked_Segments table has already been created
            If TableExist("Waybills", "WB" & txt_URCS_Year.Text & "_UnMasked_Segments") = False Then
                ' Need to create the Unmasked_Segments Table here
                Create_Unmasked_Segments_Table(My.Settings.Waybills_DB, "WB" & txt_URCS_Year.Text & "_Unmasked_Segments")
            Else
                'Otherwise don't create it
            End If

            'Check to see if PUWS table has already been created
            If TableExist("Waybills", "PUWS" & txt_URCS_Year.Text & "_Masked") = False Then
                ' Create the PUWS_Masked
                Create_PUWS_Masked_Table(txt_URCS_Year.Text)
            Else
                'Otherwise don't create it
            End If

            If TableExist("Waybills", "PUWS" & txt_URCS_Year.Text & "_Masking_Factors") = False Then
                ' Create the PUWS_Masking_Factors
                Create_PUWS_Masking_Factors_Table_Command(txt_URCS_Year.Text)
            Else
                'Otherwise don't create it
            End If

            ' open the connection to the URCS_Controls instance
            gbl_SQLConnection.Close()
            gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(My.Settings.Controls_DB)
            gbl_SQLConnection.Open()

            mCommand.Connection = gbl_SQLConnection

            ' Now we need to add the table locator entries for the tables

            txt_StatusBox.Text = "updating Locator Table..."
            Refresh()

            Create_Locator_Entry(txt_URCS_Year.Text, "AValues")
            Create_Locator_Entry(txt_URCS_Year.Text, "CRPRES")
            Create_Locator_Entry(txt_URCS_Year.Text, "CRPRES_Legacy")
            Create_Locator_Entry(txt_URCS_Year.Text, "CRPSEG")
            Create_Locator_Entry(txt_URCS_Year.Text, "CRPSEG_1")
            Create_Locator_Entry(txt_URCS_Year.Text, "CRPSEG_2")
            Create_Locator_Entry(txt_URCS_Year.Text, "CRPSEG_Legacy")
            Create_Locator_Entry(txt_URCS_Year.Text, "CRPSEG_Legacy_1")
            Create_Locator_Entry(txt_URCS_Year.Text, "CRPSEG_Legacy_2")
            Create_Locator_Entry(txt_URCS_Year.Text, "Errors")
            Create_Locator_Entry(txt_URCS_Year.Text, "EValues")
            Create_Locator_Entry(txt_URCS_Year.Text, "Substitutions")
            Create_Locator_Entry(txt_URCS_Year.Text, "Makewhole_Factors")
            Create_Locator_Entry(txt_URCS_Year.Text, "Masked")
            Create_Locator_Entry(txt_URCS_Year.Text, "Unmasked_Rev")
            Create_Locator_Entry(txt_URCS_Year.Text, "Segments")
            Create_Locator_Entry(txt_URCS_Year.Text, "Unmasked_Segments")
            Create_Locator_Entry(txt_URCS_Year.Text, "PUWS_Masked")
            Create_Locator_Entry(txt_URCS_Year.Text, "PUWS_Masking_Factors")
            Create_Locator_Entry(txt_URCS_Year.Text, "ActivityAuditLog")

            ' Add the year value to the URCS Year table (used for lookups in lots of the forms)
            Create_URCS_Year_Entry(txt_URCS_Year.Text)

            ' Add the year value to the Waybill Year table (used for lookups in lots of the forms)
            Create_Waybill_Year_Entry(txt_URCS_Year.Text)

            txt_StatusBox.Text = "Initializing Trans Data..."
            Refresh()

            Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "Trans")

            mCommand.CommandText = "DELETE FROM " & Trim(Gbl_Trans_TableName) & " WHERE year = " & txt_URCS_Year.Text
            mCommand.ExecuteNonQuery()

            Insert_AuditTrail_Record("URCS" & txt_URCS_Year.Text, "Deleted URCS_Controls->Trans Table records (if present) for " & txt_URCS_Year.Text & ".")

            ' Get the default values from the URCS_Defaults table

            Gbl_URCS_Defaults_TableName = Get_Table_Name_From_SQL("1", "URCS_DEFAULTS")

                mDataTable = New DataTable

                mSQLCmd = "SELECT * FROM " & Gbl_URCS_Defaults_TableName

                Using daAdapter As New SqlDataAdapter(mSQLCmd, gbl_SQLConnection)
                    daAdapter.Fill(mDataTable)
                End Using

                mThisLine = 0
                mLineCnt = 0

                mCommand.Connection = gbl_SQLConnection
                mCommand.CommandType = CommandType.Text

                If mDataTable.Rows.Count = 0 Then
                    bolWrite = MsgBox("FATAL ERROR - URCS Defaults table not found.", vbOK, "FATAL ERROR")
                    txt_StatusBox.Text = "Fatal Error - URCS Defaults table not found."
                    GoTo EndIt
                End If

                For mLooper = 0 To mDataTable.Rows.Count - 1
                    If mThisLine <> mDataTable.Rows(mLooper)("line") Then
                        mThisLine = mDataTable.Rows(mLooper)("line")
                        mLineCnt = mLineCnt + 1
                    End If

                    txt_StatusBox.Text = "Processing Trans Line " & mLineCnt.ToString & " of 213..."
                Refresh()

                ' create the record with values from defaults table
                mCommand.CommandText = Build_Count_Trans_SQL_Statement(
                    txt_URCS_Year.Text,
                    mDataTable.Rows(mLooper)("rricc").ToString,
                    mDataTable.Rows(mLooper)("sch").ToString,
                    mDataTable.Rows(mLooper)("line").ToString)

                    If mCommand.ExecuteScalar = 0 Then
                        mCommand.CommandText = Build_Insert_Trans_SQL_Field_Statement(
                    txt_URCS_Year.Text,
                    mDataTable.Rows(mLooper)("rricc").ToString,
                    mDataTable.Rows(mLooper)("sch").ToString,
                    mDataTable.Rows(mLooper)("line").ToString,
                    mDataTable.Rows(mLooper)("col").ToString,
                    mDataTable.Rows(mLooper)("urcs_val").ToString)
                        mCommand.ExecuteNonQuery()
                    Else
                        mCommand.CommandText = Build_Update_Trans_SQL_Field_Statement(
                        txt_URCS_Year.Text,
                    mDataTable.Rows(mLooper)("rricc").ToString,
                    mDataTable.Rows(mLooper)("sch").ToString,
                    mDataTable.Rows(mLooper)("line").ToString,
                    mDataTable.Rows(mLooper)("col").ToString,
                    mDataTable.Rows(mLooper)("urcs_val").ToString)
                        mCommand.ExecuteNonQuery()
                    End If
                Next

                Insert_AuditTrail_Record("URCS" & txt_URCS_Year.Text, "Inserted Default National Records into URCS_Controls->Trans Table.")

                mDataTable = Nothing
                gbl_SQLConnection.Close()

                txt_StatusBox.Text = "Done!"

            Else

                txt_StatusBox.Text = "Aborted."

        End If

EndIt:

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
        ' Close the Tare Weight Loader Form
        Me.Close()
    End Sub
End Class