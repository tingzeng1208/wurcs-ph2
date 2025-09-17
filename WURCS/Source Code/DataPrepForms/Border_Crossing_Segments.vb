'''////////////////////////////////////////////////////////////////////////////////////////////////////
''' <summary>   A border crossing segments. </summary>
'''
''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
'''////////////////////////////////////////////////////////////////////////////////////////////////////

Public Class Border_Crossing_Segments

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close this Form
        Me.Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Border crossing segments load. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub Border_Crossing_Segments_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        mDataTable = Get_URCS_Years_Table()

        For mLooper = 1 To mDataTable.Rows.Count
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
        '***************************************************************************
        ' This program reads the WByear_Segments table as selected on the form and then
        ' builds the SQL statement to create the corresponding Border_Crossings_Segments table.
        '***************************************************************************

        Dim cnn As New ADODB.Connection
        Dim rst As New ADODB.Recordset
        Dim mStrSQL As String, mStrOutSQL As String
        Dim mTableName As String
        Dim mBolWrite As Boolean
        Dim mSglMaxRecs As Single, mThisRec As Single
        Dim mShipTypes(8) As String
        Dim mPrevious_Segment As Class_Segments, mCurrent_Segment As Class_Segments
        Dim mSglOutRecs As Single, mSglBorderCnt As Single

        'Load the table information
        mTableName = "WB" & CStr(Me.cmb_URCS_Year.Text) & "_Segments"

        ' Load the Year combobox from the SQL database
        ' Open the SQL connection using the global variable holding the connection string
        cnn.ConnectionString = Global_Variables.gbl_SQLConnString.ToString
        cnn.Open()

        'Initialize the return recordset and run the query to load it
        rst = New ADODB.Recordset
        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Me.txt_StatusBox.Text = "Searching for records..."
        Me.Refresh()

        mStrSQL = "SELECT COUNT(*) as MyCount FROM " & mTableName

        rst.Open(mStrSQL, cnn)
        System.Windows.Forms.Cursor.Current = Cursors.Default
        Me.txt_StatusBox.Text = ""
        Me.Refresh()

        mSglMaxRecs = rst.Fields("").Value

        mBolWrite = MsgBox("The number of records returned is " & rst.Fields("MyCount").Value &
                ". Are you sure you want to write the data?", vbYesNo)

        If mBolWrite = True Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.txt_StatusBox.Text = "Wiping the output tables"
            Me.Refresh()

            ' Set the table name for the border crossings segments table
            mTableName = "WB" & Me.cmb_URCS_Year.Text & "_Border_Crossings_Segments"

            ' We will delete/drop the WBxxxx_Border_Crossings_Segments table here

            ' Check to see if the table exists before trying the DROP command
            If TableExist("Waybills", mTableName) Then
                mStrOutSQL = "DROP TABLE " & mTableName

                cnn.Execute(mStrOutSQL)
            End If

            ' We need to create it again (this drop/create routine is much faster
            ' than deleting the records from the existing table and reusing it.)

            mStrOutSQL = "CREATE TABLE " & mTableName &
                "( " &
                "Serial_No Int NOT NULL, " &
                "Seg_no tinyint NOT NULL, " &
                "Total_Segs tinyint NULL, " &
                "RR_Num smallint NULL, " &
                "RR_Alpha nvarchar(5) NULL, " &
                "RR_Dist int NULL, " &
                "RR_Cntry nvarchar(2) NULL, " &
                "RR_Rev decimal(18, 0) NULL, " &
                "RR_VC int NULL, " &
                "Seg_Type nvarchar(2) NULL, " &
                "From_Node int NULL, " &
                "To_Node int NULL, " &
                "From_Loc nvarchar(9) NULL, " &
                "From_St nvarchar(5) NULL, " &
                "To_Loc nvarchar(9) NULL, " &
                "To_St nvarchar(5) NULL) on [PRIMARY]"

            cnn.Execute(mStrOutSQL)

            ' Now we need to create the index for the table we just created.

            mStrOutSQL = "ALTER TABLE " & mTableName &
                " ADD CONSTRAINT pk_" & mTableName &
                " PRIMARY KEY CLUSTERED " &
                "(Serial_No, Seg_No) WITH (STATISTICS_NORECOMPUTE= OFF, " &
                "IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, " &
                "ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"

            cnn.Execute(mStrOutSQL)

            Me.txt_StatusBox.Text = "Getting Segments Data From Server"
            Me.Refresh()

            rst.Close()

            mStrSQL = "SELECT * From WB" & Me.cmb_URCS_Year.Text &
                "_Segments"

            rst.Open(mStrSQL, cnn)

            rst.MoveFirst()
            mThisRec = 1
            mSglMaxRecs = rst.RecordCount
            mSglOutRecs = 0
            mSglBorderCnt = 0

            ' Read the first record and set it to be the previous segment
            mPrevious_Segment = New Class_Segments

            With mPrevious_Segment
                .Serial_No = rst.Fields("serial_no").Value
                .Segment_No = rst.Fields("seg_no").Value
                .Total_Segments = rst.Fields("total_segs").Value
                .RR_Num = rst.Fields("rr_num").Value
                .RR_Alpha = rst.Fields("rr_alpha").Value
                .RR_Dist = rst.Fields("rr_dist").Value
                .RR_Cntry = rst.Fields("rr_cntry").Value
                .RR_Rev = rst.Fields("rr_rev").Value
                .RR_VC = rst.Fields("rr_vc").Value
                .Seg_Type = rst.Fields("seg_type").Value
                .From_Node = rst.Fields("from_node").Value
                .To_Node = rst.Fields("to_node").Value
                .From_Loc = rst.Fields("from_loc").Value
                .From_St = rst.Fields("from_st").Value
                .To_Loc = rst.Fields("to_loc").Value
                .To_St = rst.Fields("to_st").Value
            End With

            rst.MoveNext()
            mThisRec = mThisRec + 1

            Do While Not rst.EOF
                If mThisRec Mod 100 = 0 Then
                    Me.txt_StatusBox.Text = "Processing record " & CStr(mThisRec) & " of " & CStr(mSglMaxRecs) & " - " &
                        "Border Crossings: " & CStr(mSglBorderCnt)
                    Me.Refresh()
                    Application.DoEvents()
                End If

                mCurrent_Segment = New Class_Segments

                With mCurrent_Segment
                    .Serial_No = rst.Fields("serial_no").Value
                    .Segment_No = rst.Fields("seg_no").Value
                    .Total_Segments = rst.Fields("total_segs").Value
                    .RR_Num = rst.Fields("rr_num").Value
                    .RR_Alpha = rst.Fields("rr_alpha").Value
                    .RR_Dist = rst.Fields("rr_dist").Value
                    .RR_Cntry = rst.Fields("rr_cntry").Value
                    .RR_Rev = rst.Fields("rr_rev").Value
                    .RR_VC = rst.Fields("rr_vc").Value
                    .Seg_Type = rst.Fields("seg_type").Value
                    .From_Node = rst.Fields("from_node").Value
                    .To_Node = rst.Fields("to_node").Value
                    .From_Loc = rst.Fields("from_loc").Value
                    .From_St = rst.Fields("from_st").Value
                    .To_Loc = rst.Fields("to_loc").Value
                    .To_St = rst.Fields("to_st").Value
                End With

                If (mCurrent_Segment.Serial_No = mPrevious_Segment.Serial_No) _
                    And (mCurrent_Segment.RR_Num = mPrevious_Segment.RR_Num) Then

                    mPrevious_Segment.Seg_Type = Strings.Left(mPrevious_Segment.Seg_Type, 1) & "X"
                    mCurrent_Segment.Seg_Type = "X" & Strings.Right(mCurrent_Segment.Seg_Type, 1)
                    mSglBorderCnt = mSglBorderCnt + 1
                End If

                ' We need to write the "previous" data out to the Border Crossing table here
                mStrSQL = Build_Segment_SQL(mTableName,
                    mPrevious_Segment.Serial_No,
                    mPrevious_Segment.Segment_No,
                    mPrevious_Segment.Total_Segments,
                    mPrevious_Segment.RR_Num,
                    mPrevious_Segment.RR_Alpha,
                    mPrevious_Segment.RR_Dist,
                    mPrevious_Segment.RR_Cntry,
                    mPrevious_Segment.RR_Rev,
                    mPrevious_Segment.RR_VC,
                    mPrevious_Segment.Seg_Type,
                    mPrevious_Segment.From_Node,
                    mPrevious_Segment.To_Node,
                    mPrevious_Segment.From_Loc,
                    mPrevious_Segment.From_St,
                    mPrevious_Segment.To_Loc,
                    mPrevious_Segment.To_St).ToString

                cnn.Execute(mStrSQL)

                mSglOutRecs = mSglOutRecs + 1

                ' Now move the "current" values to "previous" segments class

                mPrevious_Segment.Serial_No = mCurrent_Segment.Serial_No
                mPrevious_Segment.Segment_No = mCurrent_Segment.Segment_No
                mPrevious_Segment.Total_Segments = mCurrent_Segment.Total_Segments
                mPrevious_Segment.RR_Num = mCurrent_Segment.RR_Num
                mPrevious_Segment.RR_Alpha = mCurrent_Segment.RR_Alpha
                mPrevious_Segment.RR_Dist = mCurrent_Segment.RR_Dist
                mPrevious_Segment.RR_Cntry = mCurrent_Segment.RR_Cntry
                mPrevious_Segment.RR_Rev = mCurrent_Segment.RR_Rev
                mPrevious_Segment.RR_VC = mCurrent_Segment.RR_VC
                mPrevious_Segment.Seg_Type = mCurrent_Segment.Seg_Type
                mPrevious_Segment.From_Node = mCurrent_Segment.From_Node
                mPrevious_Segment.To_Node = mCurrent_Segment.To_Node
                mPrevious_Segment.From_Loc = mCurrent_Segment.From_Loc
                mPrevious_Segment.From_St = mCurrent_Segment.From_St
                mPrevious_Segment.To_Loc = mCurrent_Segment.To_Loc
                mPrevious_Segment.To_St = mCurrent_Segment.To_St

                rst.MoveNext()
                mThisRec = mThisRec + 1
            Loop

            ' Lastly, write the "previous" segment out to the Border Croissings table

            mStrSQL = Build_Segment_SQL(mTableName,
                mPrevious_Segment.Serial_No,
                mPrevious_Segment.Segment_No,
                mPrevious_Segment.Total_Segments,
                mPrevious_Segment.RR_Num,
                mPrevious_Segment.RR_Alpha,
                mPrevious_Segment.RR_Dist,
                mPrevious_Segment.RR_Cntry,
                mPrevious_Segment.RR_Rev,
                mPrevious_Segment.RR_VC,
                mPrevious_Segment.Seg_Type,
                mPrevious_Segment.From_Node,
                mPrevious_Segment.To_Node,
                mPrevious_Segment.From_Loc,
                mPrevious_Segment.From_St,
                mPrevious_Segment.To_Loc,
                mPrevious_Segment.To_St).ToString

            cnn.Execute(mStrSQL)

        End If

        rst.Close()
        rst = Nothing

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Me.txt_StatusBox.Text = "Done!"
        Me.Refresh()

    End Sub
End Class