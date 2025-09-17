Imports System.Data.SqlClient
Imports System.Text
Imports System.IO

' Added Time Hacks for Log file                                 MRS 02/08/2022
' Changed references for segment_no from table to memvar        MRS 02/25/2022
' Changed log file option to append rather than overwrite       MRS 12/27/2022

Public Class Interim_Data_Loader
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Interim data loader load. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/9/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub Interim_Data_Loader_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()

        gbx_Monthly.Visible = False
        gbx_Quarter.Visible = False

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button return to data prep menu click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/9/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close this Form
        Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Radio quarterly checked changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/9/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub rdo_Quarterly_CheckedChanged(sender As Object, e As EventArgs) Handles rdo_Quarterly.CheckedChanged

        gbx_Monthly.Visible = False
        rdo_1st_Quarter.Visible = True
        rdo_2nd_Quarter.Visible = True
        rdo_3rd_Quarter.Visible = True
        rdo_4th_Quarter.Visible = True
        gbx_Quarter.Visible = True

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Radio monthly checked changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/9/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub rdo_Monthly_CheckedChanged(sender As Object, e As EventArgs) Handles rdo_Monthly.CheckedChanged

        gbx_Quarter.Visible = False
        lbl_Month_Combobox.Visible = True
        cmb_Month.Visible = True
        gbx_Monthly.Visible = True

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button input file entry click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/9/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Input_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Txt Files|*.txt|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_FilePath.Text = fd.FileName
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button execute click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/9/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Execute_Click(sender As Object, e As EventArgs) Handles btn_Execute.Click
        Dim mmsgtext As String = ""
        Dim mWorkStr As String

        If (txt_Year.Text = "") Or (Val(txt_Year.Text) < 2019) Then
            MsgBox("You must enter a valid year value.", vbOKOnly)
            GoTo EndIt
        End If

        If chk_Skip_Masked_Data_Load.Checked = False And txt_Input_FilePath.Text = "" Then
            MsgBox("You must enter an input filename.", vbOKOnly)
            GoTo EndIt
        End If

        If rdo_Monthly.Checked = True Then
            mmsgtext = "This will load/create a table of WB"
            mmsgtext = mmsgtext & txt_Year.Text & "M" & cmb_Month.Text & "_Masked. Is this correct?"
            If MsgBox(mmsgtext, vbYesNo) = vbNo Then
                txt_StatusBox.Text = "Aborted"
                GoTo EndIt
            Else
                mWorkStr = "WB" & txt_Year.Text & "M" & cmb_Month.Text
                Gbl_Interim_Masked = mWorkStr & "_Masked"
                Gbl_Unmasked_Rev_TableName = mWorkStr & "_Unmasked_Rev"
                Gbl_Segments_TableName = mWorkStr & "_Segments"
                Gbl_Unmasked_Segments_TableName = mWorkStr & "_Unmasked_Segments"
            End If
        End If

        If rdo_Quarterly.Checked = True Then
            mmsgtext = "This will load/create a table of WB"
            mmsgtext = mmsgtext & txt_Year.Text & "Q"
            mWorkStr = "WB" & txt_Year.Text & "Q"
            If rdo_1st_Quarter.Checked = True Then
                mmsgtext = mmsgtext & "1"
                gbl_Table_Name = mWorkStr & "1"
            ElseIf rdo_2nd_Quarter.Checked = True Then
                mmsgtext = mmsgtext & "2"
                gbl_Table_Name = mWorkStr & "2"
            ElseIf rdo_3rd_Quarter.Checked = True Then
                mmsgtext = mmsgtext & "3"
                gbl_Table_Name = mWorkStr & "3"
            ElseIf rdo_4th_Quarter.Checked = True Then
                mmsgtext = mmsgtext & "4"
                gbl_Table_Name = mWorkStr & "4"
            End If
            mmsgtext = mmsgtext & "_Masked. Is this correct?"
            If MsgBox(mmsgtext, vbYesNo) = vbNo Then
                txt_StatusBox.Text = "Aborted"
                GoTo EndIt
            End If
            Gbl_Interim_Masked = gbl_Table_Name & "_Masked"
            Gbl_Unmasked_Rev_TableName = gbl_Table_Name & "_Unmasked_Rev"
            Gbl_Segments_TableName = gbl_Table_Name & "_Segments"
            Gbl_Unmasked_Segments_TableName = gbl_Table_Name & "_Unmasked_Segments"
        End If

        mmsgtext = "This will load/create a table of WB"
        If rdo_Annual.Checked = True Then
            mmsgtext = mmsgtext & txt_Year.Text & "Y_Masked. Is this correct?"
            If MsgBox(mmsgtext, vbYesNo) = vbNo Then
                txt_StatusBox.Text = "Aborted"
                GoTo EndIt
            Else
                Gbl_Interim_Masked = "WB" & txt_Year.Text & "Y_Masked"
                Gbl_Unmasked_Rev_TableName = "WB" & txt_Year.Text & "Y_Unmasked_Rev"
                Gbl_Segments_TableName = "WB" & txt_Year.Text & "Y_Segments"
                Gbl_Unmasked_Segments_TableName = "WB" & txt_Year.Text & "Y_Unmasked_Segments"
                Gbl_Annual_Interim_TableName = "WB" & txt_Year.Text & "Y_Interim"
            End If
        End If

        ' One of the conditions above have met the criteria/checks
        ' We have enough to begin processing

        Load_Data_to_SQL()

EndIt:

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text year leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/9/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub txt_Year_Leave(sender As Object, e As EventArgs) Handles txt_Year.Leave

        If (txt_Year.Text = "") Or (Val(txt_Year.Text) < 2019) Then
            MsgBox("You must enter a valid year value", vbOKOnly)
        End If

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Loads data to SQL. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/9/2020. </remarks>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub Load_Data_to_SQL()

        Dim msqlCommand As New SqlCommand
        Dim mSQLCommandString As StringBuilder
        Dim mMaxRecs As BigInteger
        Dim mThisRec As BigInteger
        Dim mBPInLine As StringBuilder
        Dim sSr As StreamReader
        Dim strInline As String
        Dim mProcess As Process

        Dim sbLogFile As StreamWriter
        Dim OutString = New StringBuilder

        Dim mRun_Start As Date
        Dim mSQL_Init_Starts, mSQL_Init_Ends As Date
        Dim mLoad_Starts, mLoad_Ends As Date
        Dim mSegments_Update_Starts, mSegments_Update_Ends As Date
        Dim mBatchPro_Starts, mBatchPro_Ends As Date
        Dim mSTCC_and_BEA_Updates_Start, mSTCC_and_BEA_Updates_Ends As Date
        Dim mPUWS_Update_Starts, mPUWS_Update_Ends As Date
        Dim mUnmasked_Update_Starts, mUnmasked_Update_Ends As Date
        Dim mRun_Stops As Date

        Dim mSTCC_W49 As String, mCar_Init As String, mTState As String
        Dim mAcctMonth As String, mAcctYear As String, mU_Rev_Unmasked As Long
        Dim mWBDate As Date, mYear As Integer, mExp_Factor_Th As Integer, mStratum As Integer
        Dim mSerial_No As String = "", mU_TC_Init As String, mTotal_Rev As Long
        Dim mWbNum As Long, mTableType As String = "", mU_TC_Units As Integer
        Dim mUCarNum As Long, mTCNum As Long, mUCars As Long
        Dim mRate_Flg As Integer, mRptRR As Integer, mU_Cars As Integer, mu_TC_Num As Integer
        Dim mORR As Integer, mRebill As Integer, mTonsPerCar As Integer
        Dim mJunctions(9) As String
        Dim mRailRoads(10) As Integer, mStratumErrors As Integer = 0
        Dim mJCT_ST(9) As String
        Dim mSeg_Type As String, mFrom_Loc As Object, mTo_Loc As Object
        Dim mLooper As Integer, mLooper2 As Integer, mRR_ST As String
        Dim mTRR As Integer, mDistance As Integer, mTOFC_Serv_Code As String
        Dim mJF As Integer, mNumberOfRRs As Integer
        Dim mU_Rev As Integer, mO_FSAC As Long, mO_FSAC_ST As String, mT_FSAC As Long, mT_FSAC_ST As String
        Dim mTotal_Segs As Integer, mSeg_No As Integer, mRR_Num As Integer, mRR_Alpha As String, mJCT As String
        Dim Work_Directory As String, mRoute_Formula As Integer, mCar_Typ As Integer
        Dim mSTCC2 As Integer, mSTCC3 As Integer, mSTCC4 As Integer, mSTCC5 As Integer, mTons As Long
        Dim mServiceType As Integer, mInt_Eq_Flg As Integer, mBill_Wght_Tons As Long, mBill_Wght As Long
        Dim mFrom_ST As String, mTo_ST As String, mRR_Cntry As String, mWorkDate As Date
        Dim mList_of_FSACs_TableName As String, mList_of_JCTs_TableName As String, mList_of_RRs_TableName As String
        Dim mFrom_Lat As String, mFrom_Long As String, mTo_Lat As String, mTo_Long As String
        Dim mORR_Dist As Integer, mJRR1_Dist As Integer, mJRR2_Dist As Integer, mJRR3_Dist As Integer
        Dim mJRR4_Dist As Integer, mJRR5_Dist As Integer, mJRR6_Dist As Integer, mTRR_Dist As Integer
        Dim mTotal_Dist As Integer, mTotal_Unmasked_Rev As Integer
        Dim mOriginType As String, mUCarInit As String, mExpanded_Unmasked_Revenue As Double, mPUWS_Masking_Factor As Double
        Dim mDestinationType As String
        Dim mProcessParameters As StringBuilder
        Dim mWorkStr As String
        Dim mFrom_GeoCodeType As Integer
        Dim mFrom_GeoCodeValue As String
        Dim mTo_GeoCodeType As Integer
        Dim mTo_GeoCodeValue As String
        Dim mRR_Dist As Integer, mRR_Miles As Decimal
        Dim mErrorMsg As String
        Dim mBadFlag As Boolean
        Dim mFieldName As String
        Dim mService_Type As Integer
        Dim mArrayLoc As Integer = 0
        Dim mSPLC4toBEA_Table_Name As String = ""
        Dim mSPLC6toBEA_Table_Name As String = ""

        Dim mSegmentsTable As DataTable
        Dim mAllDateMiledTable As DataTable
        Dim mFSACTable As DataTable
        Dim mJCTTable As DataTable
        Dim mRRTable As DataTable
        Dim mMaskedTable As DataTable
        Dim mSTCCTable As DataTable
        Dim mInterim_Raw_Table As DataTable
        Dim mWorkTable As DataTable
        Dim mDatarow() As DataRow
        Dim SPLC6() As String
        Dim SPLC6_BEA() As String
        Dim SPLC4() As String
        Dim SPLC4_BEA() As String

        Dim mInterim_FileName_Root As String
        Dim mBPAllInFileName As String
        Dim mBPAllOutFilename As String
        Dim mBPGM_FileNames(4) As String
        Dim mBPIM_Filenames(4) As String
        Dim mBPCB_Filenames(4) As String
        Dim mBPAR_Filenames(4) As String

        Dim mBPGM_FileTypes(4) As String
        Dim mBPIM_FileTypes(4) As String
        Dim mBPCB_FileTypes(4) As String
        Dim mBPAR_FileTypes(4) As String

        Dim mBatchPro_Runs(4) As String

        Dim mInFileNameSuffix As String = ".IN"
        Dim mCfgFileNameSuffix As String = ".CFG"
        Dim mOutFileNameSuffix As String = ".OUT"
        Dim mRPTFileNameSuffix As String = ".RPT"
        Dim mAllInFileName As String = "_BatchPro_ALL.IN"
        Dim mAllOutFileName As String = "_BatchPro_ALL.OUT"
        Dim mRailIncErrorLogFileName As String = "_RailIncErrors.ERR"
        Dim mPeriodLength As String = ""
        Dim mAnnual_Type As String = ""

        ' Declare the streamwriters for the above files.
        Dim sBPAllInFilename As StreamWriter
        Dim sBPAllOutFile As StreamWriter
        Dim sBPGM_Streams(4) As StreamWriter
        Dim sBPIM_Streams(4) As StreamWriter
        Dim sBPCB_Streams(4) As StreamWriter
        Dim sBPAR_Streams(4) As StreamWriter
        Dim sRailIncErrorStream As StreamWriter

RunStart:

        Work_Directory = txt_Work_Directory.Text & "\"
        mRun_Start = Now()

        If chk_Skip_Masked_Data_Load.Checked = False Then
            mWorkStr = Path.GetFileName(txt_Input_FilePath.Text)
            mWorkStr = Replace(mWorkStr, ".txt", ".LOG")
            mWorkStr = Work_Directory & mWorkStr

            sbLogFile = New StreamWriter(mWorkStr, True) 'Appends to existing log file if it exists
            sbLogFile.WriteLine("Interim Processing for " & txt_Input_FilePath.Text & " started at " & mRun_Start.ToString("G") & " ")
            sbLogFile.Flush()
        Else
            mWorkStr = Work_Directory & txt_Year.Text
            If rdo_Quarterly.Checked Then
                If rdo_1st_Quarter.Checked Then
                    mWorkStr = mWorkStr & "Q1 Processing.LOG"
                ElseIf rdo_2nd_Quarter.Checked Then
                    mWorkStr = mWorkStr & "Q2 Processing.LOG"
                ElseIf rdo_3rd_Quarter.Checked Then
                    mWorkStr = mWorkStr & "Q3 Processing.LOG"
                ElseIf rdo_4th_Quarter.Checked Then
                    mWorkStr = mWorkStr & "Q4 Processing.LOG"
                End If
            End If
            sbLogFile = New StreamWriter(mWorkStr, False)

            If rdo_Quarterly.Checked Then
                If rdo_1st_Quarter.Checked Then
                    mWorkStr = txt_Year.Text & "Q1"
                ElseIf rdo_2nd_Quarter.Checked Then
                    mWorkStr = txt_Year.Text & "Q2"
                ElseIf rdo_3rd_Quarter.Checked Then
                    mWorkStr = txt_Year.Text & "Q3"
                ElseIf rdo_4th_Quarter.Checked Then
                    mWorkStr = txt_Year.Text & "Q4"
                End If
            ElseIf rdo_Monthly.Checked Then
                mWorkStr = txt_Year.Text & "M"
                mWorkStr = mWorkStr & cmb_Month.SelectedItem.ToString
            ElseIf rdo_Annual.Checked Then
                mWorkStr = txt_Year.Text & "Y"
            End If
            sbLogFile.WriteLine("Interim Processing for " & mWorkStr & " started at " & mRun_Start.ToString("G") & " ")
            sbLogFile.Flush()

        End If

        mSQL_Init_Starts = Now()

        LoadArrayData()

        ' The filetypes will always be:
        '   1 is the FSAC2FSAC file
        mBatchPro_Runs(1) = "FSAC2FSAC"
        '   2 is the FSAC2JCT file
        mBatchPro_Runs(2) = "FSAC2JCT"
        '   3 is the JCT2FSAC file
        mBatchPro_Runs(3) = "JCT2FSAC"
        '   4 is the JCT2JCT file
        mBatchPro_Runs(4) = "JCT2JCT"

        mBPGM_FileTypes(1) = "BatchPro_GM_FSAC2FSAC"
        mBPGM_FileTypes(2) = "BatchPro_GM_FSAC2JCT"
        mBPGM_FileTypes(3) = "BatchPro_GM_JCT2FSAC"
        mBPGM_FileTypes(4) = "BatchPro_GM_JCT2JCT"
        mBPIM_FileTypes(1) = "BatchPro_IM_FSAC2FSAC"
        mBPIM_FileTypes(2) = "BatchPro_IM_FSAC2JCT"
        mBPIM_FileTypes(3) = "BatchPro_IM_JCT2FSAC"
        mBPIM_FileTypes(4) = "BatchPro_IM_JCT2JCT"
        mBPCB_FileTypes(1) = "BatchPro_CB_FSAC2FSAC"
        mBPCB_FileTypes(2) = "BatchPro_CB_FSAC2JCT"
        mBPCB_FileTypes(3) = "BatchPro_CB_JCT2FSAC"
        mBPCB_FileTypes(4) = "BatchPro_CB_JCT2JCT"
        mBPAR_FileTypes(1) = "BatchPro_AR_FSAC2FSAC"
        mBPAR_FileTypes(2) = "BatchPro_AR_FSAC2JCT"
        mBPAR_FileTypes(3) = "BatchPro_AR_JCT2FSAC"
        mBPAR_FileTypes(4) = "BatchPro_AR_JCT2JCT"

        txt_StatusBox.Text = "Setting up Work Area..."
        Refresh()

        mInterim_FileName_Root = "WB" & txt_Year.Text

        If rdo_Annual.Checked = True Then
            If rdo_Annual.Checked = True Then
                mPeriodLength = "Y"
            End If
        ElseIf rdo_Monthly.Checked = True Then
            If cmb_Month.Text = "" Then
                MsgBox("You must select a month value.", vbOKOnly, "Error!")
                GoTo EndIt
            End If
            mPeriodLength = "M" & cmb_Month.Text
        Else
            If rdo_1st_Quarter.Checked Then
                mPeriodLength = "Q1"
            ElseIf rdo_2nd_Quarter.Checked Then
                mPeriodLength = "Q2"
            ElseIf rdo_3rd_Quarter.Checked Then
                mPeriodLength = "Q3"
            ElseIf rdo_4th_Quarter.Checked Then
                mPeriodLength = "Q4"
            End If
        End If

        'Set the SQL table names
        Gbl_Interim_Raw = mInterim_FileName_Root & mPeriodLength & "_Interim"
        Gbl_Interim_Masked = mInterim_FileName_Root & mPeriodLength & "_Masked"
        Gbl_Interim_Unmasked_Rev = mInterim_FileName_Root & mPeriodLength & "_Unmasked_Rev"
        Gbl_Interim_Segments = mInterim_FileName_Root & mPeriodLength & "_Segments"
        Gbl_Interim_Unmasked_Segments = mInterim_FileName_Root & mPeriodLength & "_Unmasked_Segments"
        Gbl_Interim_BatchPro_All_Miled = mInterim_FileName_Root & mPeriodLength & "_BatchPro_All_Miled"
        Gbl_interim_PUWS = mInterim_FileName_Root & mPeriodLength & "_PUWS_Rev"
        Gbl_Interim_PUWS_Masked_Rev_TableName = mInterim_FileName_Root & mPeriodLength & “_PUWS_Masked_Rev"

        'Set the ALL filenames
        mBPAllInFileName = mInterim_FileName_Root & mPeriodLength & mAllInFileName
        mBPAllOutFilename = mInterim_FileName_Root & mPeriodLength & mAllOutFileName

        If Directory.Exists(Work_Directory) = False Then
            Directory.CreateDirectory(Work_Directory)
        End If

        My.Computer.FileSystem.CopyDirectory(
            "\\amznfsxwpjnid69.stb.gov\oeeaa\URCS Work Areas\Software\PCMiler\CFG_Library ", Work_Directory, True)

        'Set the filenames
        For mLooper = 1 To 4
            mBPGM_FileNames(mLooper) = mInterim_FileName_Root & mPeriodLength & "_" & mBPGM_FileTypes(mLooper)
            mBPIM_Filenames(mLooper) = mInterim_FileName_Root & mPeriodLength & "_" & mBPIM_FileTypes(mLooper)
            mBPCB_Filenames(mLooper) = mInterim_FileName_Root & mPeriodLength & "_" & mBPCB_FileTypes(mLooper)
            mBPAR_Filenames(mLooper) = mInterim_FileName_Root & mPeriodLength & "_" & mBPAR_FileTypes(mLooper)
        Next

        ' Set up and open the connection to the database.
        Gbl_Interim_Waybills_Database_Name = My.Settings.InterimWaybills

        OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)

        'Load the arrays for unmasking functions
        LoadArrayData()

        'Load the following tables to datatables in memory for lookup values
        mList_of_FSACs_TableName = Get_Table_Name_From_SQL("1", "Active_FSACS")
        mFSACTable = New DataTable
        Load_Datatable_From_SQL(Gbl_Waybill_Database_Name, mList_of_FSACs_TableName, mFSACTable)

        mList_of_JCTs_TableName = Get_Table_Name_From_SQL("1", "Active_JCTs")
        mJCTTable = New DataTable
        Load_Datatable_From_SQL(Gbl_Waybill_Database_Name, mList_of_JCTs_TableName, mJCTTable)

        mList_of_RRs_TableName = Get_Table_Name_From_SQL("1", "Active_RRS")
        mRRTable = New DataTable
        Load_Datatable_From_SQL(Gbl_Waybill_Database_Name, mList_of_RRs_TableName, mRRTable) 'Load the arrays for unmasking functions
        LoadArrayData()

        'Load the following tables to datatables in memory for lookup values
        mList_of_FSACs_TableName = Get_Table_Name_From_SQL("1", "Active_FSACS")
        mFSACTable = New DataTable
        Load_Datatable_From_SQL(Gbl_Waybill_Database_Name, mList_of_FSACs_TableName, mFSACTable)

        mList_of_JCTs_TableName = Get_Table_Name_From_SQL("1", "Active_JCTs")
        mJCTTable = New DataTable
        Load_Datatable_From_SQL(Gbl_Waybill_Database_Name, mList_of_JCTs_TableName, mJCTTable)

        mList_of_RRs_TableName = Get_Table_Name_From_SQL("1", "Active_RRS")
        mRRTable = New DataTable
        Load_Datatable_From_SQL(Gbl_Waybill_Database_Name, mList_of_RRs_TableName, mRRTable)

        If chk_Skip_Segments_Data_Load.Checked = True Then
            sbLogFile.WriteLine("Segments Mileage Data Load from OUT files skipped.")
            sbLogFile.Flush()
            GoTo Skip_Segments_Data_Load
        End If

        If chk_Skip_BatchPro_Processing.Checked = True Then
            sbLogFile.WriteLine("BatchPro Processing skipped. Using All Miled data.")
            sbLogFile.Flush()
            GoTo Skip_BatchPro_Processing
        End If

        If chk_Skip_Masked_Data_Load.Checked = True Then
            sbLogFile.WriteLine("Masked Data Load skipped - Using data already in Masked table.")
            sbLogFile.Flush()
            GoTo Skip_Masked_Data_Load
        End If

Initialize_SQL:

        mSQL_Init_Starts = Now()

        Gbl_Interim_Waybills_Database_Name = My.Settings.InterimWaybills

        'drop the Interim table from SQL                    
        If VerifyTableExist(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Raw) = True Then
            OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)
            msqlCommand = New SqlCommand
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = "Drop table [dbo].[" & Gbl_Interim_Raw & "]"
            msqlCommand.ExecuteNonQuery()
        End If

        'drop the Masked table from SQL                    
        If VerifyTableExist(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked) = True Then
            OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)
            msqlCommand = New SqlCommand
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = "Drop table [dbo].[" & Gbl_Interim_Masked & "]"
            msqlCommand.ExecuteNonQuery()
        End If

        'drop the unmasked rev table
        If VerifyTableExist(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Unmasked_Rev) = True Then
            OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)
            msqlCommand = New SqlCommand
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = "Drop table [dbo]." & Gbl_Interim_Unmasked_Rev
            msqlCommand.ExecuteNonQuery()
        End If

        'drop the segments table
        If VerifyTableExist(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Segments) = True Then
            OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)
            msqlCommand = New SqlCommand
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = "Drop table [dbo]." & Gbl_Interim_Segments
            msqlCommand.ExecuteNonQuery()
        End If

        'drop the unmasked segments table
        If VerifyTableExist(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Unmasked_Segments) = True Then
            OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)
            msqlCommand = New SqlCommand
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = "Drop table [dbo]." & Gbl_Interim_Unmasked_Segments
            msqlCommand.ExecuteNonQuery()
        End If

        'drop the All Miled table
        If VerifyTableExist(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_BatchPro_All_Miled) = True Then
            OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)
            msqlCommand = New SqlCommand
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = "Drop table [dbo]." & Gbl_Interim_BatchPro_All_Miled
            msqlCommand.ExecuteNonQuery()
        End If

        'drop the Interim PUWS table
        If VerifyTableExist(Gbl_Interim_Waybills_Database_Name, Gbl_interim_PUWS) = True Then
            OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)
            msqlCommand = New SqlCommand
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = "Drop table [dbo]." & Gbl_interim_PUWS
            msqlCommand.ExecuteNonQuery()
        End If

        'Create the tables
        txt_StatusBox.Text = "Preparing SQL tables..."
        Refresh()

        If VerifyTableExist(Gbl_Interim_Waybills_Database_Name, "ActivityAuditLog") = False Then
            Create_AuditLog_Table(Gbl_Interim_Waybills_Database_Name, "ActivityAuditLog")
        End If

        Create_Masked_913_Table(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked)
        Create_Masked_445_Table(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Raw)
        Create_Unmasked_Rev_Table(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Unmasked_Rev)
        Create_Segments_Table(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Segments)
        Create_Unmasked_Segments_Table(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Unmasked_Segments)
        Create_BatchPro_ALL_Miled_Table(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_BatchPro_All_Miled)
        Create_Interim_PUWS_Table(Gbl_Interim_Waybills_Database_Name, Gbl_interim_PUWS)

        If Directory.Exists(Work_Directory) = False Then
            Directory.CreateDirectory(Work_Directory)
        Else
            Directory.SetCurrentDirectory(Work_Directory)
        End If

        ' Find out how many records we're loading. 
        txt_StatusBox.Text = "Examining Input Data..."
        Refresh()

        ' Delete the OUT file if it exists
        mWorkStr = Work_Directory & "WB" &
            txt_Year.Text & mPeriodLength & mRailIncErrorLogFileName
        If File.Exists(mWorkStr) Then
            File.Delete(mWorkStr)
        End If

        ' Delete the RailInc Error file if it exists
        mWorkStr = Work_Directory & "WB" &
            txt_Year.Text & mPeriodLength & mRailIncErrorLogFileName
        If File.Exists(mWorkStr) Then
            File.Delete(mWorkStr)
        End If

        mSQL_Init_Ends = Now()

        OutString = New StringBuilder
        sbLogFile.WriteLine()
        OutString.Append("SQL Initialization complete at " & mSQL_Init_Ends.ToString("G") & " ")
        sbLogFile.WriteLine(OutString)
        OutString = New StringBuilder
        OutString.Append(Return_Elapsed_Time(mSQL_Init_Starts, mSQL_Init_Ends))
        sbLogFile.WriteLine(OutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

LoadMasked:

        'Start the timer for the Load process
        mLoad_Starts = Now()

        'Open the BatchPro input file streamwriters
        sBPAllInFilename = New StreamWriter(Work_Directory & mBPAllInFileName, False)

        For mLooper = 1 To 4
            sBPGM_Streams(mLooper) = New StreamWriter(Work_Directory & mInterim_FileName_Root & mPeriodLength & "_" & mBPGM_FileTypes(mLooper) & mInFileNameSuffix)
            sBPIM_Streams(mLooper) = New StreamWriter(Work_Directory & mInterim_FileName_Root & mPeriodLength & "_" & mBPIM_FileTypes(mLooper) & mInFileNameSuffix)
            sBPCB_Streams(mLooper) = New StreamWriter(Work_Directory & mInterim_FileName_Root & mPeriodLength & "_" & mBPCB_FileTypes(mLooper) & mInFileNameSuffix)
            sBPAR_Streams(mLooper) = New StreamWriter(Work_Directory & mInterim_FileName_Root & mPeriodLength & "_" & mBPAR_FileTypes(mLooper) & mInFileNameSuffix)
        Next

        sRailIncErrorStream = New StreamWriter(Work_Directory & mInterim_FileName_Root & mPeriodLength & mRailIncErrorLogFileName)

        mMaxRecs = File.ReadAllLines(txt_Input_FilePath.Text).Length
        mThisRec = 0

        ' Open the text file
        sSr = New StreamReader(txt_Input_FilePath.Text)

        txt_StatusBox.Text = "Writing Data to SQL..."
        Refresh()

        ' Begin looping thru the input file
        While (sSr.Peek <> -1)
            strInline = sSr.ReadLine
            mThisRec = mThisRec + 1

            ' Let the user know what is going on
            If mThisRec Mod 100 = 0 Then
                txt_StatusBox.Text = "Working - Loaded " & mThisRec.ToString & " of " & mMaxRecs.ToString & " records. "
                Refresh()
                Application.DoEvents()
            End If

            ' Set the environment to default values
            mSerial_No = vbNull
            mWbNum = 0
            mRptRR = 0
            mDistance = 0
            mRate_Flg = 0
            mSTCC_W49 = 0
            mJF = 0
            mCar_Init = " "
            mWBDate = New Date
            mWorkDate = New Date
            mYear = 0
            mAcctMonth = ""
            mAcctYear = ""
            mUCarNum = 0
            mTCNum = 0
            mTState = 0
            mUCars = 0
            mU_Rev = 0
            mRebill = 0
            mSeg_Type = ""
            mFrom_Loc = ""
            mFrom_ST = ""
            mTo_Loc = ""
            mTo_ST = ""
            mNumberOfRRs = 0
            mO_FSAC = 0
            mT_FSAC = 0
            mORR = 0
            mTRR = 0
            mSeg_No = 0
            mTotal_Segs = 0
            mO_FSAC_ST = ""
            mT_FSAC_ST = ""
            mCar_Typ = 0
            mTOFC_Serv_Code = ""
            mSTCC2 = 0
            mSTCC3 = 0
            mSTCC4 = 0
            mSTCC5 = 0
            mServiceType = 0
            mInt_Eq_Flg = 0
            mBill_Wght = 0
            mBill_Wght_Tons = 0
            mTonsPerCar = 0
            mU_Cars = 0
            mU_TC_Init = ""
            mu_TC_Num = 0
            mFrom_Lat = ""
            mFrom_Long = ""
            mTo_Lat = ""
            mTo_Long = ""
            mRR_Alpha = ""
            mRR_ST = ""
            mJCT = ""
            mExp_Factor_Th = 0
            mStratum = 0
            mTotal_Rev = 0
            mU_TC_Units = 0
            mU_Rev_Unmasked = 0
            mOriginType = ""
            mDestinationType = ""

            ' For each waybill record, add it to the interim table.
            If strInline.Length > 445 Then
                Write913SQL(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Raw, strInline)
            Else
                Write445SQL(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Raw, strInline)
            End If

            ' for each waybill load what we can to the 913 masked record
            If strInline.Length > 445 Then
                Write913SQL(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, strInline)
            Else
                Write913SQLFrom445(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, strInline)
            End If

            For mLooper = 1 To 10
                mRailRoads(mLooper) = 0
            Next

            For mLooper = 1 To 9
                mJunctions(mLooper) = ""
            Next

            For mLooper = 1 To 9
                mJCT_ST(mLooper) = ""
            Next

            ' Load the required field values to memvars
            If strInline.Length > 445 Then
                mSerial_No = Mid(strInline, 1, 6)
                mRptRR = Mid(strInline, 150, 3)
                mRate_Flg = Mid(strInline, 124, 1)
                mSTCC_W49 = Mid(strInline, 58, 7)
                mWbNum = Mid(strInline, 7, 6)
                mCar_Init = Mid(strInline, 31, 4)
                mAcctMonth = Mid(strInline, 21, 2)
                mAcctYear = Mid(strInline, 23, 4)
                mO_FSAC = Mid(strInline, 153, 5)
                mT_FSAC = Mid(strInline, 217, 5)
                mU_Rev = Mid(strInline, 83, 9)
                mWBDate = CDate(Mid(strInline, 13, 2) & "/" & Mid(strInline, 15, 2) & "/" & Mid(strInline, 17, 4))
                mUCarNum = Mid(strInline, 35, 6)
                mTCNum = Val(Mid(strInline, 52, 6))
                mCar_Typ = Convert_AAR_Car_Type_To_URCS_Car_Type(Mid(strInline, 286, 4))
                mCar_Typ = Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ(mCar_Typ)
                mTOFC_Serv_Code = Trim(Mid(strInline, 41, 3))
                mU_TC_Init = Trim(Mid(strInline, 48, 4))
                mStratum = Val(Mid(strInline, 121, 1))
                mU_TC_Units = Val(Mid(strInline, 44, 4))
                mU_Rev_Unmasked = 0
                mTons = 0

                mRailRoads(1) = Val(Trim(Mid(strInline, 158, 3)))       'ORR
                mRailRoads(2) = Val(Trim(Mid(strInline, 166, 3)))       'JRR1
                mRailRoads(3) = Val(Trim(Mid(strInline, 174, 3)))       'JRR2
                mRailRoads(4) = Val(Trim(Mid(strInline, 174, 3)))       'JRR3
                mRailRoads(5) = Val(Trim(Mid(strInline, 190, 3)))       'JRR4
                mRailRoads(6) = Val(Trim(Mid(strInline, 198, 3)))       'JRR5
                mRailRoads(7) = Val(Trim(Mid(strInline, 206, 3)))       'JRR6
                mRailRoads(8) = 0                                       'JRR7
                mRailRoads(9) = 0                                       'JRR8
                mRailRoads(10) = Val(Trim(Mid(strInline, 214, 3)))      'TRR

                mJunctions(1) = Trim(Mid(strInline, 161, 5))            'JCT1
                mJunctions(2) = Trim(Mid(strInline, 169, 5))            'JCT2
                mJunctions(3) = Trim(Mid(strInline, 177, 5))            'JCT3
                mJunctions(4) = Trim(Mid(strInline, 185, 5))            'JCT4
                mJunctions(5) = Trim(Mid(strInline, 193, 5))            'JCT5
                mJunctions(6) = Trim(Mid(strInline, 201, 5))            'JCT6
                mJunctions(7) = Trim(Mid(strInline, 209, 5))            'JCT7
                mJunctions(8) = 0                                       'JCT8
                mJunctions(9) = 0                                       'JCT9
            Else
                ' 445 Record
                mSerial_No = Mid(strInline, 1, 6)
                mRptRR = Mid(strInline, 146, 3)
                mRate_Flg = Mid(strInline, 120, 1)
                mSTCC_W49 = Mid(strInline, 54, 7)
                mWbNum = Mid(strInline, 7, 6)
                mCar_Init = Mid(strInline, 27, 1)
                mAcctMonth = Mid(strInline, 19, 2)
                mAcctYear = Mid(strInline, 21, 2)

                'Adjust mAcctYear as necessary
                If Val(mAcctYear) > 80 Then
                    mAcctYear = Val(Val(mAcctYear) + 1900).ToString
                Else
                    mAcctYear = Val(Val(mAcctYear) + 2000).ToString
                End If

                mO_FSAC = Mid(strInline, 149, 5)
                mT_FSAC = Mid(strInline, 229, 5)
                mU_Rev = Mid(strInline, 79, 9)
                mWBDate = CDate(Mid(strInline, 13, 2) & "/" & Mid(strInline, 15, 2) & "/" & Mid(strInline, 17, 2))
                mUCarNum = Mid(strInline, 31, 6)
                mTCNum = Val(Mid(strInline, 48, 6))
                mCar_Typ = Convert_AAR_Car_Type_To_URCS_Car_Type(Mid(strInline, 298, 4))
                mCar_Typ = Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ(mCar_Typ)
                mTOFC_Serv_Code = Trim(Mid(strInline, 37, 3))
                mU_TC_Init = Trim(Mid(strInline, 44, 4))
                mStratum = Val(Mid(strInline, 117, 1))
                mU_TC_Units = Val(Mid(strInline, 40, 4))
                mU_Rev_Unmasked = 0
                mTons = 0

                mRailRoads(1) = Val(Trim(Mid(strInline, 154, 3)))       'ORR
                mRailRoads(2) = Val(Trim(Mid(strInline, 162, 3)))       'JRR1
                mRailRoads(3) = Val(Trim(Mid(strInline, 170, 3)))       'JRR2
                mRailRoads(4) = Val(Trim(Mid(strInline, 178, 3)))       'JRR3
                mRailRoads(5) = Val(Trim(Mid(strInline, 186, 3)))       'JRR4
                mRailRoads(6) = Val(Trim(Mid(strInline, 194, 3)))       'JRR5
                mRailRoads(7) = Val(Trim(Mid(strInline, 202, 3)))       'JRR6
                mRailRoads(8) = Val(Trim(Mid(strInline, 210, 3)))       'JRR7
                mRailRoads(9) = Val(Trim(Mid(strInline, 218, 3)))       'JRR8
                mRailRoads(10) = Val(Trim(Mid(strInline, 226, 3)))      'TRR

                mJunctions(1) = Trim(Mid(strInline, 157, 5))            'JCT1
                mJunctions(2) = Trim(Mid(strInline, 165, 5))            'JCT2
                mJunctions(3) = Trim(Mid(strInline, 173, 5))            'JCT3
                mJunctions(4) = Trim(Mid(strInline, 181, 5))            'JCT4
                mJunctions(5) = Trim(Mid(strInline, 189, 5))            'JCT5
                mJunctions(6) = Trim(Mid(strInline, 197, 5))            'JCT6
                mJunctions(7) = Trim(Mid(strInline, 205, 5))            'JCT7
                mJunctions(8) = Trim(Mid(strInline, 213, 5))            'JCT8
                mJunctions(9) = Trim(Mid(strInline, 221, 5))            'JCT9
            End If

            ' Load the railroads, railroads alpha, junctions, states, and country into the masked table
            For mLooper = 1 To 10
                If mRailRoads(mLooper) <> 0 Then
                    Select Case mLooper
                        Case 1  'ORR
                            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "ORR", mRailRoads(mLooper).ToString)
                            mDatarow = mRRTable.Select("R260_Num = '" & mRailRoads(mLooper).ToString & "'")
                            mRR_Alpha = mDatarow(0)("Road_Mark")
                            Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "ORR_Alpha", mRR_Alpha.ToString)
                            mRR_Cntry = mDatarow(0)("RR_Country")
                            Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "ORR_Cntry", mRR_Cntry.ToString)
                        Case 10 'TRR
                            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "TRR", mRailRoads(mLooper).ToString)
                            mDatarow = mRRTable.Select("R260_Num = '" & mRailRoads(mLooper).ToString & "'")
                            mRR_Alpha = mDatarow(0)("Road_Mark")
                            Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "TRR_Alpha", mRR_Alpha.ToString)
                            mRR_Cntry = mDatarow(0)("RR_Country")
                            Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "TRR_Cntry", mRR_Cntry.ToString)
                        Case Else 'JRRs
                            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "JRR" & (mLooper - 1).ToString, mRailRoads(mLooper).ToString)
                            mDatarow = mRRTable.Select("R260_Num = '" & mRailRoads(mLooper).ToString & "'")
                            mRR_Alpha = mDatarow(0)("Road_Mark")
                            Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "JRR" & (mLooper - 1).ToString & "_Alpha", mRR_Alpha.ToString)
                            mRR_Cntry = mDatarow(0)("RR_Country")
                            Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "JRR" & (mLooper - 1).ToString & "_Cntry", mRR_Cntry.ToString)
                    End Select
                End If
            Next

            OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)

            'Convert hundredweight to weight in tons.
            If strInline.Length > 445 Then
                mU_Cars = Val(Mid(strInline, 27, 4))
                mBill_Wght = Val(Mid(strInline, 65, 9))
                mBill_Wght_Tons = Math.Round(mBill_Wght / 20, 0, MidpointRounding.AwayFromZero)
            Else
                mU_Cars = Val(Mid(strInline, 23, 4))
                mBill_Wght = Val(Mid(strInline, 61, 9))
                mBill_Wght_Tons = Math.Round(mBill_Wght / 20, 0, MidpointRounding.AwayFromZero)
            End If

            ' Update the Bill_Wght and Bill_Wght_Tons fields in the Masked table
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "Bill_Wght", mBill_Wght.ToString)
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "Bill_Wght_Tons", mBill_Wght_Tons.ToString)

            'Compute the tons per car.
            mTonsPerCar = Math.Round(mBill_Wght_Tons / mU_Cars, 0, MidpointRounding.AwayFromZero)

            'Determine the Int_Eq_Flg in the waybill
            mInt_Eq_Flg = 0

            If ((mCar_Typ = 46) Or (mCar_Typ = 48) Or (mCar_Typ = 49) Or (mCar_Typ = 52)) And (((mTonsPerCar >= 17) And (mTonsPerCar <= 23)) _
                    Or ((mTonsPerCar >= 34) And (mTonsPerCar <= 46))) And (mU_TC_Init = ””) And (mTOFC_Serv_Code = ””) Then
                ' Might be Intermodal

                mInt_Eq_Flg = 1

            ElseIf ((mCar_Typ = 46) Or (mCar_Typ = 48) Or (mCar_Typ = 49) Or (mCar_Typ = 52)) And ((mCar_Init = “TCSZ”) Or (mU_TC_Init = ”TCSZ”) Or (mCar_Init = “NERZ”) Or (mU_TC_Init = ”NERZ”)) And (mTOFC_Serv_Code <> “”) Then
                ' RoadRailer

                mInt_Eq_Flg = 2


            ElseIf ((mCar_Typ = 46) Or (mCar_Typ = 48) Or (mCar_Typ = 49) Or (mCar_Typ = 52)) And (mU_TC_Init <> ””) And (mTOFC_Serv_Code <> ””) Then
                ' Intermodal

                mInt_Eq_Flg = 3

            End If

            ' Determine JF value
            mJF = 0
            For mLooper = 1 To 9
                If mJunctions(mLooper) <> "" Then
                    mJF = mJF + 1
                End If
            Next

            ' Update the Int_Eq_Flg field in the Masked table
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "Int_Eq_Flg", mInt_Eq_Flg.ToString)
            ' Update the jf field in the Masked table
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "JF", mJF.ToString)

            'Update the STB_Car_Type value in the Masked table
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "STB_Car_Typ", mCar_Typ.ToString)

            If Len(mAcctMonth) < 2 Then
                mAcctMonth = "0" & mAcctMonth
            End If

            Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "Acct_Period", mAcctMonth.ToString & mAcctYear.ToString)

            ' Zero out the Revenue, Distance, Variable Costs, and NET fields
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "ORR_REV", "0")
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "ONET", "0")

            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "ORR_Dist", "0")
            For mLooper = 1 To 6
                Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "JRR" & mLooper.ToString & "_Rev", "0")
                Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "JRR" & mLooper.ToString & "_Dist", "0")
            Next
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "TRR_Dist", "0")
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "Total_Dist", "0")

            For mLooper = 1 To 7
                Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "NET" & mLooper.ToString, "0")
            Next

            For mLooper = 1 To 8
                Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "RR" & mLooper.ToString & "_VC", 0)
            Next

            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "Total_VC", "0")
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "TRR_REV", "0")
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "TNET", "0")

            ' Let's not tie up the user computer
            Application.DoEvents()

            'Build the Unmasked Revenue records
            If mRate_Flg = 1 Then

                Select Case CInt(mRptRR)

                    Case 105, 482    'CP
                        If Val(txt_Year.Text) < 1992 Then
                            'Do nothing - they didn't mask rates prior to Year 1992
                        Else
                            mU_Rev_Unmasked = UnmaskCPValue(mSerial_No, mSTCC_W49, mU_Rev)
                        End If

                    Case 131    'CNW
                        If Val(txt_Year.Text) < 1989 Then
                            'Do nothing - they didn't mask rates prior to Year 1989
                        Else
                            mU_Rev_Unmasked = UnmaskCNWValue(mSerial_No, mTState, mUCars, mU_Rev)
                        End If

                    Case 190    'Conrail
                        If Val(txt_Year.Text) < 1987 Then
                            'Do nothing - they didn't mask rates prior to Year 1987
                        Else
                            mU_Rev_Unmasked = UnmaskConrailValue(mSerial_No, mSTCC_W49, mJF, mU_Rev)
                        End If

                    Case 555    'NS
                        If Val(txt_Year.Text) < 1990 Then
                            'Do nothing - they didn't mask rates prior to Year 1990
                        Else
                            mU_Rev_Unmasked = UnmaskNSValue(mSerial_No, mWBDate, mAcctYear, mAcctMonth, mO_FSAC, mWbNum, mUCarNum, mTCNum, mU_Rev)
                        End If

                    Case 712    'CSXT
                        If Val(txt_Year.Text) < 1990 Then
                            'Do nothing - they didn't mask rates prior to Year 1990
                        Else
                            mU_Rev_Unmasked = UnmaskCSXValue(mSerial_No, mSTCC_W49, mWbNum, mAcctMonth, mAcctYear, mU_Rev)
                        End If

                    Case 777    'BNSF
                        If Val(txt_Year.Text) < 2000 Then
                            'Do nothing - they didn't mask rates prior to Year 2000
                        Else
                            mU_Rev_Unmasked = UnmaskBNSFValue(mSerial_No, mCar_Init, mAcctMonth, mU_Rev)
                        End If

                    Case 802  'UP
                        If Val(txt_Year.Text) < 1990 Then
                            'Do nothing - they didn't mask rates prior to Year 1990
                        Else
                            mU_Rev_Unmasked = UnmaskUPValue(mSerial_No, Val(mSTCC_W49), Year(mWBDate), mWbNum, mU_Rev)
                        End If

                    Case Else   'All other railroads                        
                        If Val(txt_Year.Text) < 2001 Then
                            'Do nothing - they didn't mask rates prior to Year 2001
                        Else
                            mU_Rev_Unmasked = UnmaskGenericValue(mSerial_No, mSTCC_W49, mWBDate, mAcctYear, mU_Rev)
                        End If
                End Select

            End If

            OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)

            'start building the SQL statement to run later
            msqlCommand = New SqlCommand
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.CommandType = CommandType.Text

            mSQLCommandString = New StringBuilder
            mSQLCommandString.Append("INSERT INTO " & Gbl_Interim_Unmasked_Rev & "(")
            mSQLCommandString.Append("[unmasked_serial_no], [u_rev_unmasked]) VALUES (")
            mSQLCommandString.Append("'" & mSerial_No & "', ")
            mSQLCommandString.Append(mU_Rev_Unmasked.ToString & ")")
            msqlCommand.CommandText = mSQLCommandString.ToString
            msqlCommand.ExecuteNonQuery()

            ' Determine the expansion factor
            ' default to no expansion (this should never happen)
            mExp_Factor_Th = 1

            If mAcctYear <= 2020 Then
                Select Case mStratum
                    Case 1
                        mExp_Factor_Th = 40
                    Case 2
                        mExp_Factor_Th = 12
                    Case 3
                        mExp_Factor_Th = 4
                    Case 4
                        mExp_Factor_Th = 3
                    Case 5
                        mExp_Factor_Th = 2
                End Select
            Else  ' New sampling rates from EP 385 Sub 8 effective January 1, 2021
                Select Case mStratum
                    Case 0  ' This shouldn’t happen, but the contractor is having problems implementing the changes
                        mExp_Factor_Th = 1
                        mStratumErrors = mStratumErrors + 1
                    Case 1
                        mExp_Factor_Th = 5
                    Case 2
                        mExp_Factor_Th = 5
                    Case 3
                        mExp_Factor_Th = 4
                    Case 4
                        mExp_Factor_Th = 3
                    Case 5
                        mExp_Factor_Th = 2
                    Case 6
                        mExp_Factor_Th = 40
                    Case 7
                        mExp_Factor_Th = 5
                End Select
            End If

            'If mExp_Factor_Th = 1 Then
            '    MsgBox("Serial No " & mSerial_No & " has an invalid Stratum value.  Aborting run.", vbOKOnly, "ERROR!")
            '    GoTo EndIt
            'End If

            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "Exp_Factor_Th", mExp_Factor_Th.ToString)

            mU_Rev = Val(Get_Field_Value(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "U_Rev", "Serial_No", mSerial_No))
            mU_Cars = Val(Get_Field_Value(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "U_Cars", "Serial_No", mSerial_No))
            mU_TC_Units = Val(Get_Field_Value(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "U_TC_Units", "Serial_No", mSerial_No))
            mBill_Wght = Val(Get_Field_Value(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Bill_Wght", "Serial_No", mSerial_No))
            mTons = mExp_Factor_Th * mBill_Wght_Tons
            mU_Rev_Unmasked = Val(Get_Field_Value(Gbl_Interim_Waybills_Database_Name, Gbl_Unmasked_Rev_TableName, "U_Rev_Unmasked", "Unmasked_Serial_No", mSerial_No))
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Unmasked_Rev_TableName, "Unmasked_Serial_no", mSerial_No, "Total_Unmasked_Rev", (mU_Rev_Unmasked * mExp_Factor_Th).ToString)

            ' Update the fields that the Exp_Factor_Th is used to compute their values

            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "Total_Rev", (mExp_Factor_Th * mU_Rev).ToString)
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "Cars", (mExp_Factor_Th * mU_Cars).ToString)
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "TC_Units", (mExp_Factor_Th * mU_TC_Units).ToString)
            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "Tons", mTons.ToString)

LoadSegments:

            'Begin work for segments

            mNumberOfRRs = mJF + 1

            For mLooper = 1 To mNumberOfRRs
                If strInline.Length > 445 Then
                    mUCarNum = Mid(strInline, 35, 6)
                    mTCNum = Val(Mid(strInline, 52, 6))
                    mUCars = Val(Mid(strInline, 27, 4))
                    mU_Rev = Val(Mid(strInline, 83, 9))
                    mWBDate = Mid(strInline, 13, 2).ToString & "/" & Mid(strInline, 15, 2) & "/" & Mid(strInline, 17, 4)
                    mYear = txt_Year.Text
                    mRebill = Val(Mid(strInline, 120, 1))
                    mORR = Val(Mid(strInline, 158, 3))
                    mTRR = Val(Mid(strInline, 214, 3))
                    mSTCC2 = Mid(mSTCC_W49, 1, 2)
                    mSTCC3 = Mid(mSTCC_W49, 1, 3)
                    mSTCC4 = Mid(mSTCC_W49, 1, 4)
                    mSTCC5 = Mid(mSTCC_W49, 1, 5)
                Else
                    mUCarNum = Mid(strInline, 31, 6)
                    mTCNum = Val(Mid(strInline, 48, 6))
                    mUCars = Val(Mid(strInline, 23, 4))
                    mU_Rev = Val(Mid(strInline, 79, 9))
                    mWBDate = Mid(strInline, 13, 2).ToString & "/" & Mid(strInline, 15, 2) & "/" & Mid(strInline, 17, 2)
                    mYear = txt_Year.Text
                    mRebill = Val(Mid(strInline, 116, 1))
                    mORR = Val(Mid(strInline, 154, 3))
                    mTRR = Val(Mid(strInline, 226, 3))
                    mSTCC2 = Mid(mSTCC_W49, 1, 2)
                    mSTCC3 = Mid(mSTCC_W49, 1, 3)
                    mSTCC4 = Mid(mSTCC_W49, 1, 4)
                    mSTCC5 = Mid(mSTCC_W49, 1, 5)
                End If

                mRR_Cntry = ""
                mOriginType = ""
                mDestinationType = ""
                mSeg_Type = ""

                'Determine the Route Formula
                'default to Practical.  BatchPro practical value is 0 
                mRoute_Formula = 0

                If ((mCar_Typ = 46) Or (mCar_Typ = 48) Or (mCar_Typ = 49) Or (mCar_Typ = 52)) And (mTOFC_Serv_Code <> “”) Then
                    ' Intermodal                
                    mRoute_Formula = 2

                ElseIf (((mSTCC2 = 11) Or (mSTCC3 = 101) Or (mSTCC5 = 29913) Or (mSTCC5 = 29914)) Or
                    ((mSTCC4 = 113 And mCar_Typ = 41))) And (mUCars >= 50) Then
                    ' Coal/Bulk:  Coal, Iron Ore, Coke, or Grain in Covered Hoppers and in 50+ car shipments
                    mRoute_Formula = 3

                ElseIf (mCar_Typ = 47) And ((mSTCC5 = 37111) Or (mSTCC5 = 37112)) Then
                    ' AutoRack shipping passenger cars or trucks
                    mRoute_Formula = 4

                End If

                ' Update the Service_Type field
                If mRoute_Formula = 0 Then
                    mServiceType = 1
                Else
                    mServiceType = mRoute_Formula
                End If

                Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "Service_Type", mServiceType.ToString)

                'Create the line in the BatchProInFile
                mBPInLine = New StringBuilder
                mBPInLine.Append(mSerial_No & " ")
                mBPInLine.Append(mLooper.ToString & " ")
                mBPInLine.Append(mNumberOfRRs.ToString & " ")

                'Determine the segment type
                mSeg_No = mLooper
                mTotal_Segs = mNumberOfRRs
                If (mSeg_No = 1) And (mNumberOfRRs = 1) Then
                    Select Case mRebill
                        Case 0
                            mSeg_Type = “OT”
                        Case 1
                            mSeg_Type = “OD”
                        Case 2
                            mSeg_Type = “RD”
                        Case 3
                            mSeg_Type = “RT”
                    End Select

                    mRR_Num = mORR

                    If mRR_Num <> 0 Then
                        mDatarow = mRRTable.Select("R260_Num = " & mRR_Num)
                        mRR_Cntry = mDatarow(0)("RR_Country")
                        mRR_Alpha = mDatarow(0)("Road_Mark")
                    Else
                        mRR_Cntry = "XX"
                        mRR_Alpha = "XXXX"
                    End If

                    mFrom_Loc = Zero_Fill_Numeric(mRR_Num, 3).ToString & "-" & Zero_Fill_Numeric(mO_FSAC, 5).ToString
                    mDatarow = mFSACTable.Select("RR_FSAC = '" & mFrom_Loc & "'")
                    If mDatarow.Length = 0 Then
                        mFrom_ST = "XX"
                        mFrom_Lat = "0"
                        mFrom_Long = "0"
                    Else
                        mFrom_ST = mDatarow(0)("Loc_St")
                        mFrom_Lat = mDatarow(0)("nLatitude")
                        mFrom_Long = mDatarow(0)("nLongitude")
                        Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "O_ST", mFrom_ST.ToString)
                    End If
                    mOriginType = "FSAC"

                    mTo_Loc = Zero_Fill_Numeric(mRR_Num, 3).ToString & "-" & Zero_Fill_Numeric(mT_FSAC, 5).ToString
                    mDatarow = mFSACTable.Select("RR_FSAC = '" & mTo_Loc & "'")
                    If mDatarow.Length = 0 Then
                        mTo_ST = "XX"
                        mTo_Lat = "0"
                        mTo_Long = "0"
                    Else
                        mTo_ST = mDatarow(0)("Loc_St")
                        mTo_Lat = mDatarow(0)("nLatitude")
                        mTo_Long = mDatarow(0)("nLongitude")
                        Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "T_ST", mTo_ST.ToString)
                    End If

                    mDestinationType = "FSAC"

                    ' Continue building the BatchPro input file line for FSAC to FSAC
                    mBPInLine.Append(mSeg_Type).Append(" ")
                    mBPInLine.Append(mRoute_Formula.ToString).Append(" ")
                    mBPInLine.Append("0", (3 - Len(mRR_Num.ToString))).Append((mRR_Num.ToString)).Append(" ")
                    mBPInLine.Append(mRR_Alpha.ToString).Append(Space(5 - Len(mRR_Alpha.ToString)))
                    mBPInLine.Append("2").Append(" ")
                    mBPInLine.Append("0", 5 - Len(mO_FSAC.ToString)).Append(mO_FSAC.ToString).Append(" ")
                    mBPInLine.Append("2").Append(" ")
                    mBPInLine.Append("0", 5 - Len(mT_FSAC.ToString)).Append(mT_FSAC.ToString).Append(" ")

                ElseIf (mSeg_No = 1) And (mNumberOfRRs > 1) Then
                    Select Case mRebill
                        Case 0
                            mSeg_Type = “OD”
                        Case 1
                            mSeg_Type = “OD”
                        Case 2
                            mSeg_Type = “RD”
                        Case 3
                            mSeg_Type = “RD”
                    End Select

                    mRR_Num = mORR

                    If mRR_Num <> 0 Then
                        mDatarow = mRRTable.Select("R260_Num = " & mRR_Num)
                        If mDatarow.Length = 0 Then
                            mRR_Cntry = "XX"
                            mRR_Alpha = "XXXX"
                        Else
                            mRR_Cntry = mDatarow(0)("RR_Country")
                            mRR_Alpha = mDatarow(0)("Road_Mark")
                        End If
                    Else
                        mRR_Cntry = "XX"
                        mRR_Alpha = "XXXX"
                    End If

                    mFrom_Loc = Zero_Fill_Numeric(mRR_Num, 3).ToString & "-" & Zero_Fill_Numeric(mO_FSAC, 5).ToString
                    mDatarow = mFSACTable.Select("RR_FSAC = '" & mFrom_Loc & "'")
                    If mDatarow.Length = 0 Then
                        mFrom_ST = ""
                        mFrom_Lat = ""
                        mFrom_Long = ""
                    Else
                        mFrom_ST = mDatarow(0)("Loc_ST")
                        mFrom_Lat = mDatarow(0)("nLatitude")
                        mFrom_Long = mDatarow(0)("nLongitude")
                        Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "O_ST", mFrom_ST)
                    End If
                    mOriginType = "FSAC"

                    mTo_Loc = mJunctions(mSeg_No)
                    mDatarow = mJCTTable.Select("Rule_260 = '" & mTo_Loc & "'")
                    If mDatarow.Length = 0 Then
                        mTo_ST = "XX"
                        mTo_Lat = "0"
                        mTo_Long = "0"
                    Else
                        mTo_ST = mDatarow(0)("Loc_ST")
                        mTo_Lat = mDatarow(0)("nLatitude")
                        mTo_Long = mDatarow(0)("nLongitude")
                    End If

                    Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "T_ST", mTo_ST)
                    mDestinationType = "JCT"

                    ' Continue building the BatchPro input file line for FSAC to JCT
                    mBPInLine.Append(mSeg_Type).Append(" ")
                    mBPInLine.Append(mRoute_Formula.ToString).Append(" ")
                    mBPInLine.Append("0", (3 - Len(mRR_Num.ToString))).Append((mRR_Num.ToString)).Append(" ")
                    mBPInLine.Append(mRR_Alpha.ToString).Append(Space(5 - Len(mRR_Alpha.ToString)))
                    mBPInLine.Append("2").Append(" ")
                    mBPInLine.Append("0", 5 - Len(mO_FSAC.ToString)).Append(mO_FSAC.ToString).Append(" ")
                    mBPInLine.Append("5").Append(" ")
                    mBPInLine.Append(mTo_Loc.ToString).Append(" ", 6 - Len(mTo_Loc.ToString))

                ElseIf (mSeg_No > 1) And (mSeg_No < mNumberOfRRs) Then
                    Select Case mRebill
                        Case 0
                            mSeg_Type = “RD”
                        Case 1
                            mSeg_Type = “RD”
                        Case 2
                            mSeg_Type = “RD”
                        Case 3
                            mSeg_Type = “RD”
                    End Select

                    mRR_Num = mRailRoads(mSeg_No)

                    If mRR_Num <> 0 Then
                        mDatarow = mRRTable.Select("R260_Num = " & mRR_Num)
                        If mDatarow.Length = 0 Then
                            mRR_Cntry = "XX"
                            mRR_Alpha = "XXXX"
                        Else
                            mRR_Cntry = mDatarow(0)("RR_Country")
                            mRR_Alpha = mDatarow(0)("Road_Mark")
                        End If
                    Else
                        mRR_Cntry = "XX"
                        mRR_Alpha = "XXXX"
                    End If

                    mFrom_Loc = mJunctions(mSeg_No - 1)
                    mDatarow = mJCTTable.Select("Rule_260 = '" & mFrom_Loc & "'")
                    If mDatarow.Length = 0 Then
                        mFrom_ST = "XX"
                        mFrom_Lat = "0"
                        mFrom_Long = "0"
                    Else
                        mFrom_ST = mDatarow(0)("Loc_ST")
                        mFrom_Lat = mDatarow(0)("nLatitude")
                        mFrom_Long = mDatarow(0)("nLongitude")
                    End If

                    Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "JCT" & (mSeg_No - 1).ToString & "_ST", mFrom_ST.ToString)
                    mOriginType = "JCT"

                    mTo_Loc = mJunctions(mSeg_No)
                    mDatarow = mJCTTable.Select("Rule_260 = '" & mTo_Loc & "'")
                    If mDatarow.Length = 0 Then
                        mTo_ST = "XX"
                        mTo_Lat = "0"
                        mTo_Long = "0"
                    Else
                        mTo_ST = mDatarow(0)("Loc_ST")
                        mTo_Lat = mDatarow(0)("nLatitude")
                        mTo_Long = mDatarow(0)("nLongitude")
                    End If

                    Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "JCT" & mSeg_No.ToString & "_ST", mFrom_ST.ToString)
                    mDestinationType = "JCT"

                    ' Continue building the BatchPro input file line for JCT to JCT
                    mBPInLine.Append(mSeg_Type).Append(" ")
                    mBPInLine.Append(mRoute_Formula.ToString).Append(" ")
                    mBPInLine.Append("0", (3 - Len(mRR_Num.ToString))).Append((mRR_Num.ToString)).Append(" ")
                    mBPInLine.Append(mRR_Alpha.ToString).Append(Space(5 - Len(mRR_Alpha.ToString)))
                    mBPInLine.Append("5 ")
                    mBPInLine.Append(mFrom_Loc.ToString).Append(" ", 5 - Len(mFrom_Loc.ToString)).Append(" ")
                    mBPInLine.Append("5 ")
                    mBPInLine.Append(mTo_Loc.ToString).Append(" ", 5 - Len(mTo_Loc.ToString)).Append(" ")

                ElseIf (mSeg_No > 1) And (mSeg_No = mNumberOfRRs) Then
                    Select Case mRebill
                        Case 0
                            mSeg_Type = “RT”
                        Case 1
                            mSeg_Type = “RD”
                        Case 2
                            mSeg_Type = “RD”
                        Case 3
                            mSeg_Type = “RT”
                    End Select

                    mRR_Num = mTRR

                    If mRR_Num <> 0 Then
                        mDatarow = mRRTable.Select("R260_Num = " & mRR_Num)
                        mRR_Cntry = mDatarow(0)("RR_Country")
                        mRR_Alpha = mDatarow(0)("Road_Mark")
                    Else
                        mRR_Cntry = "XX"
                        mRR_Alpha = "XXXX"
                    End If

                    mFrom_Loc = mJunctions(mSeg_No - 1)
                    mDatarow = mJCTTable.Select("Rule_260 = '" & mFrom_Loc & "'")
                    If mDatarow.Length = 0 Then
                        mFrom_ST = "XX"
                        mFrom_Lat = "0"
                        mFrom_Long = "0"
                    Else
                        mFrom_ST = mDatarow(0)("Loc_ST")
                        mFrom_Lat = mDatarow(0)("nLatitude")
                        mFrom_Long = mDatarow(0)("nLongitude")
                    End If

                    Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_no", mSerial_No, "JCT" & (mSeg_No - 1).ToString & "_ST", mFrom_ST.ToString)
                    mOriginType = "JCT"

                    mDatarow = mRRTable.Select("R260_Num = " & mRR_Num)
                    mTo_Loc = Zero_Fill_Numeric(mRR_Num, 3).ToString & "-" & Zero_Fill_Numeric(mT_FSAC, 5).ToString
                    mDatarow = mFSACTable.Select("RR_FSAC = '" & mTo_Loc & "'")
                    If mDatarow.Length = 0 Then
                        mTo_ST = ""
                        mTo_Lat = ""
                        mTo_Long = ""
                    Else
                        mTo_ST = mDatarow(0)("Loc_St")
                        mTo_Lat = mDatarow(0)("nLatitude")
                        mTo_Long = mDatarow(0)("nLongitude")
                        Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name, Gbl_Interim_Masked, "Serial_No", mSerial_No, "T_ST", mTo_ST.ToString)
                    End If
                    mDestinationType = "FSAC"

                    ' Continue building the BatchPro input file line for JCT to FSAC
                    mBPInLine.Append(mSeg_Type).Append(" ")
                    mBPInLine.Append(mRoute_Formula.ToString).Append(" ")
                    mBPInLine.Append("0", (3 - Len(mRR_Num.ToString))).Append((mRR_Num.ToString)).Append(" ")
                    mBPInLine.Append(mRR_Alpha.ToString).Append(Space(5 - Len(mRR_Alpha.ToString)))
                    mBPInLine.Append("5").Append(" ")
                    mBPInLine.Append(mFrom_Loc.ToString).Append(Space(6 - Len(mFrom_Loc.ToString)))
                    mBPInLine.Append("2").Append(" ")
                    mBPInLine.Append("0", 5 - Len(mT_FSAC.ToString)).Append(mT_FSAC.ToString).Append(" ")

                End If

                'write the line to the ALL IN file
                sBPAllInFilename.WriteLine(mBPInLine.ToString)

                ' Change the route formula for the all file to 0
                mWorkStr = mBPInLine.ToString
                mWorkStr = Mid(mWorkStr, 1, 14) & "0" & Mid(mWorkStr, 16)

                'Write the line to the appropriate .IN text file
                Select Case mRoute_Formula
                    Case 0  'Practical/GM
                        If mOriginType = "FSAC" Then    '2
                            Select Case mDestinationType
                                Case "FSAC" '2 - FSAC2FSAC
                                    sBPGM_Streams(1).WriteLine(mBPInLine.ToString)
                                Case "JCT"  '5 - FSAC2JCT
                                    sBPGM_Streams(2).WriteLine(mBPInLine.ToString)
                            End Select
                        Else  '5
                            Select Case mDestinationType
                                Case "FSAC" '2 - JCT2FSAC
                                    sBPGM_Streams(3).WriteLine(mBPInLine.ToString)
                                Case "JCT"  '5 - JCT2JCT
                                    sBPGM_Streams(4).WriteLine(mBPInLine.ToString)
                            End Select
                        End If
                    Case 2  'Intermodal
                        If mOriginType = "FSAC" Then    '2
                            Select Case mDestinationType
                                Case "FSAC" '2 - FSAC2FSAC
                                    sBPIM_Streams(1).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(1).WriteLine(mWorkStr)
                                Case "JCT"  '5 - FSAC2JCT
                                    sBPIM_Streams(2).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(2).WriteLine(mWorkStr)
                            End Select
                        Else  '5
                            Select Case mDestinationType
                                Case "FSAC" '2 - JCT2FSAC
                                    sBPIM_Streams(3).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(3).WriteLine(mWorkStr)
                                Case "JCT"  '5 - JCT2JCT
                                    sBPIM_Streams(4).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(4).WriteLine(mWorkStr)
                            End Select
                        End If
                    Case 3  'Coal/Bulk
                        If mOriginType = "FSAC" Then    '2
                            Select Case mDestinationType
                                Case "FSAC" '2 - FSAC2FSAC
                                    sBPCB_Streams(1).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(1).WriteLine(mWorkStr)
                                Case "JCT"  '5 - FSAC2JCT
                                    sBPCB_Streams(2).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(2).WriteLine(mWorkStr)
                            End Select
                        Else  '2
                            Select Case mDestinationType
                                Case "FSAC" '2 - JCT2FSAC
                                    sBPCB_Streams(3).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(3).WriteLine(mWorkStr)
                                Case "JCT"  '5 - JCT2JCT
                                    sBPCB_Streams(4).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(4).WriteLine(mWorkStr)
                            End Select
                        End If
                    Case 4  'Auto Rack
                        If mOriginType = "FSAC" Then '2
                            Select Case mDestinationType
                                Case "FSAC" '2 - FSAC2FSAC
                                    sBPAR_Streams(1).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(1).WriteLine(mWorkStr)
                                Case "JCT"  '5 - FSAC2JCT
                                    sBPAR_Streams(2).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(2).WriteLine(mWorkStr)
                            End Select
                        Else  '5
                            Select Case mDestinationType
                                Case "FSAC" '2 - JCT2FSAC
                                    sBPAR_Streams(3).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(3).WriteLine(mWorkStr)
                                Case "JCT"  '5 - JCT2JCT
                                    sBPAR_Streams(4).WriteLine(mBPInLine.ToString)
                                    sBPGM_Streams(4).WriteLine(mWorkStr)
                            End Select
                        End If
                End Select

                'Start building the insert SQL command to create the segments record
                mSQLCommandString = New StringBuilder
                With mSQLCommandString
                    .Append("INSERT INTO " & Gbl_Interim_Segments & "(")
                    .Append("[serial_no], [seg_no], [total_segs], [RR_Num], ")
                    .Append("[RR_Alpha], [RR_Dist], [RR_Cntry], [RR_Rev], [RR_VC], ")
                    .Append("[Seg_Type], [From_Node], [To_Node], [From_Loc], ")
                    .Append("[From_St], [To_Loc], [To_St], [From_Latitude], ")
                    .Append("[From_Longitude], [To_Latitude], [To_Longitude]) VALUES (")

                    .Append("'" & mSerial_No & "',")                                       'SerialNo
                    .Append(mLooper.ToString & ",")                                        'Seg_no
                    .Append(mNumberOfRRs.ToString & ",")                                   'Total_Segs
                    .Append(mRR_Num.ToString & ",")                                        'RR_num
                    .Append("'" & Get_Field_Value(Gbl_Waybill_Database_Name, mList_of_RRs_TableName, "Road_Mark", "R260_Num", mRR_Num.ToString) & "',") 'RR_Alpha
                    .Append("0,")                                                          'RR_Dist
                    .Append("'" & mRR_Cntry.ToString & "',")                               'RR_Cntry
                    .Append("0,")                                                          'RR_Rev
                    .Append("0,")                                                          'RR_VC
                    .Append("'" & mSeg_Type.ToString & "',")                                'Seg_Type
                    .Append("0,")                                                          'From_Node
                    .Append("0,")                                                          'To_Node
                    .Append("'" & mFrom_Loc.ToString & "',")                               'From_Loc
                    .Append("'" & mFrom_ST.ToString & "',")                                'From_St
                    .Append("'" & mTo_Loc.ToString & "',")                                 'To_Loc
                    .Append("'" & mTo_ST.ToString & "',")                                   'to_St
                    .Append(ReturnDecimal(mFrom_Lat).ToString & ",")                                'From_Latitude
                    .Append(ReturnDecimal(mFrom_Long).ToString & ",")                               'From_Longitude
                    .Append(ReturnDecimal(mTo_Lat).ToString & ",")                                  'To_Latitude
                    .Append(ReturnDecimal(mTo_Long).ToString & ")")                                 'To_Longitude
                End With

                OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)

                msqlCommand = New SqlCommand
                msqlCommand.CommandType = CommandType.Text
                msqlCommand.Connection = gbl_SQLConnection
                msqlCommand.CommandText = mSQLCommandString.ToString
                msqlCommand.ExecuteNonQuery()

                OpenSQLConnection(My.Settings.InterimWaybills)

                'start building the SQL statement to load the unmasked segments data
                mSQLCommandString = New StringBuilder
                mSQLCommandString.Append("INSERT INTO " & Gbl_Interim_Unmasked_Segments & "(")
                mSQLCommandString.Append("[Serial_No], ")
                mSQLCommandString.Append("[Seg_no], ")
                mSQLCommandString.Append("[RR_Unmasked_Rev]) VALUES (")

                msqlCommand = New SqlCommand
                msqlCommand.Connection = gbl_SQLConnection
                msqlCommand.CommandType = CommandType.Text

                mSQLCommandString.Append("'" & mSerial_No & "', ")
                mSQLCommandString.Append(mLooper.ToString & ", ")

                ' Commented out until method of calculation is updated
                '=====================================================
                'If mLooper = 1 Then
                '    mSQLCommandString.Append(mU_Rev.ToString & ")")
                'Else
                mSQLCommandString.Append("0)")
                'End If

                msqlCommand.CommandText = mSQLCommandString.ToString
                msqlCommand.ExecuteNonQuery()

            Next

        End While

        'Log the time hack

        mLoad_Ends = Now()

        OutString = New StringBuilder
        OutString.Append("Masked Data Table Load complete at " & mLoad_Ends.ToString("G") & " ")
        sbLogFile.WriteLine(OutString)
        OutString = New StringBuilder
        OutString.Append(Return_Elapsed_Time(mLoad_Starts, mLoad_Ends))
        sbLogFile.WriteLine(OutString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        'Flush and Close the text files

        For mLooper = 1 To 4
            sBPGM_Streams(mLooper).Flush()
            sBPGM_Streams(mLooper).Close()
            sBPIM_Streams(mLooper).Flush()
            sBPIM_Streams(mLooper).Close()
            sBPCB_Streams(mLooper).Flush()
            sBPCB_Streams(mLooper).Close()
            sBPAR_Streams(mLooper).Flush()
            sBPAR_Streams(mLooper).Close()
        Next

        sBPAllInFilename.Flush()
        sBPAllInFilename.Close()

        mSQLCommandString = New StringBuilder
        mSQLCommandString.Append("Run Errors")
        mSQLCommandString.AppendLine()
        mSQLCommandString.Append("Stratum Errors: " & mStratumErrors.ToString)

        sRailIncErrorStream.Write(mSQLCommandString)
        sRailIncErrorStream.Flush()
        sRailIncErrorStream.Close()

        ' Close the reader
        sSr.Close()

Skip_Masked_Data_Load:

BatchPro:

        'Write the time hack to the log file
        mBatchPro_Starts = Now()

        ' Call the Batch Pro program for each filetype.  Note that BatchPro does not work if drive letters are referenced, thus "Mid(Work_Directory, 3))"
        For mLooper = 1 To 4
            Select Case mLooper
                Case 1  'GM runs
                    For mLooper2 = 1 To 4
                        txt_StatusBox.Text = "Working - Computing Mileage for GM " & mBatchPro_Runs(mLooper2) & " Records - Step " & mLooper2.ToString & " of 4."
                        Refresh()
                        Application.DoEvents()
                        mProcessParameters = New StringBuilder
                        With mProcessParameters
                            .Append(" -input:" & Chr(34) & Mid(Work_Directory, 3))
                            .Append(mBPGM_FileNames(mLooper2) & mInFileNameSuffix & Chr(34) & " ")
                            .Append(" -config:" & Chr(34) & Mid(Work_Directory, 3))
                            .Append(mBPGM_FileTypes(mLooper2) & mCfgFileNameSuffix & Chr(34))
                            mProcess.Start(My.Settings.BatchProLoc, .ToString).WaitForExit()
                        End With
                    Next mLooper2

                Case 2  'IM Runs
                    For mLooper2 = 1 To 4
                        txt_StatusBox.Text = "Working - Computing Mileage for IM " & mBatchPro_Runs(mLooper2) & " Records - Step " & mLooper2.ToString & " of 4."
                        Refresh()
                        Application.DoEvents()
                        mProcessParameters = New StringBuilder
                        With mProcessParameters
                            .Append(" -input:" & Chr(34) & Mid(Work_Directory, 3))
                            .Append(mBPIM_Filenames(mLooper2) & mInFileNameSuffix & Chr(34) & " ")
                            .Append(" -config:" & Chr(34) & Mid(Work_Directory, 3))
                            .Append(mBPIM_FileTypes(mLooper2) & mCfgFileNameSuffix & Chr(34))
                            mProcess.Start(My.Settings.BatchProLoc, .ToString).WaitForExit()
                        End With
                    Next mLooper2

                Case 3  'CB Runs
                    For mLooper2 = 1 To 4
                        txt_StatusBox.Text = "Working - Computing Mileage for CB " & mBatchPro_Runs(mLooper2) & " Records - Step " & mLooper2.ToString & " of 4."
                        Refresh()
                        Application.DoEvents()
                        mProcessParameters = New StringBuilder
                        With mProcessParameters
                            .Append(" -input:" & Chr(34) & Mid(Work_Directory, 3))
                            .Append(mBPCB_Filenames(mLooper2) & mInFileNameSuffix & Chr(34) & " ")
                            .Append(" -config:" & Chr(34) & Mid(Work_Directory, 3))
                            .Append(mBPCB_FileTypes(mLooper2) & mCfgFileNameSuffix & Chr(34))
                            mProcess.Start(My.Settings.BatchProLoc, .ToString).WaitForExit()
                        End With
                    Next mLooper2

                Case 4  'AR Runs
                    For mLooper2 = 1 To 4
                        txt_StatusBox.Text = "Working - Computing Mileage for AR " & mBatchPro_Runs(mLooper2) & " Records - Step " & mLooper2.ToString & " of 4."
                        Refresh()
                        Application.DoEvents()
                        mProcessParameters = New StringBuilder
                        With mProcessParameters
                            .Append(" -input:" & Chr(34) & Mid(Work_Directory, 3))
                            .Append(mBPAR_Filenames(mLooper2) & mInFileNameSuffix & Chr(34) & " ")
                            .Append(" -config:" & Chr(34) & Mid(Work_Directory, 3))
                            .Append(mBPAR_FileTypes(mLooper2) & mCfgFileNameSuffix & Chr(34))
                            mProcess.Start(My.Settings.BatchProLoc, .ToString).WaitForExit()
                        End With
                    Next mLooper2
            End Select

        Next mLooper


        ' Build the ALL.OUT file and also save the data to the mBPAllOut data table.
        'Open the OUT File in overwrite mode
        sBPAllOutFile = New StreamWriter(mBPAllOutFilename, False)

        ' Load those segments that the Route Formula <> 0 (GM)
        For mLooper = 1 To 4
            Select Case mLooper
                Case 1 'GM Runs
                    txt_StatusBox.Text = "Reading GM Mileage data..."
                    Refresh()
                    Application.DoEvents()
                    For mLooper2 = 1 To 4
                        ' open the streamreader
                        mWorkStr = Work_Directory & "WB" &
                                         txt_Year.Text & mPeriodLength & "_" & mBPGM_FileTypes(mLooper2) & mOutFileNameSuffix

                        sSr = New StreamReader(mWorkStr)
                        While (sSr.Peek <> -1)
                            mWorkStr = sSr.ReadLine
                            ' Save the line to the ALL.Out file                            
                            sBPAllOutFile.WriteLine(mWorkStr)
                        End While

                        sSr.Close()
                    Next

                Case 2 'IM Runs
                    txt_StatusBox.Text = "Reading IM Mileage data..."
                    Refresh()
                    Application.DoEvents()
                    For mLooper2 = 1 To 4
                        ' open the streamreader
                        mWorkStr = Work_Directory & "WB" &
                            txt_Year.Text & mPeriodLength & "_" & mBPIM_FileTypes(mLooper2) & mOutFileNameSuffix

                        sSr = New StreamReader(mWorkStr)
                        While (sSr.Peek <> -1)
                            mWorkStr = sSr.ReadLine
                            ' Save the line to the ALL.Out file                            
                            sBPAllOutFile.WriteLine(mWorkStr)
                        End While

                        sSr.Close()

                    Next
                Case 3 'CB Runs
                    txt_StatusBox.Text = "Reading CB Mileage data..."
                    Refresh()
                    Application.DoEvents()
                    For mLooper2 = 1 To 4
                        ' open the streamreader
                        mWorkStr = Work_Directory & "\WB" &
                           txt_Year.Text & mPeriodLength & "_" & mBPCB_FileTypes(mLooper2) & mOutFileNameSuffix

                        sSr = New StreamReader(mWorkStr)
                        While (sSr.Peek <> -1)
                            mWorkStr = sSr.ReadLine
                            ' Save the line to the ALL.Out file                            
                            sBPAllOutFile.WriteLine(mWorkStr)
                        End While

                        sSr.Close()

                    Next
                Case 4 'AR Runs
                    txt_StatusBox.Text = "Reading AR Mileage data..."
                    Refresh()
                    Application.DoEvents()
                    For mLooper2 = 1 To 4
                        ' open the streamreader
                        mWorkStr = Work_Directory & "\WB" &
                            txt_Year.Text & mPeriodLength & "_" & mBPAR_FileTypes(mLooper2) & mOutFileNameSuffix

                        sSr = New StreamReader(mWorkStr)
                        While (sSr.Peek <> -1)
                            mWorkStr = sSr.ReadLine
                            ' Save the line to the ALL.Out file                            
                            sBPAllOutFile.WriteLine(mWorkStr)
                        End While

                        sSr.Close()
                    Next
            End Select
        Next

        ' Add the All OUT Data to the SQL table.
        sBPAllOutFile.Flush()
        sBPAllOutFile.Close()

        mBatchPro_Ends = Now()

        OutString = New StringBuilder
        OutString.Append("BatchPro runs complete at " & mBatchPro_Ends.ToString("G") & " ")
        sbLogFile.WriteLine(OutString.ToString)
        OutString = New StringBuilder
        OutString.Append(Return_Elapsed_Time(mBatchPro_Starts, mBatchPro_Ends))
        sbLogFile.WriteLine(OutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

Skip_BatchPro_Processing:

        mSegments_Update_Starts = Now()

        txt_StatusBox.Text = "Updating Segments with Mileage data..."
        Refresh()
        Application.DoEvents()

        ' open the streamreader
        sSr = New StreamReader(txt_Work_Directory.Text & "\" & mBPAllOutFilename)

        While (sSr.Peek() <> -1)
            mWorkStr = sSr.ReadLine

            ' Parse the data values from the stream
            mSerial_No = Mid(mWorkStr, 1, 6)
            mSeg_No = Val(Mid(mWorkStr, 8, 1))
            mTotal_Segs = Val(Mid(mWorkStr, 10, 1))
            mSeg_Type = Mid(mWorkStr, 12, 2)
            mRoute_Formula = Val(Mid(mWorkStr, 15, 1))
            mRR_Num = Val(Mid(mWorkStr, 17, 3))
            mRR_Alpha = Trim(Mid(mWorkStr, 21, 4))
            mFrom_GeoCodeType = Val(Mid(mWorkStr, 26, 1))
            mFrom_GeoCodeValue = Mid(mWorkStr, 28, 5)
            mTo_GeoCodeType = Val(Mid(mWorkStr, 34, 1))
            mTo_GeoCodeValue = Mid(mWorkStr, 36, 5)
            mRR_Miles = Val(Mid(mWorkStr, 41, 9))
            mErrorMsg = Trim(Mid(mWorkStr, 51))

            'If mRR_Dist > 0 Then
            '    mRR_Dist = Val(Mid(mWorkStr, 41, 10)) * 10
            If mRR_Miles > 0 Then
                mRR_Dist = 10 * mRR_Miles
            ElseIf Len(Trim(mErrorMsg)) = 0 Then
                mRR_Dist = 10
            Else
                mRR_Dist = -1
            End If

            If chk_Skip_Masked_Data_Load.Checked = False Then
                ' If the Masked data load was skipped, then this record already exists in the
                ' All Miled table.

                ' Save the data to the Gbl_Interim_BatchPro_All_Miled SQL table
                Insert_All_Miled_Record(Gbl_Interim_Waybills_Database_Name,
                                    Gbl_Interim_BatchPro_All_Miled,
                                    mSerial_No,
                                    mSeg_No.ToString,
                                    mTotal_Segs.ToString,
                                    mSeg_Type.ToString,
                                    mRoute_Formula.ToString,
                                    mRR_Num.ToString,
                                    mRR_Alpha.ToString,
                                    mFrom_GeoCodeType.ToString,
                                    mFrom_GeoCodeValue.ToString,
                                    mTo_GeoCodeType.ToString,
                                    mTo_GeoCodeValue.ToString,
                                    mRR_Dist.ToString,
                                    mErrorMsg.ToString)
            End If
        End While

Skip_Segments_Data_Load:

        txt_StatusBox.Text = "Updating Masked records with Mileage data..."
        Refresh()
        Application.DoEvents()

        ' Now we will run the data against the segments and masked data
        OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)

        mSQLCommandString = New StringBuilder
        mSQLCommandString.Append("SELECT " & Gbl_Interim_Segments & ".Serial_No, ")
        mSQLCommandString.Append(Gbl_Interim_Segments & ".Seg_no, ")
        mSQLCommandString.Append(Gbl_Interim_Segments & ".Total_Segs, ")
        mSQLCommandString.Append(Gbl_Interim_Segments & ".RR_Dist, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".Service_Type ")
        mSQLCommandString.Append("FROM " & Gbl_Interim_Segments & " INNER JOIN ")
        mSQLCommandString.Append(Gbl_Interim_Masked & " ON ")
        mSQLCommandString.Append(Gbl_Interim_Segments & ".Serial_No = " & Gbl_Interim_Masked & ".Serial_No ")
        mSQLCommandString.Append("ORDER BY Serial_no, Seg_no ASC")

        mSegmentsTable = New DataTable

        Using daAdapter As New SqlDataAdapter(mSQLCommandString.ToString, gbl_SQLConnection)
            daAdapter.Fill(mSegmentsTable)
        End Using

        mORR_Dist = 0
        mJRR1_Dist = 0
        mJRR2_Dist = 0
        mJRR3_Dist = 0
        mJRR4_Dist = 0
        mJRR5_Dist = 0
        mJRR6_Dist = 0
        mTRR_Dist = 0
        mTotal_Dist = 0
        mBadFlag = False

        mMaxRecs = mSegmentsTable.Rows.Count
        mThisRec = 0

        For mLooper = 0 To mSegmentsTable.Rows.Count - 1

            mThisRec = mThisRec + 1

            ' Let the user know what is going on
            If mThisRec Mod 100 = 0 Then
                txt_StatusBox.Text = "Working - Updating Segment " & mThisRec.ToString & " of " & mMaxRecs.ToString & " records. "
                Refresh()
                Application.DoEvents()
            End If

            mService_Type = mSegmentsTable.Rows(mLooper)("Service_Type")
            mRoute_Formula = mService_Type

            If (mRoute_Formula = 1) Then
                mRoute_Formula = 0
            End If

            mSerial_No = mSegmentsTable(mLooper)("Serial_No")
            mSeg_No = mSegmentsTable(mLooper)("Seg_No")

            mSQLCommandString = New StringBuilder
            mSQLCommandString.Append("SELECT * FROM ")
            mSQLCommandString.Append(Gbl_Interim_BatchPro_All_Miled)
            mSQLCommandString.Append(" WHERE Serial_No = ")
            mSQLCommandString.Append("'" & mSerial_No & "'")
            mSQLCommandString.Append(" AND Seg_No = ")
            'mSQLCommandString.Append(mSegmentsTable(mLooper)("Seg_No").ToString)   ' MRS 2/25/2022
            mSQLCommandString.Append(mSeg_No.ToString)
            mSQLCommandString.Append(" AND Route_Formula = " & mRoute_Formula.ToString)

            msqlCommand = New SqlCommand
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.CommandText = mSQLCommandString.ToString

            mAllDateMiledTable = New DataTable

            Using daAdapter As New SqlDataAdapter(mSQLCommandString.ToString, gbl_SQLConnection)
                daAdapter.Fill(mAllDateMiledTable)
            End Using

            If mAllDateMiledTable.Rows.Count > 0 Then
                mRR_Dist = mAllDateMiledTable(0)("RR_Dist")

                If mRR_Dist = 0 Then
                    mRR_Dist = -1
                End If

            End If

            If mRR_Dist < 0 And mRoute_Formula <> 0 Then
                ' The distance is invalid for this non-zero Route Formula, lookup the distance again using a RouteFormula=0
                mSQLCommandString = New StringBuilder
                mSQLCommandString.Append("SELECT * FROM ")
                mSQLCommandString.Append(Gbl_Interim_BatchPro_All_Miled)
                mSQLCommandString.Append(" WHERE Serial_No = ")
                mSQLCommandString.Append("'" & mSerial_No & "'")
                mSQLCommandString.Append(" AND Seg_No = ")
                'mSQLCommandString.Append(mSegmentsTable(mLooper)("Seg_No").ToString)   ' MRS 2/25/2022
                mSQLCommandString.Append(mSeg_No.ToString)
                mSQLCommandString.Append(" AND Route_Formula = 0")

                msqlCommand = New SqlCommand
                msqlCommand.CommandType = CommandType.Text
                msqlCommand.Connection = gbl_SQLConnection
                msqlCommand.CommandText = mSQLCommandString.ToString

                mAllDateMiledTable = New DataTable

                Using daAdapter As New SqlDataAdapter(mSQLCommandString.ToString, gbl_SQLConnection)
                    daAdapter.Fill(mAllDateMiledTable)
                End Using

                If mAllDateMiledTable.Rows.Count = 0 Then
                    mRR_Dist = -1
                Else
                    mRR_Dist = mAllDateMiledTable(0)("RR_Dist")
                End If
            End If

            mSerial_No = mAllDateMiledTable.Rows(0)("Serial_No")
            mSeg_No = mAllDateMiledTable.Rows(0)("Seg_No")
            mTotal_Segs = mAllDateMiledTable.Rows(0)("Total_Segs")

            Update_Segment_Field(Gbl_Interim_Waybills_Database_Name,
                                     Gbl_Interim_Segments,
                                     mSerial_No,
                                     mSeg_No.ToString,
                                     "RR_Dist",
                                     mRR_Dist.ToString)
            mFieldName = ""

            If (mSeg_No = 1) And (mSeg_No = mTotal_Segs) Then
                mORR_Dist = mRR_Dist
                mFieldName = "ORR_Dist"
                Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name,
                            Gbl_Interim_Masked,
                            "Serial_No",
                            mSerial_No,
                            mFieldName.ToString,
                            mRR_Dist.ToString)
                mTRR_Dist = mRR_Dist
                mFieldName = "TRR_Dist"
                Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name,
                            Gbl_Interim_Masked,
                            "Serial_No",
                            mSerial_No,
                            mFieldName.ToString,
                            mRR_Dist.ToString)

            ElseIf (mSeg_No = 1) And (mSeg_No < mTotal_Segs) Then
                mORR_Dist = mRR_Dist
                mFieldName = "ORR_Dist"

                Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name,
                             Gbl_Interim_Masked,
                             "Serial_No",
                             mSerial_No,
                             mFieldName.ToString,
                             mRR_Dist.ToString)

            ElseIf (mSeg_No > 1) And (mSeg_No < mTotal_Segs) Then
                If (mSeg_No = 2) Then
                    mJRR1_Dist = mRR_Dist
                    mFieldName = "JRR1_Dist"
                ElseIf (mSeg_No = 3) Then
                    mJRR2_Dist = mRR_Dist
                    mFieldName = "JRR2_Dist"
                ElseIf (mSeg_No = 4) Then
                    mJRR3_Dist = mRR_Dist
                    mFieldName = "JRR3_Dist"
                ElseIf (mSeg_No = 5) Then
                    mJRR4_Dist = mRR_Dist
                    mFieldName = "JRR4_Dist"
                ElseIf (mSeg_No = 6) Then
                    mJRR5_Dist = mRR_Dist
                    mFieldName = "JRR5_Dist"
                ElseIf (mSeg_No = 7) Then
                    mJRR6_Dist = mRR_Dist
                    mFieldName = "JRR6_Dist"
                End If

                Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name,
                             Gbl_Interim_Masked,
                             "Serial_No",
                             mSerial_No,
                             mFieldName.ToString,
                             mRR_Dist.ToString)

            ElseIf (mSeg_No > 1) And (mSeg_No = mTotal_Segs) Then
                mTRR_Dist = mRR_Dist
                mFieldName = "TRR_Dist"

                Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name,
                             Gbl_Interim_Masked,
                             "Serial_No",
                             mSerial_No,
                             mFieldName.ToString,
                             mRR_Dist.ToString)
            End If
        Next

        ' Now we will run the data against the segments and masked data
        OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)

        mSQLCommandString = New StringBuilder
        With mSQLCommandString
            .Append("SELECT Serial_No, ")
            .Append("ORR_Dist, ")
            .Append("JRR1_Dist, ")
            .Append("JRR2_Dist, ")
            .Append("JRR3_Dist, ")
            .Append("JRR4_Dist, ")
            .Append("JRR5_Dist, ")
            .Append("JRR6_Dist, ")
            .Append("TRR_Dist, ")
            .Append("Total_Dist, ")
            .Append("JF ")
            .Append("FROM " & Gbl_Interim_Masked)
        End With

        mMaskedTable = New DataTable

        Using daAdapter As New SqlDataAdapter(mSQLCommandString.ToString, gbl_SQLConnection)
            daAdapter.Fill(mMaskedTable)
        End Using

        mMaxRecs = mMaskedTable.Rows.Count
        mThisRec = 0

        For mLooper = 0 To mMaskedTable.Rows.Count - 1

            mThisRec = mThisRec + 1

            ' Let the user know what is going on
            If mThisRec Mod 100 = 0 Then
                txt_StatusBox.Text = "Working - Updating Total Distances " & mThisRec.ToString & " of " & mMaxRecs.ToString & " records. "
                Refresh()
                Application.DoEvents()
            End If

            With mMaskedTable
                mORR_Dist = .Rows(mLooper)("ORR_Dist")
                mJRR1_Dist = .Rows(mLooper)("JRR1_Dist")
                mJRR2_Dist = .Rows(mLooper)("JRR2_Dist")
                mJRR3_Dist = .Rows(mLooper)("JRR3_Dist")
                mJRR4_Dist = .Rows(mLooper)("JRR4_Dist")
                mJRR5_Dist = .Rows(mLooper)("JRR5_Dist")
                mJRR6_Dist = .Rows(mLooper)("JRR6_Dist")
                mTRR_Dist = .Rows(mLooper)("TRR_Dist")
            End With

            If (mORR_Dist = -1 Or
                    mJRR1_Dist = -1 Or
                    mJRR2_Dist = -1 Or
                    mJRR3_Dist = -1 Or
                    mJRR4_Dist = -1 Or
                    mJRR5_Dist = -1 Or
                    mJRR6_Dist = -1 Or
                    mTRR_Dist = -1) Then
                mTotal_Dist = -1
            Else

                mTotal_Dist = mORR_Dist + mJRR1_Dist + mJRR2_Dist + mJRR3_Dist + mJRR4_Dist + mJRR5_Dist + mJRR6_Dist
                If (mMaskedTable.Rows(mLooper)("JF") > 0) Then
                    mTotal_Dist = mTotal_Dist + mTRR_Dist
                End If

            End If

            Update_Numeric_Field(Gbl_Interim_Waybills_Database_Name,
                                Gbl_Interim_Masked,
                                "Serial_No",
                                mMaskedTable.Rows(mLooper)("Serial_No").ToString,
                                "Total_Dist",
                                mTotal_Dist.ToString)
        Next

        mSegments_Update_Ends = Now()

        OutString = New StringBuilder
        OutString.Append("Interim Mileage data update complete at " & mSegments_Update_Ends.ToString("G") & " ")
        sbLogFile.WriteLine(OutString.ToString)
        OutString = New StringBuilder
        OutString.Append(Return_Elapsed_Time(mSegments_Update_Starts, mSegments_Update_Ends))
        sbLogFile.WriteLine(OutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

UpdateSTCCs:

        mSTCC_and_BEA_Updates_Start = Now()

        ' Update the STCC fields (STCC & STCC_W49) from the interim table.
        txt_StatusBox.Text = "Updating STCC/BEA information..."
        Refresh()
        Application.DoEvents()

        ' Now we will run the data against the segments and masked data
        OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)

        mSQLCommandString = New StringBuilder
        mSQLCommandString.Append("SELECT Serial_No, STCC_W49 FROM " & Gbl_Interim_Raw)

        mSTCCTable = New DataTable

        Using daAdapter As New SqlDataAdapter(mSQLCommandString.ToString, gbl_SQLConnection)
            daAdapter.Fill(mSTCCTable)
        End Using

        mThisRec = 1

        ' Looping thru the rows in the Interim table
        For mLooper = 0 To mSTCCTable.Rows.Count - 1

            ' Let the user know what is going on
            If mThisRec Mod 100 = 0 Then
                txt_StatusBox.Text = "Working - Updated STCC Codes " & mThisRec.ToString & " of " & mSTCCTable.Rows.Count.ToString & " records. "
                Refresh()
                Application.DoEvents()
            End If

            ' Update the STCC translation.  If the STCC_W49 has no translation, the sub will return the STCC_W49 value
            Update_Alpha_Field(Gbl_Interim_Waybills_Database_Name,
                               Gbl_Interim_Masked,
                               "Serial_No",
                               mSTCCTable.Rows(mLooper)("Serial_No"),
                               "STCC",
                                Get_STCC_W49_Translation(mSTCCTable.Rows(mLooper)("STCC_W49")))
            mThisRec = mThisRec + 1
        Next

UpdateBEAs:

        'mSTCC_and_BEA_Updates_Start = Now() 

        ' Update the BEA fields (O_BEA & T_BEA) from the interim table.
        txt_StatusBox.Text = "Updating BEA information..."
        Refresh()
        Application.DoEvents()

        'Get the SPLC to BEA table names
        mSPLC6toBEA_Table_Name = Get_Table_Name_From_SQL("1", "SPLC6toBEA")
        mSPLC4toBEA_Table_Name = Get_Table_Name_From_SQL("1", "SPLC4toBEA")

        mWorkTable = New DataTable

        ' Load the SPLC6 to BEA translation data
        OpenSQLConnection(Gbl_Waybill_Database_Name)

        mSQLCommandString = New StringBuilder
        mSQLCommandString.Append("SELECT * FROM " & mSPLC6toBEA_Table_Name)

        Using daAdapter As New SqlDataAdapter(mSQLCommandString.ToString, gbl_SQLConnection)
            daAdapter.Fill(mWorkTable)
        End Using

        SPLC6 = Nothing
        SPLC6_BEA = Nothing

        For mLooper = 1 To mWorkTable.Rows.Count - 1
            ReDim Preserve SPLC6(mLooper)
            ReDim Preserve SPLC6_BEA(mLooper)

            SPLC6(mLooper) = mWorkTable.Rows(mLooper)("SPLC6")
            SPLC6_BEA(mLooper) = mWorkTable.Rows(mLooper)("BEA")
        Next

        mWorkTable = New DataTable

        SPLC4 = Nothing
        SPLC4_BEA = Nothing

        mSQLCommandString = New StringBuilder
        mSQLCommandString.Append("SELECT * FROM " & mSPLC4toBEA_Table_Name)

        Using daAdapter As New SqlDataAdapter(mSQLCommandString.ToString, gbl_SQLConnection)
            daAdapter.Fill(mWorkTable)
        End Using

        For mLooper = 1 To mWorkTable.Rows.Count - 1
            ReDim Preserve SPLC4(mLooper)
            ReDim Preserve SPLC4_BEA(mLooper)

            SPLC4(mLooper) = mWorkTable.Rows(mLooper)("SPLC4")
            SPLC4_BEA(mLooper) = mWorkTable.Rows(mLooper)("BEA")
        Next

        ' Now we will run the data against the segments and masked data
        OpenSQLConnection(Gbl_Interim_Waybills_Database_Name)

        mSQLCommandString = New StringBuilder
        mSQLCommandString.Append("SELECT Serial_No, O_SPLC, T_SPLC FROM " & Gbl_Interim_Raw)

        mInterim_Raw_Table = New DataTable

        Using daAdapter As New SqlDataAdapter(mSQLCommandString.ToString, gbl_SQLConnection)
            daAdapter.Fill(mInterim_Raw_Table)
        End Using

        mThisRec = 1

        ' Looping thru the rows in the Interim table
        For mLooper = 0 To mInterim_Raw_Table.Rows.Count - 1

            ' Let the user know what is going on
            If mThisRec Mod 100 = 0 Then
                txt_StatusBox.Text = "Working - Updated BEA Codes " & mThisRec.ToString & " of " & mInterim_Raw_Table.Rows.Count.ToString & " records. "
                Refresh()
                Application.DoEvents()
            End If

            mSerial_No = mInterim_Raw_Table.Rows(mLooper)("Serial_No")

            ' Try to match the SPLC6 O_BEA translation.  If not found, mArrayLoc will equal 0
            mArrayLoc = SPLC6.ToList().IndexOf(mInterim_Raw_Table.Rows(mLooper)("O_SPLC"))

            If mArrayLoc > 0 Then
                'We have a match.  Update masked record
                mWorkStr = SPLC6_BEA(mArrayLoc)

            Else

                ' look for match with SPLC4
                mArrayLoc = SPLC4.ToList().IndexOf(Mid(mInterim_Raw_Table.Rows(mLooper)("O_SPLC"), 1, 4))

                If mArrayLoc > 0 Then
                    'We have a match
                    mWorkStr = SPLC4_BEA(mArrayLoc)
                Else
                    ' No match found.  BEA will be set to zero.
                    mWorkStr = "0"
                End If

            End If

            ' Execute the update
            mSQLCommandString = New StringBuilder

            mSQLCommandString.Append("UPDATE " & Gbl_Interim_Masked & " ")
            mSQLCommandString.Append("SET O_BEA " & " = '" & mWorkStr & "' WHERE Serial_No = '" & mSerial_No.ToString & "'")

            msqlCommand = New SqlCommand
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = mSQLCommandString.ToString
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.ExecuteNonQuery()

            ' Try to match the SPLC4 T_BEA translation.  If not found, mArrayLoc will equal 0
            mArrayLoc = SPLC6.ToList().IndexOf(mInterim_Raw_Table.Rows(mLooper)("T_SPLC"))

            If mArrayLoc > 0 Then
                'We have a match.  Update masked record
                mWorkStr = SPLC6_BEA(mArrayLoc)

            Else

                ' look for match with SPLC4
                mArrayLoc = SPLC4.ToList().IndexOf(Mid(mInterim_Raw_Table.Rows(mLooper)("T_SPLC"), 1, 4))

                If mArrayLoc > 0 Then
                    'We have a match
                    mWorkStr = SPLC4_BEA(mArrayLoc)
                Else
                    ' No match found.  BEA will be set to zero.
                    mWorkStr = "0"
                End If

            End If

            ' Execute the update
            mSQLCommandString = New StringBuilder

            mSQLCommandString.Append("UPDATE " & Gbl_Interim_Masked & " ")
            mSQLCommandString.Append("SET T_BEA " & " = '" & mWorkStr & "' WHERE Serial_No = '" & mSerial_No.ToString & "'")

            msqlCommand = New SqlCommand
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = mSQLCommandString.ToString
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.ExecuteNonQuery()

            mThisRec = mThisRec + 1
        Next

        mSTCC_and_BEA_Updates_Ends = Now()

        OutString = New StringBuilder
        OutString.Append("STCC/BEA data update complete at " & mSTCC_and_BEA_Updates_Ends.ToString("G") & " ")
        sbLogFile.WriteLine(OutString.ToString)
        OutString = New StringBuilder
        OutString.Append(Return_Elapsed_Time(mSTCC_and_BEA_Updates_Start, mSTCC_and_BEA_Updates_Ends))
        sbLogFile.WriteLine(OutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

Update_Interim_Unmasked:

        mUnmasked_Update_Starts = Now()

        ' Mask the U_rev for those roads that do not mask their own
        mSQLCommandString = New StringBuilder
        mSQLCommandString.Append("SELECT " & Gbl_Interim_Masked & ".Serial_No, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".Acct_Period, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".U_Rev, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".Report_RR, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".STCC_W49, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".WB_Date, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".Exp_Factor_Th, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".T_FSAC, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".WB_Num, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".U_Car_Num, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".U_TC_Num, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".U_Car_Init, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".Rate_Flg FROM ")
        mSQLCommandString.Append(Gbl_Interim_Masked & " INNER JOIN " & Gbl_Interim_Raw & " ON ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".Serial_No = " & Gbl_Interim_Raw & ".Serial_No")

        mMaskedTable = New DataTable

        Using daAdapter As New SqlDataAdapter(mSQLCommandString.ToString, gbl_SQLConnection)
            daAdapter.Fill(mMaskedTable)
        End Using

        mThisRec = 1

        For mLooper = 0 To mMaskedTable.Rows.Count - 1
            ' Let the user know what is going on
            If mThisRec Mod 100 = 0 Then
                txt_StatusBox.Text = "Working - Updated Unmasked Record " & mThisRec.ToString & " of " & mMaskedTable.Rows.Count.ToString & " records. "
                Refresh()
                Application.DoEvents()
            End If

            mSerial_No = mMaskedTable.Rows(mLooper)("Serial_No")
            mWBDate = mMaskedTable.Rows(mLooper)("wb_Date")
            mSTCC_W49 = mMaskedTable.Rows(mLooper)("STCC_W49")
            mAcctMonth = Mid(mMaskedTable.Rows(mLooper)("acct_period"), 1, 2)
            mAcctYear = CInt(Mid(mMaskedTable.Rows(mLooper)("acct_period"), 3))
            mU_Rev = mMaskedTable.Rows(mLooper)("U_Rev")
            mT_FSAC = mMaskedTable.Rows(mLooper)("T_FSAC")
            mWbNum = mMaskedTable.Rows(mLooper)("wb_Num")
            mUCarNum = mMaskedTable.Rows(mLooper)("U_Car_Num")
            mu_TC_Num = mMaskedTable.Rows(mLooper)("u_tc_num")
            mUCarInit = Mid(mMaskedTable.Rows(mLooper)("u_car_init"), 1, 1)

            'Default values in case rate_flg <> 1
            mU_Rev = mMaskedTable.Rows(mLooper)("U_Rev")
            mU_Rev_Unmasked = mU_Rev

            Select Case mMaskedTable.Rows(mLooper)("Report_RR")
                ' Roads that mask their own revenues
                Case 105  'CP

                    If mMaskedTable.Rows(mLooper)("Rate_Flg") = 1 Then
                        mU_Rev = mMaskedTable.Rows(mLooper)("U_Rev")
                        ' Unmask revenues and store in unmasked table
                        mU_Rev_Unmasked = UnmaskCPValue(mSerial_No, mSTCC_W49, mU_Rev)
                    End If
                Case 555  'NS
                    If mMaskedTable.Rows(mLooper)("Rate_Flg") = 1 Then
                        mU_Rev = mMaskedTable.Rows(mLooper)("U_Rev")
                        ' Unmask revenues and store in unmasked table
                        mU_Rev_Unmasked = UnmaskNSValue(mSerial_No, mWBDate, mAcctYear, mAcctMonth, mT_FSAC, mWbNum, mUCarNum, mu_TC_Num, mU_Rev)
                    End If
                Case 712 'CSXT
                    If mMaskedTable.Rows(mLooper)("Rate_Flg") = 1 Then
                        mU_Rev = mMaskedTable.Rows(mLooper)("U_Rev")
                        ' Unmask revenues and store in unmasked table
                        mU_Rev_Unmasked = UnmaskCSXValue(mSerial_No, mSTCC_W49, mWbNum, mAcctMonth, mAcctYear, mU_Rev)
                    End If
                Case 777  'BNSF
                    If mMaskedTable.Rows(mLooper)("Rate_Flg") = 1 Then
                        mU_Rev = mMaskedTable.Rows(mLooper)("U_Rev")
                        ' Unmask revenues and store in unmasked table
                        mU_Rev_Unmasked = UnmaskBNSFValue(mSerial_No, mUCarInit, mAcctMonth, mU_Rev)
                    End If
                Case 802  'UP
                    If mMaskedTable.Rows(mLooper)("Rate_Flg") = 1 Then
                        mU_Rev = mMaskedTable.Rows(mLooper)("U_Rev")
                        ' Unmask revenues and store in unmasked table
                        mU_Rev_Unmasked = UnmaskUPValue(mSerial_No, mSTCC_W49, mAcctYear, mWbNum, mU_Rev)
                    End If
                Case Else  'Mask the revenues for those roads that don't mask their own revenues
                    If mMaskedTable.Rows(mLooper)("Rate_Flg") = 1 Then
                        ' Store the U_Rev to unmasked table
                        mU_Rev_Unmasked = mU_Rev
                        'mask the u_rev field
                        mU_Rev = MaskGenericValue(mSTCC_W49, mWBDate, mAcctYear, mU_Rev)
                    Else
                        mU_Rev = mMaskedTable.Rows(mLooper)("U_Rev")
                        mU_Rev_Unmasked = mU_Rev
                    End If
            End Select

            'Calculate the Total_Rev and total_unmasked_Rev by multiplying U_Rev amd U_Unmasked_Rev by the Exp_Factor_Th.
            mTotal_Rev = mU_Rev * mMaskedTable.Rows(mLooper)("Exp_Factor_Th")
            mTotal_Unmasked_Rev = mU_Rev_Unmasked * mMaskedTable.Rows(mLooper)("Exp_Factor_Th")

            'Update the U_Rev field in masked table
            mSQLCommandString = New StringBuilder

            mSQLCommandString.Append("UPDATE " & Gbl_Interim_Masked & " ")
            mSQLCommandString.Append("SET U_Rev " & " = " & mU_Rev & " WHERE Serial_No = '" & mSerial_No.ToString & "'")

            msqlCommand = New SqlCommand
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = mSQLCommandString.ToString
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.ExecuteNonQuery()

            'Update the U_Rev_Unmasked field
            mSQLCommandString = New StringBuilder

            mSQLCommandString.Append("UPDATE " & Gbl_Interim_Unmasked_Rev & " ")
            mSQLCommandString.Append("SET U_Rev_Unmasked " & " = " & mU_Rev_Unmasked & " WHERE Unmasked_Serial_No = '" & mSerial_No.ToString & "'")

            msqlCommand = New SqlCommand
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = mSQLCommandString.ToString
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.ExecuteNonQuery()

            'Update the Total_Rev field
            mSQLCommandString = New StringBuilder

            mSQLCommandString.Append("UPDATE " & Gbl_Interim_Masked & " ")
            mSQLCommandString.Append("SET Total_Rev " & " = " & mTotal_Rev & " WHERE Serial_No = '" & mSerial_No.ToString & "'")

            msqlCommand = New SqlCommand
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = mSQLCommandString.ToString
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.ExecuteNonQuery()

            mSQLCommandString = New StringBuilder

            mSQLCommandString.Append("UPDATE " & Gbl_Interim_Unmasked_Rev & " ")
            mSQLCommandString.Append("SET Total_Unmasked_Rev " & " = " & mTotal_Unmasked_Rev & " WHERE Unmasked_Serial_No = '" & mSerial_No.ToString & "'")

            msqlCommand = New SqlCommand
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = mSQLCommandString.ToString
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.ExecuteNonQuery()

            mThisRec += 1

        Next

        mUnmasked_Update_Ends = Now()

        OutString = New StringBuilder
        OutString.Append("Unmasked Data Update Complete at " & mUnmasked_Update_Ends.ToString("G") & " ")
        sbLogFile.WriteLine(OutString.ToString)
        OutString = New StringBuilder
        OutString.Append(Return_Elapsed_Time(mUnmasked_Update_Starts, mUnmasked_Update_Ends))
        sbLogFile.WriteLine(OutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

Update_PUWS:

        mPUWS_Update_Starts = Now()

        'finally, we can create the PUWS table data
        'Get the revenues from SQL server for this serial number for anything other than UP & PAL
        mSQLCommandString = New StringBuilder
        mSQLCommandString.Append("SELECT " & Gbl_Interim_Masked & ".serial_no, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".total_rev, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".u_rev, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".report_rr, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".rate_flg, ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".exp_factor_th, ")
        mSQLCommandString.Append(Gbl_Interim_Unmasked_Rev & ".u_rev_unmasked ")
        mSQLCommandString.Append(" FROM " & Gbl_Interim_Masked & " INNER JOIN ")
        mSQLCommandString.Append(Gbl_Interim_Unmasked_Rev & " ON " & Gbl_Interim_Masked & ".Serial_No = ")
        mSQLCommandString.Append(Gbl_Interim_Unmasked_Rev & ".Unmasked_Serial_No WHERE ")
        mSQLCommandString.Append(Gbl_Interim_Masked & ".serial_no = " & Gbl_Interim_Unmasked_Rev & ".Unmasked_Serial_No")

        mWorkTable = New DataTable

        Using daAdapter As New SqlDataAdapter(mSQLCommandString.ToString, gbl_SQLConnection)
            daAdapter.Fill(mWorkTable)
        End Using

        mThisRec = 1
        For mLooper = 0 To mWorkTable.Rows.Count - 1

            If mThisRec Mod 100 = 0 Then
                txt_StatusBox.Text = "Creating PUWS record " & mThisRec.ToString & " of " & mWorkTable.Rows.Count.ToString
                Refresh()
                Application.DoEvents()
            End If

            mTotal_Rev = mWorkTable.Rows(mLooper)("total_rev")
            mU_Rev = mWorkTable.Rows(mLooper)("u_rev")
            mPUWS_Masking_Factor = 1
            mExpanded_Unmasked_Revenue = 0

            If mWorkTable.Rows(mLooper)("rate_flg") > 0 Then

                Randomize()

                Select Case mWorkTable.Rows(mLooper)("report_rr")
                    Case 555, 712, 777, 105
                        'N/A
                    Case 802, 907
                        ' UP & P&L (P&L added for 2013+)
                        Do Until mPUWS_Masking_Factor > 1.04 And mPUWS_Masking_Factor <= 1.55
                            mPUWS_Masking_Factor = 1 + Rnd()
                            mPUWS_Masking_Factor = System.Math.Round(mPUWS_Masking_Factor, 8, MidpointRounding.AwayFromZero)
                        Loop

                        ' For these roads we mask the unmasked revenues
                        If mU_Rev > 0 And mTotal_Rev > 0 Then
                            ' Calculate the U_Rev * exp_factor_th.  It is now the expanded revenue
                            mExpanded_Unmasked_Revenue = mWorkTable.Rows(mLooper)("u_rev_unmasked") * mWorkTable.Rows(mLooper)("exp_factor_th")
                            ' Calculate the PWS total revenue by multiplying the expanded unmasked revenue by the masking factor
                            mTotal_Rev = System.Math.Round(mExpanded_Unmasked_Revenue * mPUWS_Masking_Factor, 0, MidpointRounding.AwayFromZero)
                        End If
                End Select
            End If

            If Is_Record_In_Table(Gbl_Interim_Waybills_Database_Name,
                                  Gbl_interim_PUWS,
                                  "PUWS_Serial_No",
                                  mWorkTable.Rows(mLooper)("Serial_No")) = False Then
                ' Now we need to record/write the masking factor to its table
                mSQLCommandString = New StringBuilder
                mSQLCommandString.Append("INSERT INTO " & Gbl_interim_PUWS & " (PUWS_Serial_No, PUWS_Total_Rev, PUWS_Masking_Factor) VALUES ('")
                mSQLCommandString.Append(mWorkTable.Rows(mLooper)("Serial_No").ToString & "', ")
                mSQLCommandString.Append(mTotal_Rev.ToString & ", ")
                mSQLCommandString.Append(Format(mPUWS_Masking_Factor, "0.00000000") & ")")
            Else
                ' We need to update it
                mSQLCommandString = New StringBuilder
                mSQLCommandString.Append("UPDATE " & Gbl_interim_PUWS & " ")
                mSQLCommandString.Append("SET PUWS_Total_Rev = " & mTotal_Rev & ", ")
                mSQLCommandString.Append("PUWS_Masking_Factor = " & Format(mPUWS_Masking_Factor, "0.00000000") & " ")
                mSQLCommandString.Append("WHERE PUWS_Serial_No = '" & mSerial_No.ToString & "'")
            End If

            msqlCommand = New SqlCommand
            msqlCommand.CommandType = CommandType.Text
            msqlCommand.CommandText = mSQLCommandString.ToString
            msqlCommand.Connection = gbl_SQLConnection
            msqlCommand.ExecuteNonQuery()

            mThisRec += 1

        Next

        mPUWS_Update_Ends = Now()

        OutString = New StringBuilder
        OutString.Append("PUWS Update Complete at " & mPUWS_Update_Ends.ToString("G") & " ")
        sbLogFile.WriteLine(OutString.ToString)
        OutString = New StringBuilder
        OutString.Append(Return_Elapsed_Time(mPUWS_Update_Starts, mPUWS_Update_Ends))
        sbLogFile.WriteLine(OutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

EndIt:

        sbLogFile.Flush()

        mRun_Stops = Now()

        OutString = New StringBuilder
        OutString.Append("*** Run Completed at " & mRun_Stops.ToString("G") & " ")
        sbLogFile.WriteLine(OutString.ToString)
        OutString = New StringBuilder
        OutString.Append(Return_Elapsed_Time(mRun_Start, mRun_Stops))
        sbLogFile.WriteLine(OutString.ToString)
        sbLogFile.Flush()

        ' Finally, we're done
        txt_StatusBox.Text = "Done!"

    End Sub

    Private Sub rdo_Annual_CheckedChanged(sender As Object, e As EventArgs) Handles rdo_Annual.CheckedChanged

        rdo_1st_Quarter.Visible = False
        rdo_2nd_Quarter.Visible = False
        rdo_3rd_Quarter.Visible = False
        rdo_4th_Quarter.Visible = False
        gbx_Monthly.Visible = False
        gbx_Quarter.Visible = False

    End Sub

    Private Sub btn_Select_Work_Directory_Click(sender As Object, e As EventArgs) Handles btn_Select_Work_Directory.Click
        Dim folderDlg As New FolderBrowserDialog

        folderDlg.ShowNewFolderButton = True
        folderDlg.SelectedPath = "C:\"
        folderDlg.Description = "It is HIGHLY suggested that a folder on this machine's hard drive be used!"

        If (folderDlg.ShowDialog() = DialogResult.OK) Then

            txt_Work_Directory.Text = folderDlg.SelectedPath

            Dim root As Environment.SpecialFolder = folderDlg.RootFolder

        End If

    End Sub

    Private Sub chk_Skip_BatchPro_Processing_CheckedChanged(sender As Object, e As EventArgs) Handles chk_Skip_BatchPro_Processing.CheckedChanged
        If chk_Skip_BatchPro_Processing.Checked = True Then
            chk_Skip_Masked_Data_Load.Checked = True
            btn_Input_File_Entry.Enabled = False
        Else
            chk_Skip_Masked_Data_Load.Checked = False
            btn_Input_File_Entry.Enabled = True
        End If
    End Sub

    Private Sub chk_Skip_Segments_Data_Load_CheckedChanged(sender As Object, e As EventArgs) Handles chk_Skip_Segments_Data_Load.CheckedChanged
        If chk_Skip_Segments_Data_Load.Checked = True Then
            chk_Skip_Masked_Data_Load.Checked = True
            chk_Skip_BatchPro_Processing.Checked = True
            btn_Input_File_Entry.Enabled = False
        Else
            chk_Skip_Masked_Data_Load.Checked = False
            chk_Skip_BatchPro_Processing.Checked = False
            btn_Input_File_Entry.Enabled = True
        End If

    End Sub

    Private Sub chk_Skip_Masked_Data_Load_CheckedChanged(sender As Object, e As EventArgs) Handles chk_Skip_Masked_Data_Load.CheckedChanged
        If chk_Skip_Masked_Data_Load.Checked = True Then
            btn_Input_File_Entry.Enabled = False
            txt_Input_FilePath.Enabled = False
        Else
            btn_Input_File_Entry.Enabled = True
            txt_Input_FilePath.Enabled = True
        End If
    End Sub

End Class

