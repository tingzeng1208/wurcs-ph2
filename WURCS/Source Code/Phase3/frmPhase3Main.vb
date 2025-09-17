'**********************************************************************
' Title:        Phase 3 Costing
' Author:       Michael Sanders and Michael Boyles
' Purpose:      Produces the Balance Report to make data integrity checking easier for URCS.
' Revisions:    Initial Creation - Spring 2017
'               Added logic to utilize data choices for either "Legacy" or "Current" - Jun 14 2017 
'               Removed ADODB and replaced with SQL Dataset - Jun 20 2018
'               Added Makewhole adjustments calculations to spreadsheets and adjusted pgm logic - Summer 2018
'               Added option to cost select year using XML from another year - Fall 2018
' 
' This program is US Government Property - For Official Use Only
'**********************************************************************

Imports SpreadsheetGear
Imports System.Text
Imports System.Data.SqlClient

Public Class frmPhase3Main

    Private Const BadCarType = 1
    Private Const BadTonnage = 2
    Private Const BadRevenue = 3
    Private Const BadDistance = 4
    Private Const BadIntermodalSpec = 5
    Private Const NonUSSegment = 6
    Private Const NegativeCost = 7

    Private Const IntermodalMove = 1
    Private Const SingleMove = 2
    Private Const MultiMove = 3
    Private Const UnitMove = 4

#Region "Form events"

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        mDataTable = Get_URCS_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            cmb_URCS_Year.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
            cmb_Different_Year.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
        Next

        mDataTable = Nothing

        chk100Series.Checked = False
        chk200Series.Checked = False
        chk300Series.Checked = False
        chk400Series.Checked = False
        chk500Series.Checked = False
        chk600Series.Checked = False
        chkSaveResults.Checked = False
        chk_Save_CRPRESRecords.Checked = True
        chk_Cost_All_Segments.Checked = True
        chk_SaveToSQL.Checked = True

        chk_Cost_All_Segments.Visible = False

        txt_Target_Server_Name.Text = gbl_Server_Name

        Me.CenterToScreen()

    End Sub

    Private Sub btn_Return_To_MainMenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_Return_To_MainMenu.Click
        ' Open the Main Menu Form
        Dim frmNew As New frm_MainMenu()
        frmNew.Show()
        ' Close this Menu
        Me.Close()
    End Sub

    Private Sub btnSelectExcelFile_Click(sender As Object, e As EventArgs) Handles btnSelectP3File.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm|All Files|*.*"

        If fd.ShowDialog() = DialogResult.OK Then
            txtP3FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub btnSelectOutputFolder_Click(sender As Object, e As EventArgs) Handles btnSelectOutputFolder.Click
        Dim oFolderBrowserDialog As FolderBrowserDialog

        oFolderBrowserDialog = New FolderBrowserDialog
        oFolderBrowserDialog.ShowDialog()

        If oFolderBrowserDialog.SelectedPath.Length > 0 Then
            txtFolder.Text = oFolderBrowserDialog.SelectedPath
        End If
    End Sub

    Private Sub btnSelectMWFFile_Click(sender As Object, e As EventArgs) Handles btnSelectMWFFile.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm|All Files|*.*"

        If fd.ShowDialog() = DialogResult.OK Then
            txtMakeWholeFilePath.Text = fd.FileName
        End If
    End Sub

#End Region

#Region "Spreadsheet Data Handlers"

    Private Function Get100Series(ByVal mSheet As IWorksheet) As String
        ' This returns the K5:U5 values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 75 To 84        ' K thru T
            mOutString.Append(mSheet.Cells(Chr(mLooper) & "5").Value & ",")
        Next
        mOutString.Append(mSheet.Cells("U5").Value)
        Get100Series = mOutString.ToString

    End Function

    Private Function Get100SeriesHeaders(ByVal mSheet As IWorksheet) As String
        ' This returns the header values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 75 To 84        ' K thru T
            mOutString.Append(mSheet.Cells(Chr(mLooper) & "4").Value & ",")
        Next
        mOutString.Append(mSheet.Cells("U4").Value)
        Get100SeriesHeaders = mOutString.ToString

    End Function

    Private Function Get200Series(ByVal mSheet As IWorksheet) As String
        ' This returns the V5:DF5 values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 86 To 90          ' V thru Z
            mOutString.Append(mSheet.Cells(Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 90          ' AA thru AZ
            mOutString.Append(mSheet.Cells("A" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 90          ' BA thru BZ
            mOutString.Append(mSheet.Cells("B" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 90          ' CA thru CZ
            mOutString.Append(mSheet.Cells("C" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 70          ' DA thru DF
            mOutString.Append(mSheet.Cells("D" & Chr(mLooper) & "5").Value & ",")
        Next
        Get200Series = mOutString.ToString

    End Function

    Private Function Get200SeriesHeaders(ByVal mSheet As IWorksheet) As String
        ' This returns the Header values V4:DF4 values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 86 To 90          ' V thru Z
            mOutString.Append(mSheet.Cells(Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 90          ' AA thru AZ
            mOutString.Append(mSheet.Cells("A" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 90          ' BA thru BZ
            mOutString.Append(mSheet.Cells("B" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 90          ' CA thru CZ
            mOutString.Append(mSheet.Cells("C" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 69          ' DA thru DE
            mOutString.Append(mSheet.Cells("D" & Chr(mLooper) & "4").Value & ",")
        Next
        mOutString.Append(mSheet.Cells("DF4").Value)
        Get200SeriesHeaders = mOutString.ToString

    End Function

    Private Function Get300Series(ByVal mSheet As IWorksheet) As String
        ' This returns the DG5:EN5 values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 71 To 90          ' DG thru DZ
            mOutString.Append(mSheet.Cells("D" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 77          ' EA thru EM
            mOutString.Append(mSheet.Cells("E" & Chr(mLooper) & "5").Value & ",")
        Next
        mOutString.Append(mSheet.Cells("EN5").Value)
        Get300Series = mOutString.ToString

    End Function

    Private Function Get300SeriesHeaders(ByVal mSheet As IWorksheet) As String
        ' This returns the headers values DG5:EN5 values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 71 To 90          ' DG thru DZ
            mOutString.Append(mSheet.Cells("D" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 77          ' EA thru EM
            mOutString.Append(mSheet.Cells("E" & Chr(mLooper) & "4").Value & ",")
        Next
        mOutString.Append(mSheet.Cells("EN4").Value)
        Get300SeriesHeaders = mOutString.ToString

    End Function

    Private Function Get400Series(ByVal mSheet As IWorksheet) As String
        ' This returns the EO5:IM5 values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 79 To 90          ' EO thru EZ
            mOutString.Append(mSheet.Cells("E" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 90          ' FA thru FZ
            mOutString.Append(mSheet.Cells("F" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 90          ' GA thru GZ
            mOutString.Append(mSheet.Cells("G" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 90          ' HA thru HZ
            mOutString.Append(mSheet.Cells("H" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 76          ' IA thru IL
            mOutString.Append(mSheet.Cells("I" & Chr(mLooper) & "5").Value & ",")
        Next
        mOutString.Append(mSheet.Cells("IM5").Value)
        Get400Series = mOutString.ToString

    End Function

    Private Function Get400SeriesHeaders(ByVal mSheet As IWorksheet) As String
        ' This returns the header value EO5:IM5 values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 79 To 90          ' EO thru EZ
            mOutString.Append(mSheet.Cells("E" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 90          ' FA thru FZ
            mOutString.Append(mSheet.Cells("F" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 90          ' GA thru GZ
            mOutString.Append(mSheet.Cells("G" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 90          ' HA thru HZ
            mOutString.Append(mSheet.Cells("H" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 76          ' IA thru IL
            mOutString.Append(mSheet.Cells("I" & Chr(mLooper) & "4").Value & ",")
        Next
        mOutString.Append(mSheet.Cells("IM4").Value)
        Get400SeriesHeaders = mOutString.ToString

    End Function

    Private Function Get500Series(ByVal mSheet As IWorksheet) As String
        ' This returns the IN5:LR5 values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 78 To 90          ' IN thru IZ
            mOutString.Append(mSheet.Cells("I" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 90          ' JA thru JZ
            mOutString.Append(mSheet.Cells("J" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 90          ' KA thru KZ
            mOutString.Append(mSheet.Cells("K" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 81          ' LA thru LQ
            mOutString.Append(mSheet.Cells("L" & Chr(mLooper) & "5").Value & ",")
        Next
        mOutString.Append(mSheet.Cells("LR5").Value)
        Get500Series = mOutString.ToString

    End Function

    Private Function Get500SeriesHeaders(ByVal mSheet As IWorksheet) As String
        ' This returns the header value IN5:LR5 values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 78 To 90          ' IN thru IZ
            mOutString.Append(mSheet.Cells("I" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 90          ' JA thru JZ
            mOutString.Append(mSheet.Cells("J" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 90          ' KA thru KZ
            mOutString.Append(mSheet.Cells("K" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 81          ' LA thru LQ
            mOutString.Append(mSheet.Cells("L" & Chr(mLooper) & "4").Value & ",")
        Next
        mOutString.Append(mSheet.Cells("LR4").Value)
        Get500SeriesHeaders = mOutString.ToString

    End Function

    Private Function Get600Series(ByVal mSheet As IWorksheet) As String
        ' This returns the LS5:PQ5 values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 83 To 90          ' LS thru LZ
            mOutString.Append(mSheet.Cells("L" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 90          ' MA thru MZ
            mOutString.Append(mSheet.Cells("M" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 90          ' NA thru NZ
            mOutString.Append(mSheet.Cells("N" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 90          ' OA thru OZ
            mOutString.Append(mSheet.Cells("O" & Chr(mLooper) & "5").Value & ",")
        Next
        For mLooper = 65 To 80          ' PA thru PP
            mOutString.Append(mSheet.Cells("P" & Chr(mLooper) & "5").Value & ",")
        Next
        mOutString.Append(mSheet.Cells("PQ5").Value)
        Get600Series = mOutString.ToString

    End Function

    Private Function Get600SeriesHeaders(ByVal mSheet As IWorksheet) As String
        ' This returns the header value LS5:PQ5 values from the P3 BatchOutput sheet, comma delimited
        Dim mLooper As Integer
        Dim mOutString As StringBuilder

        mOutString = New StringBuilder
        For mLooper = 83 To 90          ' LS thru LZ
            mOutString.Append(mSheet.Cells("L" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 90          ' MA thru MZ
            mOutString.Append(mSheet.Cells("M" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 90          ' NA thru NZ
            mOutString.Append(mSheet.Cells("N" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 90          ' OA thru OZ
            mOutString.Append(mSheet.Cells("O" & Chr(mLooper) & "4").Value & ",")
        Next
        For mLooper = 65 To 80          ' PA thru PP
            mOutString.Append(mSheet.Cells("P" & Chr(mLooper) & "4").Value & ",")
        Next
        mOutString.Append(mSheet.Cells("PQ4").Value)
        Get600SeriesHeaders = mOutString.ToString

    End Function


#End Region

    Private Sub btnExecute_Click(sender As Object, e As EventArgs) Handles btnExecute.Click

        Dim mDataTable As DataTable, mMaskedTable As DataTable
        Dim mStrSQL As String
        Dim mLooper As Integer, mLooper2 As Integer
        Dim mCounter As Integer
        Dim mPass1Starts As Date
        Dim mPass2Starts As Date
        Dim mPass3Starts As Date
        Dim mEValuesUpdateStarts As Date
        Dim mUnitTrainDef As Integer = 0
        Dim mTOFCMove As Boolean = False
        Dim mVal As Object
        Dim m600Procd As Boolean = False
        Dim mProcess As Boolean = True
        Dim mGoodSegment As Boolean = True
        Dim mXMLSheetColumn As String = Nothing
        Dim mVLookup As String = Nothing
        Dim mThisCost As Integer = 0
        Dim mThisRoad As Integer
        Dim mColumn As String = Nothing             ' Added 6/12/2019 
        Dim mLookupOkay As Boolean = True           ' Added 6/12/2019
        Dim mTS As TimeSpan                         ' Added 12/27/2002
        Dim mWorkStr As String = ""                ' Added 12/27/2022

        ' Used when updating EValues table
        Dim mSQLcmd As SqlCommand

        'Used for testing purposes
        Dim mThreesOnly As Boolean = False
        Dim mTestOne As Boolean = False

        ' Used by SpreadsheetGear for the Phase III workbook
        Dim mP3WorkbookSet As IWorkbookSet
        Dim mP3Workbook As IWorkbook
        Dim mP3RailroadCostProgramSheet As IWorksheet
        Dim mP3BatchCostProgramSheet As IWorksheet
        Dim mP3BatchOutputSheet As IWorksheet
        Dim mP3DetailedParametersSheet As IWorksheet
        Dim mP3RailroadUnitCostXMLSheet As IWorksheet
        Dim mP3CellRange As IRange

        ' Used by SpreadsheetGear for the MakeWhole workbook
        Dim mMWFWorkbookSet As IWorkbookSet
        Dim mMWFWorkbook As IWorkbook
        Dim mMWFMovement As IWorksheet
        Dim mMWFCRPRESRecord As IWorksheet
        Dim mMWFCRPRES As IWorksheet
        Dim mMWFXMLValues As IWorksheet
        Dim mMWFCellRange As IRange

        ' Local Memvars for the waybill to Spreadsheet data
        Dim mSerial_No As String = ""
        Dim mSeg_No As Integer
        Dim mExp_Factor_Th As Integer
        Dim mResiduals(9, 18) As Long
        Dim mSerialChanged As Boolean = False

        ' Variables for output files/SQL tables
        Dim sb100Series_SysAvg As StreamWriter
        Dim sb100Series_EffAdj As StreamWriter
        Dim sb100Series_Costed As StreamWriter
        Dim sb200Series_SysAvg As StreamWriter
        Dim sb200Series_EffAdj As StreamWriter
        Dim sb200Series_Costed As StreamWriter
        Dim sb300Series_SysAvg As StreamWriter
        Dim sb300Series_EffAdj As StreamWriter
        Dim sb300Series_Costed As StreamWriter
        Dim sb400Series_SysAvg As StreamWriter
        Dim sb400Series_EffAdj As StreamWriter
        Dim sb400Series_Costed As StreamWriter
        Dim sb500Series_SysAvg As StreamWriter
        Dim sb500Series_EffAdj As StreamWriter
        Dim sb500Series_Costed As StreamWriter
        Dim sb600Series_SysAvg As StreamWriter
        Dim sb600Series_EffAdj As StreamWriter
        Dim sb600Series_Costed As StreamWriter
        Dim sbResults_SysAvg As StreamWriter
        Dim sbResults_EffAdj As StreamWriter
        Dim sbResults_Costed As StreamWriter
        Dim sbCRPRES As StreamWriter
        Dim sbCRPRESRecord As StreamWriter
        Dim sbLogFile As StreamWriter
        Dim mOutString As StringBuilder
        Dim mOutputMaskedTableName As String
        Dim mOutputSegmentsTableName As String
        Dim mLogFileName As String = ""

        ' Counters/memvars used for Log file
        Dim mNumOfWaybillRecords As Integer = 0
        Dim mNumOfCostedWaybills As Integer = 0
        Dim mNumOfCostedRecords As Integer = 0
        Dim mNumOfCostedSegments As Integer = 0
        Dim mNumOfMaskedSegments As Integer = 0
        Dim mNumOfRejectedSegments As Integer = 0
        Dim mNumOfBadLookups As Integer = 0                 'Added 6/12/2019
        Dim mCostedSegmentsByRailroad(9) As Integer
        Dim mRejectedSegments(7) As Integer
        Dim mRejectedWaybills(7) As Integer
        Dim mCostedSegments(4) As Integer
        Dim mCostedWaybills(4) As Integer
        Dim mPreviousSerialNo As String = ""
        Dim mTotalVariableCosts(9) As Double
        Dim mGrandTotal As Double = 0

        'Array for CRPRES data collection
        Dim mCRPRESArray(9, 19) As Double

        'Array for the Cost Breakdown data collection
        Dim mCostBreakdown(9, 4) As Double
        Dim mCostBreakdownCol3 As Double = 0
        Dim mCostBreakdownCol4 As Double = 0
        Dim mCostBreakdownCol1Total As Double = 0
        Dim mCostBreakdownCol2Total As Double = 0
        Dim mCostBreakdownCol3Total As Double = 0
        Dim mCostBreakdownCol4Total As Double = 0

        'Check to make sure that the form has all of the info we need.
        If txtP3FilePath.Text = "" Then
            MsgBox("No Excel file chosen!", vbOKOnly, "ERROR!")
            GoTo EndIt
        End If

        If txtFolder.Text = "" Then
            MsgBox("No output folder chosen!", vbOKOnly, "ERROR!")
            GoTo EndIt
        End If

        If File.Exists(txtFolder.Text & "\" & cmb_URCS_Year.Text & "*.CSV") Then
            If MsgBox("This will overwrite CSV files in your output directory!  Proceed?", vbYesNo, "WARNING") = vbNo Then
                GoTo EndIt
            End If
        End If

        If (Not rdo_Legacy.Checked) And (Not rdo_Current.Checked) Then
            MsgBox("No Costing Method chosen!", vbOKOnly, "ERROR!")
            GoTo EndIt
        End If

        mPass1Starts = Now

        txtStatus.Text = "Initializing Spreadsheets..."
        Refresh()
        Application.DoEvents()

        'Open the Phase III Excel workbook.
        mP3WorkbookSet = Factory.GetWorkbookSet(System.Globalization.CultureInfo.CurrentCulture)
        mP3WorkbookSet.Calculation = Calculation.Manual
        mP3Workbook = mP3WorkbookSet.Workbooks.Open(txtP3FilePath.Text)

        'Open the Make Whole Excel workbook.
        mMWFWorkbookSet = Factory.GetWorkbookSet(System.Globalization.CultureInfo.CurrentCulture)
        mMWFWorkbookSet.Calculation = Calculation.Manual
        mMWFWorkbook = mMWFWorkbookSet.Workbooks.Open(txtMakeWholeFilePath.Text)

        'Set up the P3 worksheets for use
        mP3RailroadCostProgramSheet = mP3Workbook.Worksheets("RailroadCostProgram")
        mP3BatchCostProgramSheet = mP3Workbook.Worksheets("BatchCostProgram")
        mP3BatchOutputSheet = mP3Workbook.Worksheets("BatchOutput")
        mP3DetailedParametersSheet = mP3Workbook.Worksheets("DetailedParameters")
        mP3RailroadUnitCostXMLSheet = mP3Workbook.Worksheets("RailroadUnitCostXML")

        'Set up the Make Whole worksheets for use
        mMWFMovement = mMWFWorkbook.Worksheets("Movement")
        mMWFCRPRESRecord = mMWFWorkbook.Worksheets("CRPRESRecord")
        mMWFCRPRES = mMWFWorkbook.Worksheets("CRPRES")
        mMWFXMLValues = mMWFWorkbook.Worksheets("XMLValues")

        ' We need to find the column which the current year is in RailroadUnitCostXML sheet
        mP3CellRange = mP3RailroadUnitCostXMLSheet.Cells.Find(cmb_URCS_Year.Text,
                                                              mP3RailroadUnitCostXMLSheet.Cells("G1:Z1"),
                                                              FindLookIn.Values,
                                                              LookAt.Whole,
                                                              SearchOrder.ByRows,
                                                              SearchDirection.Next,
                                                              False)
        If IsNothing(mP3CellRange) Then
            ' The column doesn't exist
            ' For now we'll report this as an error
            ' In the future, we'll insert the column
            MsgBox("Current Year column not found in P3XL spreadsheet!", vbOKOnly, "ERROR!")
            GoTo EndIt
        Else
            mXMLSheetColumn = Convert_Zero_Based_Column_Number_To_Text(mP3CellRange.Column)
        End If

        'We need to load the recordset with the Segments table data for the year selected.
        gbl_Database_Name = Get_Database_Name_From_SQL(cmb_URCS_Year.SelectedItem, "SEGMENTS")
        Gbl_Segments_TableName = Trim(Get_Table_Name_From_SQL(cmb_URCS_Year.SelectedItem, "SEGMENTS"))

        mOutputMaskedTableName = "Not Used"
        mOutputSegmentsTableName = "Not Used"

        ' If we are using a different year for the waybill data, we need to make accomodations
        If chk_UseDifferentYear.Checked = True Then
            Gbl_Masked_TableName = Trim(Get_Table_Name_From_SQL(cmb_Different_Year.SelectedItem, "MASKED"))
            If chk_SaveToSQL.Checked = True Then
                mOutputMaskedTableName = "WB" & cmb_Different_Year.Text & "_Masked_Using_" & cmb_URCS_Year.Text
                mOutputSegmentsTableName = "WB" & cmb_Different_Year.Text & "_Segments_Using_" & cmb_URCS_Year.Text
            End If
        Else
            Gbl_Masked_TableName = Trim(Get_Table_Name_From_SQL(cmb_URCS_Year.SelectedItem, "MASKED"))
        End If

        ' Used when we update it later
        Gbl_EValues_TableName = Trim(Get_Table_Name_From_SQL(cmb_URCS_Year.Text, "EValues"))

        'Check to see if the log file exists before creating a new one
        mLogFileName = txtFolder.Text & "\URC" & cmb_URCS_Year.Text & ".LOG"
        If My.Computer.FileSystem.FileExists(mLogFileName) Then
            For mLooper = 2 To 9
                mLogFileName = txtFolder.Text & "\URC" & cmb_URCS_Year.Text & "_v" & mLooper.ToString & ".LOG"
                If My.Computer.FileSystem.FileExists(mLogFileName) = False Then
                    Exit For
                End If
            Next
        End If
        sbLogFile = New StreamWriter(mLogFileName, False)
        sbLogFile.WriteLine("URCS Phase 3 Processing Log File")
        sbLogFile.WriteLine()
        sbLogFile.WriteLine("SQL Server selected: " & gbl_Server_Name)
        sbLogFile.WriteLine()
        sbLogFile.WriteLine("*** Processing started at " & Now.ToString("G"))

        sbLogFile.WriteLine()
        sbLogFile.Flush()

        mOutString = New StringBuilder

        If chk_UseDifferentYear.Checked = True Then
            mOutString.Append("Processed " & cmb_Different_Year.Text & " XML to " & cmb_Different_Year.Text & " Waybills")
        Else
            mOutString.Append("Processed " & cmb_URCS_Year.Text)
        End If

        If mThreesOnly = True Then
            mOutString.Append(" - 10% Run (Serial_No ends with '3')")
        End If
        If rdo_Legacy.Checked Then
            mOutString.Append(" using Legacy methodology")
        Else
            mOutString.Append(" using Current methodology")
            If chk_Cost_All_Segments.Checked = True Then
                mOutString.Append(", Costing All Segments (i.e., US, CA, MX)")
            Else
                mOutString.Append(", Costing US Segments Only")
            End If
        End If
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.WriteLine("Phase III Spreadsheet File: " & txtP3FilePath.Text)
        sbLogFile.WriteLine("MakeWhole Factor Spreadsheet File: " & txtMakeWholeFilePath.Text)
        sbLogFile.WriteLine()
        sbLogFile.WriteLine("*** Data Sources")
        sbLogFile.WriteLine()
        sbLogFile.WriteLine("Database: " & Trim(gbl_Database_Name))

        If chk_UseDifferentYear.Checked = True Then
            sbLogFile.WriteLine("Input Masked Table: " & Trim(Gbl_Masked_TableName))
            sbLogFile.WriteLine("Input Segments Table: " & Trim(Gbl_Segments_TableName))
            sbLogFile.WriteLine("Output Masked Table: " & mOutputMaskedTableName)
            sbLogFile.WriteLine("Output Segments Table: " & mOutputSegmentsTableName)
        Else
            sbLogFile.WriteLine("Masked Table: " & Trim(Gbl_Masked_TableName))
            sbLogFile.WriteLine("Segments Table: " & Trim(Gbl_Segments_TableName))
        End If

        sbLogFile.WriteLine()
        sbLogFile.WriteLine("*** Rejections")
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        btn_Return_To_MainMenu.Enabled = False
        btnExecute.Enabled = False

        'We need to load the recordset with the Segments table data for the year selected.
        gbl_Database_Name = "Waybills"
        Gbl_Segments_TableName = Trim(Get_Table_Name_From_SQL(cmb_URCS_Year.SelectedItem, "SEGMENTS"))
        Gbl_Masked_TableName = Trim(Get_Table_Name_From_SQL(cmb_URCS_Year.SelectedItem, "MASKED"))
        ' Used when we update it later
        Gbl_EValues_TableName = Trim(Get_Table_Name_From_SQL(cmb_URCS_Year.Text, "EValues"))

        OpenSQLConnection(gbl_Database_Name)

        ' If we're using a different waybill year data, we may have to clone the actual tables to the output tables
        ' both masked and segments tables
        If chk_UseDifferentYear.Checked = True And chk_SaveToSQL.Checked = True Then

            txtStatus.Text = "Creating Output Masked Table..."
            Refresh()
            Application.DoEvents()

            'Copy the Masked table
            mOutputMaskedTableName = "WB" & cmb_Different_Year.Text & "_Masked_Using_" &
                cmb_URCS_Year.Text

            If VerifyTableExist("Waybills", mOutputMaskedTableName) = True Then
                mStrSQL = "DROP TABLE " & mOutputMaskedTableName
                mSQLcmd = New SqlCommand
                mSQLcmd.Connection = gbl_SQLConnection
                mSQLcmd.CommandType = CommandType.Text
                mSQLcmd.CommandText = mStrSQL
                mSQLcmd.ExecuteNonQuery()
            End If

            mStrSQL = "Select * INTO " & mOutputMaskedTableName & " FROM WB" & cmb_URCS_Year.Text & "_Masked"
            mSQLcmd = New SqlCommand
            mSQLcmd.Connection = gbl_SQLConnection
            mSQLcmd.CommandType = CommandType.Text
            mSQLcmd.CommandText = mStrSQL
            mSQLcmd.ExecuteNonQuery()

            txtStatus.Text = "Creating Output Masked Table Index..."
            Refresh()
            Application.DoEvents()

            ' Create the index for the Masked table
            mStrSQL = "CREATE INDEX idx_" & mOutputMaskedTableName & " ON " & mOutputMaskedTableName & " (Serial_no)"
            mSQLcmd = New SqlCommand
            mSQLcmd.Connection = gbl_SQLConnection
            mSQLcmd.CommandType = CommandType.Text
            mSQLcmd.CommandText = mStrSQL
            mSQLcmd.ExecuteNonQuery()

            txtStatus.Text = "Creating Output Segments Table..."
            Refresh()
            Application.DoEvents()

            'Copy the Segments table
            mOutputSegmentsTableName = "WB" & cmb_Different_Year.Text & "_Segments_Using_" &
                cmb_URCS_Year.Text

            If VerifyTableExist("Waybills", mOutputSegmentsTableName) = True Then
                mStrSQL = "DROP TABLE " & mOutputSegmentsTableName
                mSQLcmd = New SqlCommand
                mSQLcmd.Connection = gbl_SQLConnection
                mSQLcmd.CommandType = CommandType.Text
                mSQLcmd.CommandText = mStrSQL
                mSQLcmd.ExecuteNonQuery()
            End If

            mStrSQL = "Select * INTO " & mOutputSegmentsTableName & " FROM WB" & cmb_URCS_Year.Text & "_Segments"
            mSQLcmd = New SqlCommand
            mSQLcmd.Connection = gbl_SQLConnection
            mSQLcmd.CommandType = CommandType.Text
            mSQLcmd.CommandText = mStrSQL
            mSQLcmd.ExecuteNonQuery()

            txtStatus.Text = "Creating Output Segments Table Index..."
            Refresh()
            Application.DoEvents()

            ' Create the index for the Masked table
            mStrSQL = "CREATE INDEX idx_" & mOutputSegmentsTableName & " ON " & mOutputSegmentsTableName & " (Serial_No, Seg_No)"
            mSQLcmd = New SqlCommand
            mSQLcmd.Connection = gbl_SQLConnection
            mSQLcmd.CommandType = CommandType.Text
            mSQLcmd.CommandText = mStrSQL
            mSQLcmd.ExecuteNonQuery()

            ' Set the global variables to the new tables
            Gbl_Masked_TableName = mOutputMaskedTableName
            Gbl_Segments_TableName = mOutputSegmentsTableName

        End If

        txtStatus.Text = "Getting Segments Data from SQL..."
        Refresh()

        mDataTable = New DataTable

        mStrSQL = "Select " & Gbl_Segments_TableName & ".Serial_No, " &
            Gbl_Segments_TableName & ".Seg_no, " &
            Gbl_Segments_TableName & ".RR_Num, " &
            Gbl_Segments_TableName & ".rr_Alpha, " &
            Gbl_Segments_TableName & ".RR_Dist, " &
            Gbl_Segments_TableName & ".Seg_Type, " &
            Gbl_Segments_TableName & ".Total_Segs, " &
            Gbl_Segments_TableName & ".RR_Cntry, " &
            Gbl_Masked_TableName & ".TOFC_Serv_Code, " &
            Gbl_Masked_TableName & ".STB_Car_Typ, " &
            Gbl_Masked_TableName & ".U_Cars, " &
            Gbl_Masked_TableName & ".U_TC_Units, " &
            Gbl_Masked_TableName & ".Car_Own, " &
            Gbl_Masked_TableName & ".U_Car_Init, " &
            Gbl_Masked_TableName & ".Bill_Wght_Tons, " &
            Gbl_Masked_TableName & ".STCC, " &
            Gbl_Masked_TableName & ".Int_Eq_Flg, " &
            Gbl_Masked_TableName & ".ORR, " &
            Gbl_Masked_TableName & ".ORR_Dist, JRR1_Dist, JRR2_Dist, JRR3_Dist, JRR4_Dist, JRR5_Dist, JRR6_Dist, TRR_Dist, JF," &
            Gbl_Masked_TableName & ".Exp_Factor_Th, " &
            Gbl_Masked_TableName & ".Total_Rev, " &
            Gbl_Masked_TableName & ".Total_Dist FROM " & Gbl_Segments_TableName &
            " INNER JOIN " & Gbl_Masked_TableName & " On " &
            Gbl_Segments_TableName & ".Serial_No = " & Gbl_Masked_TableName & ".Serial_No "

        If mThreesOnly = True Then
            'mStrSQL = mStrSQL & " where (right(" & Gbl_Masked_TableName & ".serial_no,1) = 3) And " & Gbl_Masked_TableName & ".serial_no < 250000)"
            mStrSQL = mStrSQL & " where right(" & Gbl_Masked_TableName & ".serial_no,2) = 33"
        End If

        Using daAdapter As New SqlClient.SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        If mDataTable.Rows.Count = 0 Then
            MsgBox("No records found In Segments table!", vbOKOnly, "ERROR!")
            GoTo EndIt
        Else
            mNumOfMaskedSegments = mDataTable.Rows.Count
        End If

        ' Ensure that the Residuals array is cleared
        For mLooper = 1 To 9
            For mLooper2 = 1 To 17
                mResiduals(mLooper, mLooper2) = 0
            Next
        Next

        ' Ensure that the rejected segments and waybill counters are initialized to zero
        For mLooper = 1 To 7
            mRejectedSegments(mLooper) = 0
            mRejectedWaybills(mLooper) = 0
        Next

        ' Clear out the P3 BatchInput sheet rows that we don't need or use
        mP3CellRange = mP3BatchCostProgramSheet.Cells(4, 0, mP3BatchCostProgramSheet.UsedRange.RowCount, mP3BatchCostProgramSheet.UsedRange.ColumnCount)
        mP3CellRange.Clear()

        ' Do the same for the P3 BatchOutput sheet
        mP3CellRange = mP3BatchOutputSheet.Cells(5, 0, mP3BatchOutputSheet.UsedRange.RowCount, mP3BatchOutputSheet.UsedRange.ColumnCount)
        mP3CellRange.Clear()

        ' Set the CRPRES sheet, cells B5:S13 to zeros
        mMWFCellRange = mMWFCRPRES.Cells("B5:S13")
        mMWFCellRange.Value = 0

        If InStr(mP3RailroadCostProgramSheet.Cells("D77").Formula, "AdjustLUMs") > 0 Then
            ' Process with multicar shipment size at 75 cars
            mUnitTrainDef = 75
        Else
            ' Process with multicar shipment size at 50 cars
            mUnitTrainDef = 50
        End If

        mCounter = 0

        ' Check to see what, if any, csv files are requested via checkboxes on the form
        sb100Series_SysAvg = Nothing
        sb100Series_EffAdj = Nothing
        sb100Series_Costed = Nothing
        If chk100Series.Checked Then
            If mUnitTrainDef = 50 Then
                sb100Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L100Series_SysAvg.CSV", False)
                sb100Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L100Series_EffAdj.CSV", False)
                sb100Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L100Series_Costed.CSV", False)
            Else
                sb100Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L100Series_EP431Sub4_SysAvg.CSV", False)
                sb100Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L100Series_EP431Sub4_EffAdj.CSV", False)
                sb100Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L100Series_EP431Sub4_Costed.CSV", False)
            End If
            mOutString = New StringBuilder
            mOutString.Append("Serial_no,Seg_no,")
            mOutString.Append(Get100SeriesHeaders(mP3BatchOutputSheet))
            sb100Series_SysAvg.WriteLine(mOutString.ToString)
            sb100Series_EffAdj.WriteLine(mOutString.ToString)
            sb100Series_Costed.WriteLine(mOutString.ToString)
        End If

        sb200Series_SysAvg = Nothing
        sb200Series_EffAdj = Nothing
        sb200Series_Costed = Nothing
        If chk200Series.Checked Then
            If mUnitTrainDef = 50 Then
                sb200Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L200Series_SysAvg.CSV", False)
                sb200Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L200Series_EffAdj.CSV", False)
                sb200Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L200Series_Costed.CSV", False)
            Else
                sb200Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L200Series_EP431Sub4_SysAvg.CSV", False)
                sb200Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L200Series_EP431Sub4_EffAdj.CSV", False)
                sb200Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L200Series_EP431Sub4_Costed.CSV", False)
            End If
            mOutString = New StringBuilder
            mOutString.Append("Serial_no,Seg_no,")
            mOutString.Append(Get200SeriesHeaders(mP3BatchOutputSheet))
            sb200Series_SysAvg.WriteLine(mOutString.ToString)
            sb200Series_EffAdj.WriteLine(mOutString.ToString)
            sb200Series_Costed.WriteLine(mOutString.ToString)
        End If

        sb300Series_SysAvg = Nothing
        sb300Series_EffAdj = Nothing
        sb300Series_Costed = Nothing
        If chk300Series.Checked Then
            If mUnitTrainDef = 50 Then
                sb300Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L300Series_SysAvg.CSV", False)
                sb300Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L300Series_EffAdj.CSV", False)
                sb300Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L300Series_Costed.CSV", False)
            Else
                sb300Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L300Series_EP431Sub4_SysAvg.CSV", False)
                sb300Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L300Series_EP431Sub4_EffAdj.CSV", False)
                sb300Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L300Series_EP431Sub4_Costed.CSV", False)
            End If
            mOutString = New StringBuilder
            mOutString.Append("Serial_no,Seg_no,")
            mOutString.Append(Get300SeriesHeaders(mP3BatchOutputSheet))
            sb300Series_SysAvg.WriteLine(mOutString.ToString)
            sb300Series_EffAdj.WriteLine(mOutString.ToString)
            sb300Series_Costed.WriteLine(mOutString.ToString)
        End If

        sb400Series_SysAvg = Nothing
        sb400Series_EffAdj = Nothing
        sb400Series_Costed = Nothing
        If chk400Series.Checked Then
            If mUnitTrainDef = 50 Then
                sb400Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L400Series_SysAvg.CSV", False)
                sb400Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L400Series_EffAdj.CSV", False)
                sb400Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L400Series_Costed.CSV", False)
            Else
                sb400Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L400Series_EP431Sub4_SysAvg.CSV", False)
                sb400Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L400Series_EP431Sub4_EffAdj.CSV", False)
                sb400Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L400Series_EP431Sub4_Costed.CSV", False)
            End If
            mOutString = New StringBuilder
            mOutString.Append("Serial_no,Seg_no,")
            mOutString.Append(Get400SeriesHeaders(mP3BatchOutputSheet))
            sb400Series_SysAvg.WriteLine(mOutString.ToString)
            sb400Series_EffAdj.WriteLine(mOutString.ToString)
            sb400Series_Costed.WriteLine(mOutString.ToString)
        End If

        sb500Series_SysAvg = Nothing
        sb500Series_EffAdj = Nothing
        sb500Series_Costed = Nothing
        If chk500Series.Checked Then
            If mUnitTrainDef = 50 Then
                sb500Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L500Series_SysAvg.CSV", False)
                sb500Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L500Series_EffAdj.CSV", False)
                sb500Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L500Series_Costed.CSV", False)
            Else
                sb500Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L500Series_EP431Sub4_SysAvg.CSV", False)
                sb500Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L500Series_EP431Sub4_EffAdj.CSV", False)
                sb500Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L500Series_EP431Sub4_Costed.CSV", False)
            End If
            mOutString = New StringBuilder
            mOutString.Append("Serial_no,Seg_no,")
            mOutString.Append(Get500SeriesHeaders(mP3BatchOutputSheet))
            sb500Series_SysAvg.WriteLine(mOutString.ToString)
            sb500Series_EffAdj.WriteLine(mOutString.ToString)
            sb500Series_Costed.WriteLine(mOutString.ToString)
        End If

        sb600Series_SysAvg = Nothing
        sb600Series_EffAdj = Nothing
        sb600Series_Costed = Nothing
        If chk600Series.Checked Then
            If mUnitTrainDef = 50 Then
                sb600Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L600Series_SysAvg.CSV", False)
                sb600Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L600Series_EffAdj.CSV", False)
                sb600Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L600Series_Costed.CSV", False)
            Else
                sb600Series_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L600Series_EP431Sub4_SysAvg.CSV", False)
                sb600Series_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L600Series_EP431Sub4_EffAdj.CSV", False)
                sb600Series_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_L600Series_EP431Sub4_Costed.CSV", False)
            End If
            mOutString = New StringBuilder
            mOutString.Append("Serial_no,Seg_no,")
            mOutString.Append(Get600SeriesHeaders(mP3BatchOutputSheet))
            sb600Series_SysAvg.WriteLine(mOutString.ToString)
            sb600Series_EffAdj.WriteLine(mOutString.ToString)
            sb600Series_Costed.WriteLine(mOutString.ToString)
        End If

        sbCRPRES = Nothing
        sbCRPRESRecord = Nothing
        If chk_Save_CRPRESRecords.Checked Then

            sbCRPRESRecord = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_CRPRESRecord.CSV", False)
            mOutString = New StringBuilder
            mOutString.Append("ID,")
            For mLooper = 1 To 18
                mOutString.Append("C" & mLooper.ToString & ",")
            Next
            mOutString.Append("C19")
            sbCRPRESRecord.WriteLine(mOutString.ToString)
        End If

        sbCRPRES = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_CRPRES.CSV", False)
        mOutString = New StringBuilder
        mOutString.Append("RRICC,")
        For mLooper = 2 To 18
            mOutString.Append("C" & mLooper.ToString & ",")
        Next
        mOutString.Append("C19")
        sbCRPRES.WriteLine(mOutString.ToString)

        sbResults_SysAvg = Nothing
        sbResults_EffAdj = Nothing
        sbResults_Costed = Nothing
        If chkSaveResults.Checked Then
            If mUnitTrainDef = 50 Then
                sbResults_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_Results_SysAvg.CSV", False)
                sbResults_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_Results_EffAdj.CSV", False)
                sbResults_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_Results_Costed.CSV", False)
            Else
                sbResults_SysAvg = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_Results_EP431Sub4_SysAvg.CSV", False)
                sbResults_EffAdj = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_Results_EP431Sub4_EffAdj.CSV", False)
                sbResults_Costed = New StreamWriter(txtFolder.Text & "\" & cmb_URCS_Year.Text & "_Results_EP431Sub4_Costed.CSV", False)
            End If

            mOutString = New StringBuilder
            mOutString.Append("Serial_no,Seg_no,")

            'Get headers for Results files from line 4 in P3 BatchOutput sheet
            mP3CellRange = mP3BatchOutputSheet.Cells
            For Each cell As IRange In mP3CellRange("B4:J4")
                mVal = cell.Value
                mOutString.Append(mVal.ToString & ",")
            Next cell
            ' Set the header for L569
            mOutString.Append(mP3CellRange("KZ4").Value & ",")
            ' Set the header for L587
            mOutString.Append(mP3CellRange("LR4").Value & ",")
            ' Set the header for L696
            mOutString.Append(mP3CellRange("PJ4").Value & ",")
            ' Set the header for L699
            mOutString.Append(mP3CellRange("PM4").Value & ",")
            ' Set the header for L700
            mOutString.Append(mP3CellRange("PN4").Value & ",")
            ' Set the header for Exp_Factor_Th
            mOutString.Append("Exp_Factor_Th" & ",")
            ' Set the header for ID
            mOutString.Append(mP3CellRange("A4").Value)
            sbResults_SysAvg.WriteLine(mOutString.ToString)
            sbResults_EffAdj.WriteLine(mOutString.ToString)
            sbResults_Costed.WriteLine(mOutString.ToString)
        End If

        ' Initialize the CRPRES array and Cost breakdown array
        For mLooper = 1 To 9
            For mLooper2 = 2 To 19
                mCRPRESArray(mLooper, mLooper2) = 0
            Next

            For mLooper2 = 1 To 4
                mCostBreakdown(mLooper, mLooper2) = 0
            Next
        Next

        mCostBreakdownCol4 = 0
        mCostBreakdownCol1Total = 0
        mCostBreakdownCol2Total = 0
        mCostBreakdownCol3Total = 0
        mCostBreakdownCol4Total = 0

        mPass1Starts = Now()

        mOutString = New StringBuilder
        mOutString.Append("*** First pass thru Segments table to calculate the Make-Whole Factors started at " & DateTime.Now.ToString("G") & ". ")
        sbLogFile.WriteLine(mOutString.ToString)

        Text = "Pass 1 of 3"
        txtStatus.Text = "Pass 1 of 3 - Creating Make-Whole Factors for Segments record 1 of " _
                    & mDataTable.Rows.Count.ToString & "..."
        Refresh()
        Application.DoEvents()

        ' Start looping thru the DataTable
        For mLooper = 0 To mDataTable.Rows.Count - 1
            mCounter = mCounter + 1
            If mCounter Mod 100 = 0 Then
                txtStatus.Text = "Pass 1 of 3  - Creating Make-Whole Factors for Segments record " & mCounter.ToString & " of " _
                    & mDataTable.Rows.Count.ToString & " - " & ((mLooper / mDataTable.Rows.Count) * 100).ToString("N1") & "%..."
                Text = "Pass 1 of 3 - " & ((mLooper / mDataTable.Rows.Count) * 100).ToString("N1") & "%"
                Refresh()
                Application.DoEvents()
            End If

            If mCounter Mod 100000 = 0 Then
                sbLogFile.WriteLine(“*** Processing record " & mCounter.ToString & " of the Segments table at " & DateTime.Now.ToString("G"))
            End If

            ' Increment Serial Number Counter if serial # is different
            If mSerial_No <> mDataTable.Rows(mLooper)("Serial_No").ToString Then
                mNumOfWaybillRecords = mNumOfWaybillRecords + 1
                mSerialChanged = True
            Else
                mSerialChanged = False
            End If

            'Clean up any old data
            mP3CellRange = mP3RailroadCostProgramSheet.Cells("D6: D11")
            mP3CellRange.Value = ""
            mP3CellRange = mP3RailroadCostProgramSheet.Cells("E6:H14")
            mP3CellRange.Value = ""

            ' Set the Efficiency Adjustment option to "Y" for the first calculation
            mP3DetailedParametersSheet.Cells("C90").Value = "Y"

            ' Save this for the MakeWhole Factor Movement sheet
            mExp_Factor_Th = mDataTable.Rows(mLooper)("Exp_Factor_Th")

            'Start laying in the values to line 5 base zero of the spreadsheet
            mSerial_No = mDataTable.Rows(mLooper)("Serial_no")
            mSeg_No = mDataTable.Rows(mLooper)("Seg_no")

            ' Input Parameter: ID
            mP3BatchCostProgramSheet.Cells("A5").Value = mDataTable.Rows(mLooper)("Serial_No") & (0.1 * mDataTable.Rows(mLooper)("Seg_No").ToString)

            ' Input Parameter: RR
            mP3BatchCostProgramSheet.Cells("B5").Value = mDataTable.Rows(mLooper)("RR_Num")

            ' Input Parameter: DIS
            mP3BatchCostProgramSheet.Cells("C5").Value = mDataTable.Rows(mLooper)("RR_Dist") * 0.1

            ' Input Parameter: SG
            mP3BatchCostProgramSheet.Cells("D5").Value = mDataTable.Rows(mLooper)("Seg_Type")

            ' Input Parameter: FC
            If rdo_Legacy.Checked = True Then
                ' Logic if running as Legacy 
                mTOFCMove = False
                Select Case mDataTable.Rows(mLooper)("STB_Car_Typ")
                    Case 46, 49, 52, 54
                        If (mDataTable.Rows(mLooper)("TOFC_Serv_Code") <> "") Or (mDataTable.Rows(mLooper)("Int_Eq_Flg") = 2) Then
                            mTOFCMove = True
                            mP3BatchCostProgramSheet.Cells("E5").Value = 46
                        Else
                            mP3BatchCostProgramSheet.Cells("E5").Value = mDataTable.Rows(mLooper)("STB_Car_typ")
                        End If
                    Case Else
                        mP3BatchCostProgramSheet.Cells("E5").Value = mDataTable.Rows(mLooper)("STB_Car_typ")
                End Select
            Else
                ' Check to see if we have a value for the Car Type.  If not, use default
                If String.IsNullOrEmpty(mDataTable.Rows(mLooper)("STB_Car_typ")) Then
                    mP3BatchCostProgramSheet.Cells("E5").Value = 52
                Else
                    mP3BatchCostProgramSheet.Cells("E5").Value = mDataTable.Rows(mLooper)("STB_Car_typ")
                End If

                mTOFCMove = False
                Select Case mDataTable.Rows(mLooper)("STB_Car_Typ")
                    Case 46, 48, 49, 52, 54
                        If (mDataTable.Rows(mLooper)("TOFC_Serv_Code") <> "") Or (mDataTable.Rows(mLooper)("Int_Eq_Flg") = 2) Then
                            ' If RoadRailer (Int_Eq_Flg = 2) set the cartyp to TOFC flat Car (46)
                            If mDataTable.Rows(mLooper)("Int_Eq_Flg") = 2 Then
                                mP3BatchCostProgramSheet.Cells("E5").Value = 46
                            End If
                            mTOFCMove = True
                        End If
                End Select
            End If

            ' Input Parameter: NC
            'Check for the larger of U_Cars or U_TC_Units
            If mDataTable.Rows(mLooper)("U_Cars") > mDataTable.Rows(mLooper)("U_TC_Units") Then
                mP3BatchCostProgramSheet.Cells("F5").Value = mDataTable.Rows(mLooper)("U_Cars")
            Else
                mP3BatchCostProgramSheet.Cells("F5").Value = mDataTable.Rows(mLooper)("U_TC_Units")
            End If

            ' Input Parameter: OWN
            If rdo_Legacy.Checked = True Then
                ' Logic if running as Legacy
                Select Case Trim(mDataTable.Rows(mLooper)("U_Car_init"))
                    Case "ABOX", "RBOX", "CSX", "CSXT", "GONX"
                        mP3BatchCostProgramSheet.Cells("G5").Value = "R"
                    Case Else
                        If Strings.Right(Trim(mDataTable.Rows(mLooper)("U_Car_Init")), 1) = "X" Then
                            mP3BatchCostProgramSheet.Cells("G5").Value = "P"
                        Else
                            mP3BatchCostProgramSheet.Cells("G5").Value = "R"
                        End If
                End Select
            Else
                Select Case CInt(cmb_URCS_Year.Text)
                    Case 2015, 2016, 2017
                        If Strings.Right(Trim(mDataTable.Rows(mLooper)("U_Car_Init")), 1) = "X" Then
                            mP3BatchCostProgramSheet.Cells("G5").Value = "P"
                        Else
                            mP3BatchCostProgramSheet.Cells("G5").Value = "R"
                        End If
                    Case Else
                        ' Use value in Car_Own field unless that is blank
                        Select Case Trim(mDataTable.Rows(mLooper)("Car_Own"))
                            Case "R", "P", "T"
                                mP3BatchCostProgramSheet.Cells("G5").Value = mDataTable.Rows(mLooper)("Car_Own")
                            Case Else
                                If Strings.Right(Trim(mDataTable.Rows(mLooper)("U_Car_Init")), 1) = "X" Then
                                    mP3BatchCostProgramSheet.Cells("G5").Value = "P"
                                Else
                                    mP3BatchCostProgramSheet.Cells("G5").Value = "R"
                                End If
                        End Select
                End Select
            End If

            '' Use value in Car_Own field unless that is blank
            'Select Case Trim(mDataTable.Rows(mLooper)("Car_Own"))
            '    Case "R", "P", "T"
            '        mP3BatchCostProgramSheet.Cells("G5").Value = mDataTable.Rows(mLooper)("Car_Own")
            '    Case Else
            '        If Strings.Right(Trim(mDataTable.Rows(mLooper)("U_Car_Init")), 1) = "X" Then
            '            mP3BatchCostProgramSheet.Cells("G5").Value = "P"
            '        Else
            '            mP3BatchCostProgramSheet.Cells("G5").Value = "R"
            '        End If
            'End Select


            ' Input Parameter: WT
            ' Calculate the tons per car or tons per TCU value
            If mTOFCMove Then
                If mDataTable.Rows(mLooper)("U_TC_Units") > 0 Then
                    mP3BatchCostProgramSheet.Cells("H5").Value = mDataTable.Rows(mLooper)("Bill_Wght_Tons") / mDataTable.Rows(mLooper)("U_TC_Units")
                Else
                    mP3BatchCostProgramSheet.Cells("H5").Value = 0
                End If
            Else
                If mDataTable.Rows(mLooper)("U_Cars") > 0 Then
                    mP3BatchCostProgramSheet.Cells("H5").Value = mDataTable.Rows(mLooper)("Bill_Wght_Tons") / mDataTable.Rows(mLooper)("U_Cars")
                Else
                    mP3BatchCostProgramSheet.Cells("H5").Value = 0
                End If
            End If

            ' Input Parameter: COM
            mP3BatchCostProgramSheet.Cells("I5").Value = "'" & Strings.Left(mDataTable.Rows(mLooper)("STCC"), 5)

            ' Input Parameter: SZ (Set the shipment size - legacy v. EP431 Sub 4)
            If mTOFCMove = True Then
                mP3BatchCostProgramSheet.Cells("J5").Value = "Intermodal"
            ElseIf mDataTable.Rows(mLooper)("U_Cars") <= 5 Then
                mP3BatchCostProgramSheet.Cells("J5").Value = "Single"
            ElseIf mDataTable.Rows(mLooper)("U_Cars") < mUnitTrainDef Then
                mP3BatchCostProgramSheet.Cells("J5").Value = "Multi"
            Else
                mP3BatchCostProgramSheet.Cells("J5").Value = "Unit"
            End If

            ' Input Parameter: L102 (Set the Circuity to 1)
            mP3BatchCostProgramSheet.Cells("K5").Value = 1

            ' Input Parameter: L569
            If mDataTable.Rows(mLooper)("Int_Eq_Flg") = 2 Then
                mP3BatchCostProgramSheet.Cells("L5").Value = "TCS"
            Else
                mP3BatchCostProgramSheet.Cells("L5").Value = Trim(mDataTable.Rows(mLooper)("TOFC_Serv_Code"))
            End If

            ' Input Parameter: TDIS
            mP3BatchCostProgramSheet.Cells("M5").Value = mDataTable.Rows(mLooper)("Total_Dist") * 0.1

            ' Input Parameter: ORR
            mP3BatchCostProgramSheet.Cells("N5").Value = mDataTable.Rows(mLooper)("ORR")

            ' Copy the formulas for P5:AC5
            With mP3BatchCostProgramSheet
                mP3CellRange = .Cells
                mP3CellRange("P1:AC1").Copy(.Cells("P5:AC5"))
            End With

            ' Implement Current Waybill Logic that excludes certain records from being costed
            mProcess = True
            mGoodSegment = True

            ' Skip record if Freight Car (FC) equals 0
            If mP3BatchCostProgramSheet.Cells("E5").Value = 0 Then
                If mSerialChanged = True Then
                    mRejectedWaybills(BadCarType) = mRejectedWaybills(BadCarType) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        " - Invalid Car Type (STB_Car_Typ = " &
                                        mP3BatchCostProgramSheet.Cells("E5").Value & ")")
                mRejectedSegments(BadCarType) = mRejectedSegments(BadCarType) + 1
                mGoodSegment = False
            End If

            ' Skip record if Weight (WT) is less than 1 ton
            If mP3BatchCostProgramSheet.Cells("H5").Value < 1 Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadTonnage) = mRejectedWaybills(BadTonnage) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        " - Weight < 1 ton (WT = " &
                                        CDbl(mP3BatchCostProgramSheet.Cells("H5").Value).ToString("N2") & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadTonnage) = mRejectedSegments(BadTonnage) + 1
                    mGoodSegment = False
                End If
            End If

            ' Skip record if Total Revenue is zero
            If mDataTable.Rows(mLooper)("total_rev") = 0 Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadRevenue) = mRejectedWaybills(BadRevenue) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        " - Total Revenue = " & mDataTable.Rows(mLooper)("total_rev").ToString)
                If mGoodSegment = True Then
                    mRejectedSegments(BadRevenue) = mRejectedSegments(BadRevenue) + 1
                    mGoodSegment = False
                End If
            End If

            ' Skip record if 1 railroad (Total_Segs=1) and TDIS < 1.5 (rounded)
            If (mDataTable.Rows(mLooper)("total_Segs") = 1) And mP3BatchCostProgramSheet.Cells("M5").Value < 1.5 Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        " - Single-Carrier Move < 2 miles (Total_Segs = " &
                                        mDataTable.Rows(mLooper)("total_Segs").ToString & ", " &
                                        "TDIS = " & mP3BatchCostProgramSheet.Cells("M5").Value & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (mDataTable.Rows(mLooper)("total_Segs") > 1) And mP3BatchCostProgramSheet.Cells("M5").Value < 4.5 Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                    " - Multi-Carrier Move < 5 miles (Total_Segs = " &
                                        mDataTable.Rows(mLooper)("total_Segs").ToString & ", " &
                                        "TDIS = " & mP3BatchCostProgramSheet.Cells("M5").Value & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("ORR_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 0) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                    " - Invalid ORR_Dist (ORR_Dist = " &
                                    (mDataTable.Rows(mLooper)("ORR_Dist") * 0.1) & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = False) And ((mDataTable.Rows(mLooper)("ORR_Dist") * 0.1) < 0.5) And
                    (mDataTable.Rows(mLooper)("JF") >= 0) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                            ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                            " - Invalid ORR_Dist (ORR_Dist = " & (mDataTable.Rows(mLooper)("ORR_Dist") * 0.1).ToString & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("TRR_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 1) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid TRR_Dist (TRR_Dist = " &
                    (mDataTable.Rows(mLooper)("TRR_Dist") * 0.1) & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = False) And ((mDataTable.Rows(mLooper)("TRR_Dist") * 0.1) < 0.5) And
                (mDataTable.Rows(mLooper)("JF") >= 1) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid TRR_Dist (TRR_Dist = " &
                    (mDataTable.Rows(mLooper)("TRR_Dist") * 0.1) & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR1_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 2) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR1_Dist (JRR1_Dist = " &
                    (mDataTable.Rows(mLooper)("JRR1_Dist") * 0.1) & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = False) And ((mDataTable.Rows(mLooper)("JRR1_Dist") * 0.1) < 0.5) And
                (mDataTable.Rows(mLooper)("JF") >= 2) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR1_Dist (JRR1_Dist = " &
                    (mDataTable.Rows(mLooper)("JRR1_Dist") * 0.1).ToString & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR2_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 3) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR2_Dist (JRR2_Dist = " &
                   (mDataTable.Rows(mLooper)("JRR2_Dist") * 0.1) & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = False) And
                   ((mDataTable.Rows(mLooper)("JRR2_Dist") * 0.1) < 0.5) And
                   (mDataTable.Rows(mLooper)("JF") >= 3) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR2_Dist (JRR2_Dist = " &
                    (mDataTable.Rows(mLooper)("JRR2_Dist") * 0.1).ToString & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR3_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 4) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR3_Dist (JRR3_Dist = " &
                    (mDataTable.Rows(mLooper)("JRR3_Dist") * 0.1) & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = False) And
                   ((mDataTable.Rows(mLooper)("JRR3_Dist") * 0.1) < 0.5) And
                   (mDataTable.Rows(mLooper)("JF") >= 4) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR3_Dist (JRR3_Dist = " &
                    (mDataTable.Rows(mLooper)("JRR3_Dist") * 0.1).ToString & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR4_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 5) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR4_Dist (JRR4_Dist = " &
                    (mDataTable.Rows(mLooper)("JRR4_Dist") * 0.1) & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = False) And
                ((mDataTable.Rows(mLooper)("JRR4_Dist") * 0.1) < 0.5) And (mDataTable.Rows(mLooper)("JF") >= 5) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR4_Dist (JRR4_Dist = " &
                    (mDataTable.Rows(mLooper)("JRR4_Dist") * 0.1).ToString & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR5_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 6) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR5_Dist (JRR5_Dist = " &
                    (mDataTable.Rows(mLooper)("JRR5_Dist") * 0.1) & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = False) And
                   ((mDataTable.Rows(mLooper)("JRR5_Dist") * 0.1) < 0.5) And
                   (mDataTable.Rows(mLooper)("JF") >= 6) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR5_Dist (JRR5_Dist = " &
                    (mDataTable.Rows(mLooper)("JRR5_Dist") * 0.1).ToString & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR6_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 7) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR6_Dist (JRR6_Dist = " &
                    (mDataTable.Rows(mLooper)("JRR6_Dist") * 0.1) & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            ElseIf (rdo_Legacy.Checked = False) And
                   ((mDataTable.Rows(mLooper)("JRR6_Dist") * 0.1) < 0.5) And
                   (mDataTable.Rows(mLooper)("JF") >= 7) Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                    " - Invalid JRR6_Dist (JRR6_Dist = " &
                    (mDataTable.Rows(mLooper)("JRR6_Dist") * 0.1).ToString & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
                    mGoodSegment = False
                End If
            End If

            ' Skip record if FC = 46 and SZ not Intermodal
            If (mP3BatchCostProgramSheet.Cells("E5").Value = "46") And (mP3BatchCostProgramSheet.Cells("J5").Value <> "Intermodal") Then
                If mSerialChanged = True And mProcess = True Then
                    mRejectedWaybills(BadIntermodalSpec) = mRejectedWaybills(BadIntermodalSpec) + 1
                End If
                mProcess = False
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                " - Invalid Intermodal Definition (FC = 46 / SZ = " &
                                mP3BatchCostProgramSheet.Cells("J5").Value.ToString & ")")
                If mGoodSegment = True Then
                    mRejectedSegments(BadIntermodalSpec) = mRejectedSegments(BadIntermodalSpec) + 1
                    mGoodSegment = False
                End If
            End If

            If rdo_Legacy.Checked = False Then  ' Non-Legacy processing
                ' If selected cost non-US segment
                If (chk_Cost_All_Segments.Checked = True) And (mProcess = True) Then
                    ' We will cost this segment
                Else
                    If (mDataTable.Rows(mLooper)("RR_Cntry") <> "US") Then
                        If mSerialChanged = True And mProcess = True Then
                            mRejectedWaybills(NonUSSegment) = mRejectedWaybills(NonUSSegment) + 1
                        End If
                        mProcess = False
                        sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                    ", Segment " & mDataTable.Rows(mLooper)("Seg_no") &
                    " ignored - Non US Segment (RR_Cntry = " &
                    Trim(mDataTable.Rows(mLooper)("RR_Cntry")) & ")")
                        If mGoodSegment = True Then
                            mRejectedSegments(NonUSSegment) = mRejectedSegments(NonUSSegment) + 1
                            mGoodSegment = False
                        End If
                    End If
                End If
            End If

            sbLogFile.Flush()

            ' Force the SpreadsheetGear sheet to "translate" the SQL data into user input data
            mP3Workbook.WorkbookSet.Calculate()

            ' Check the results in P5:AC5 for any bad VLOOKUPs
            mLookupOkay = True

            ' Check RR
            If mP3BatchCostProgramSheet.Cells("Q5").Value.ToString = "NA" Then
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        "  WARNING Bad RR Translation = " & mP3BatchCostProgramSheet.Cells("B5").Value)
                mLookupOkay = False
            End If

            ' Check SG
            If mP3BatchCostProgramSheet.Cells("S5").Value.ToString = "NA" Then
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        "  WARNING Bad SG Translation = " & mP3BatchCostProgramSheet.Cells("D5").Value)
                mLookupOkay = False
            End If

            ' Check FC
            If mP3BatchCostProgramSheet.Cells("T5").Value.ToString = "NA" Then
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        "  WARNING Bad FC Translation = " & mP3BatchCostProgramSheet.Cells("E5").Value)
                mLookupOkay = False
            End If

            ' Check OWN
            If mP3BatchCostProgramSheet.Cells("V5").Value.ToString = "NA" Then
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        "  WARNING Bad OWN Translation = " & mP3BatchCostProgramSheet.Cells("G5").Value)
                mLookupOkay = False
            End If

            ' Check COM
            If mP3BatchCostProgramSheet.Cells("X5").Value.ToString = "NA" Then
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        "  WARNING Bad COM Translation = " & mP3BatchCostProgramSheet.Cells("I5").Value)
                mLookupOkay = False
            End If

            ' Check SZ
            If mP3BatchCostProgramSheet.Cells("Y5").Value.ToString = "NA" Then
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        "  WARNING Bad SZ Translation = " & mP3BatchCostProgramSheet.Cells("J5").Value)
                mLookupOkay = False
            End If

            ' Check TOFC
            If mP3BatchCostProgramSheet.Cells("AB5").Value.ToString = "NA" Then
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        "  WARNING Bad TOFC Translation = " & mP3BatchCostProgramSheet.Cells("L5").Value)
                mLookupOkay = False
            End If

            ' Check ORR
            If mP3BatchCostProgramSheet.Cells("AC5").Value.ToString = "NA" Then
                sbLogFile.WriteLine("Serial " & mDataTable.Rows(mLooper)("Serial_no") &
                                        ", Seg " & mDataTable.Rows(mLooper)("Seg_no") &
                                        "  WARNING Bad ORR Translation = " & mP3BatchCostProgramSheet.Cells("N5").Value)
                mLookupOkay = False
            End If

            sbLogFile.Flush()

            If mLookupOkay = False Then
                mNumOfBadLookups = mNumOfBadLookups + 1
                mProcess = False
            End If

            If mProcess = True Then

                ' Count this waybill as being costed
                If mPreviousSerialNo.ToString <> mSerial_No Then
                    mNumOfCostedWaybills = mNumOfCostedWaybills + 1
                    If mTOFCMove = True Then
                        mCostedWaybills(IntermodalMove) = mCostedWaybills(IntermodalMove) + 1
                    ElseIf mDataTable.Rows(mLooper)("U_Cars") <= 5 Then
                        mCostedWaybills(SingleMove) = mCostedWaybills(SingleMove) + 1
                    ElseIf mDataTable.Rows(mLooper)("U_Cars") < mUnitTrainDef Then
                        mCostedWaybills(MultiMove) = mCostedWaybills(MultiMove) + 1
                    Else
                        mCostedWaybills(UnitMove) = mCostedWaybills(UnitMove) + 1
                    End If
                    mPreviousSerialNo = mSerial_No
                End If

                ' Force the SpreadsheetGear sheet to "translate" the SQL data into user input data
                ' mP3Workbook.WorkbookSet.Calculate()

                If chk_UseDifferentYear.Checked = True Then
                    mP3RailroadCostProgramSheet.Cells("D1").Value = cmb_Different_Year.SelectedItem.ToString                    'Year
                Else
                    mP3RailroadCostProgramSheet.Cells("D1").Value = cmb_URCS_Year.SelectedItem.ToString                         'Year
                End If

                mP3RailroadCostProgramSheet.Cells("D5").Value = mP3BatchCostProgramSheet.Cells("Q5").Value                      'RR
                mP3RailroadCostProgramSheet.Cells("D6").Value = mP3BatchCostProgramSheet.Cells("R5").Value                      'DIS
                mP3RailroadCostProgramSheet.Cells("D7").Value = mP3BatchCostProgramSheet.Cells("S5").Value                      'SG
                mP3RailroadCostProgramSheet.Cells("D8").Value = mP3BatchCostProgramSheet.Cells("T5").Value                      'FC
                mP3RailroadCostProgramSheet.Cells("D9").Value = mP3BatchCostProgramSheet.Cells("U5").Value                      'NC
                mP3RailroadCostProgramSheet.Cells("D10").Value = mP3BatchCostProgramSheet.Cells("V5").Value                     'OWN
                mP3RailroadCostProgramSheet.Cells("D11").Value = mP3BatchCostProgramSheet.Cells("W5").Value                     'WT
                mP3RailroadCostProgramSheet.Cells("D12").Value = Replace(mP3BatchCostProgramSheet.Cells("X5").Value, "'", ";")  'COM
                mP3RailroadCostProgramSheet.Cells("D13").Value = mP3BatchCostProgramSheet.Cells("Y5").Value                     'SZ

                ' Copy the OtherRRDist (if any) from the BatchCostProgram sheet to the RailroadCostProgram sheet
                mP3RailroadCostProgramSheet.Cells("H6").Value = mP3BatchCostProgramSheet.Cells("Z5").Value

                ' Copy the Circuity (if any) from the BatchCostProgram sheet to the DetailedParameters sheet
                mP3DetailedParametersSheet.Cells("C5").Value = mP3BatchCostProgramSheet.Cells("AA5").Value

                ' Copy the TOFC Plan (if any) from the BatchCostProgram sheet to the DetailedParameters sheet
                ' mP3DetailedParametersSheet.Cells("C14").Value = mP3BatchCostProgramSheet.Cells("L5").Value
                mP3DetailedParametersSheet.Cells("C14").Value = mP3BatchCostProgramSheet.Cells("AB5").Value

                ' Copy the ORR (if any) from the BatchCostProgram sheet to the DetailedParameters sheet
                ' mP3DetailedParametersSheet.Cells("J11").Value = mP3BatchCostProgramSheet.Cells("N5").Value
                mP3DetailedParametersSheet.Cells("J11").Value = mP3BatchCostProgramSheet.Cells("AC5").Value

                'Set the efficiency adjustment flag to Y in the DetailedParameters sheet
                mP3DetailedParametersSheet.Cells("C90").Value = "Y"

                'Calculate the costs for this segment
                mP3Workbook.WorkbookSet.Calculate()

                ' Copy P5:Y5 on the BatchcostProgram sheet to the BatchOutput sheet
                mP3CellRange = mP3BatchCostProgramSheet.Cells("P5:Y5")
                mP3CellRange.Copy(mP3BatchOutputSheet.Cells("A5"), PasteType.Values, PasteOperation.None, False, False)

                ' Copy K1:PQ1 on the BatchOutput sheet to K5:PQ5 on the same sheet
                mP3CellRange = mP3BatchOutputSheet.Cells("K1:PQ1")
                mP3CellRange.Copy(mP3BatchOutputSheet.Cells("K5"), PasteType.Values, PasteOperation.None, False, False)

                ' Increment the processed segment counter for this railroad.
                mCostedSegmentsByRailroad(mMWFCRPRESRecord.Cells("B5").Value) = mCostedSegmentsByRailroad(mMWFCRPRESRecord.Cells("B5").Value) + 1

                'Increment the Segments Costed counters
                Select Case mP3BatchCostProgramSheet.Cells("J5").Value
                    Case "Intermodal"
                        mCostedSegments(IntermodalMove) = mCostedSegments(IntermodalMove) + 1
                    Case "Single"
                        mCostedSegments(SingleMove) = mCostedSegments(SingleMove) + 1
                    Case "Multi"
                        mCostedSegments(MultiMove) = mCostedSegments(MultiMove) + 1
                    Case "Unit"
                        mCostedSegments(UnitMove) = mCostedSegments(UnitMove) + 1
                End Select

                'Increment the total number of Segments Costed
                mNumOfCostedSegments = mNumOfCostedSegments + 1

                ' Extract the calculated values with the efficiency adjustment and save them to CSV files

                If chk100Series.Checked = True Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get100Series(mP3BatchOutputSheet))
                    sb100Series_EffAdj.WriteLine(mOutString)
                End If

                If chk200Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get200Series(mP3BatchOutputSheet))
                    sb200Series_EffAdj.WriteLine(mOutString)
                End If

                If chk300Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get300Series(mP3BatchOutputSheet))
                    sb300Series_EffAdj.WriteLine(mOutString)
                End If

                If chk400Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get400Series(mP3BatchOutputSheet))
                    sb400Series_EffAdj.WriteLine(mOutString)
                End If

                If chk500Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get500Series(mP3BatchOutputSheet))
                    sb500Series_EffAdj.WriteLine(mOutString)
                End If

                If chk600Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get600Series(mP3BatchOutputSheet))
                    sb600Series_EffAdj.WriteLine(mOutString)
                End If

                If chkSaveResults.Checked Then
                    mP3CellRange = mP3BatchCostProgramSheet.Cells
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_No") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_No") & ",")
                    For Each cell As IRange In mP3CellRange("Q5:Y5")
                        mVal = cell.Value
                        mOutString.Append(mVal.ToString & ",")
                    Next cell

                    ' Switch to the Batch Output sheet
                    mP3CellRange = mP3BatchOutputSheet.Cells

                    ' Get the results for L569 – the Intermodal Plan
                    mOutString.Append(mP3CellRange("KZ1").Value & ",")

                    ' Get the results for L587 – the total Make-Whole Costs
                    mOutString.Append(mP3CellRange("LR1").Value & ",")

                    ' Get the results for L696 – the total costs less the loss & damage costs
                    mOutString.Append(mP3CellRange("PJ1").Value & ",")

                    ' Get the results for L699 – the loss & damage costs
                    mOutString.Append(mP3CellRange("PM1").Value & ",")

                    ' Get the results for L700
                    mOutString.Append(mP3CellRange("PN1").Value & ",")

                    ' Get the Exp_Factor_Th field
                    mOutString.Append(mDataTable.Rows(mLooper)("Exp_Factor_Th") & ",")

                    ' Get the ID
                    mOutString.Append(mP3CellRange("A5").Value)

                    sbResults_EffAdj.WriteLine(mOutString.ToString)

                End If

                ' Copy the BatchOutout data to p3 to the MWF Movement sheet, line 5
                mP3CellRange = mP3BatchOutputSheet.Cells("A5:PQ5")
                mP3CellRange.Copy(mMWFMovement.Cells("A5"))

                ' Set the Efficiancy Adjustment option to N
                mP3DetailedParametersSheet.Cells("C90").Value = "N"

                'Calculate the costs for this segment without Efficiency Adjustment
                mP3Workbook.WorkbookSet.Calculate()

                ' Copy P5:Y5 on the BatchcostProgram sheet to the BatchOutput sheet
                mP3CellRange = mP3BatchCostProgramSheet.Cells("P5:Y5")
                mP3CellRange.Copy(mP3BatchOutputSheet.Cells("A5"), PasteType.Values, PasteOperation.None, False, False)

                ' Copy K1:PQ1 on the BatchOutput sheet to K5:PQ5 on the same sheet
                mP3CellRange = mP3BatchOutputSheet.Cells("K1:PQ1")
                mP3CellRange.Copy(mP3BatchOutputSheet.Cells("K5"), PasteType.Values, PasteOperation.None, False, False)

                ' Extract the calculated values without the efficiency adjustment and save them to CSV files

                If chk100Series.Checked = True Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get100Series(mP3BatchOutputSheet))
                    sb100Series_SysAvg.WriteLine(mOutString)
                End If

                If chk200Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get200Series(mP3BatchOutputSheet))
                    sb200Series_SysAvg.WriteLine(mOutString)
                End If

                If chk300Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get300Series(mP3BatchOutputSheet))
                    sb300Series_SysAvg.WriteLine(mOutString)
                End If

                If chk400Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get400Series(mP3BatchOutputSheet))
                    sb400Series_SysAvg.WriteLine(mOutString)
                End If

                If chk500Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get500Series(mP3BatchOutputSheet))
                    sb500Series_SysAvg.WriteLine(mOutString)
                End If

                If chk600Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get600Series(mP3BatchOutputSheet))
                    sb600Series_SysAvg.WriteLine(mOutString)
                End If

                If chkSaveResults.Checked Then
                    mP3CellRange = mP3BatchCostProgramSheet.Cells
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_No") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_No") & ",")
                    For Each cell As IRange In mP3CellRange("Q5:Y5")
                        mVal = cell.Value
                        mOutString.Append(mVal.ToString & ",")
                    Next cell

                    ' Switch to the Batch Output sheet
                    mP3CellRange = mP3BatchOutputSheet.Cells

                    'get the results for L569
                    mOutString.Append(mP3CellRange("KZ1").Value & ",")
                    'get the results for L587
                    mOutString.Append(mP3CellRange("LR1").Value & ",")
                    'get the results for L696
                    mOutString.Append(mP3CellRange("PJ1").Value & ",")
                    'get the results for L699
                    mOutString.Append(mP3CellRange("PM1").Value & ",")
                    ' Get the results for L700
                    mOutString.Append(mP3CellRange("PN1").Value & ",")
                    'get the results for Exp_Factor_Th
                    mOutString.Append(mDataTable.Rows(mLooper)("Exp_Factor_Th") & ",")
                    ' Get the results for the ID
                    mOutString.Append(mP3CellRange("A5").Value)

                    sbResults_SysAvg.WriteLine(mOutString.ToString)
                End If

                ' Copy the BatchOutout data to p3 to the MWF Movement sheet, line 6
                mP3CellRange = mP3BatchOutputSheet.Cells("A5:PQ5")
                mP3CellRange.Copy(mMWFMovement.Cells("A6"), PasteType.Values, PasteOperation.None, False, False)

                ' Set the expansion factor into the sheet
                mMWFMovement.Cells("B8").Value = mExp_Factor_Th

                ' Force the MakeWhole sheet to calculate
                mMWFWorkbook.WorkbookSet.Calculate()

                ' Save the values from MWF CRPRESRecord sheet
                If chk_Save_CRPRESRecords.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_No") & (0.1 * mDataTable.Rows(mLooper)("Seg_No")).ToString & ",")
                    mOutString.Append(mMWFCRPRESRecord.Cells("B5").Value & ",")
                    For mLooper2 = 2 To 18
                        mOutString.Append(mMWFCRPRESRecord.Cells(Chr(65 + mLooper2) & "5").Value & ",")
                    Next
                    mOutString.Append(mMWFCRPRESRecord.Cells("T5").Value)
                    sbCRPRESRecord.WriteLine(mOutString.ToString)
                End If

                For mLooper2 = 2 To 19
                    mCRPRESArray(mMWFCRPRESRecord.Cells("B5").Value, mLooper2) =
                        mCRPRESArray(mMWFCRPRESRecord.Cells("B5").Value, mLooper2) + mMWFCRPRESRecord.Cells(Chr(65 + mLooper2) & "5").Value
                Next

                ' Update the mCostBreakdown Summary Matrix

                ' Column 1 is Exp_Factor_Th * (L696 – L587 + L699) of row 6 in the CRPRESRecord sheet
                mCostBreakdown(mMWFCRPRESRecord.Cells("B5").Value, 1) = mCostBreakdown(mMWFCRPRESRecord.Cells("B5").Value, 1) +
                    mMWFMovement.Cells("B8").Value * (mMWFMovement.Cells("PJ6").Value - mMWFMovement.Cells("LR6").Value +
                    mMWFMovement.Cells("PM6").Value)

                ' Column 2 is Exp_Factor_Th * (L696 – L587 + L699) of row 5 in the CRPRESRecord sheet
                mCostBreakdown(mMWFCRPRESRecord.Cells("B5").Value, 2) = mCostBreakdown(mMWFCRPRESRecord.Cells("B5").Value, 2) +
                    mMWFMovement.Cells("B8").Value * (mMWFMovement.Cells("PJ5").Value - mMWFMovement.Cells("LR5").Value +
                    mMWFMovement.Cells("PM5").Value)

                ' Column 3 is the sum of C2-C11 of the CRPRESRecord sheet
                mCostBreakdownCol3 = 0
                For mLooper2 = 2 To 11
                    mCostBreakdownCol3 = mCostBreakdownCol3 + mMWFCRPRESRecord.Cells(Chr(65 + mLooper2) & "5").Value
                Next

                mCostBreakdown(mMWFCRPRESRecord.Cells("B5").Value, 3) = mCostBreakdown(mMWFCRPRESRecord.Cells("B5").Value, 3) + mCostBreakdownCol3

                If mTestOne = True Then
                    Exit For
                End If

            End If
        Next

        mPass2Starts = Now

        sbLogFile.WriteLine(" ")
        mOutString = New StringBuilder
        mOutString.Append("*** First pass thru Segments table to calculate the Make-Whole Factors completed at " & DateTime.Now.ToString("G") &
                          ". " & Return_Elapsed_Time(mPass1Starts, mPass2Starts))
        sbLogFile.WriteLine(mOutString.ToString)

        mOutString = New StringBuilder
        mTS = New TimeSpan
        mTS = mPass2Starts - mPass1Starts
        mOutString.Append("Processed approx. " & Math.Round(mDataTable.Rows.Count / mTS.TotalSeconds, 0, MidpointRounding.AwayFromZero) & " Segments per second.")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        If chk100Series.Checked Then
            sb100Series_SysAvg.Flush()
            sb100Series_SysAvg.Close()
            sb100Series_EffAdj.Flush()
            sb100Series_EffAdj.Close()
        End If

        If chk200Series.Checked Then
            sb200Series_SysAvg.Flush()
            sb200Series_SysAvg.Close()
            sb200Series_EffAdj.Flush()
            sb200Series_EffAdj.Close()
        End If

        If chk300Series.Checked Then
            sb300Series_SysAvg.Flush()
            sb300Series_SysAvg.Close()
            sb300Series_EffAdj.Flush()
            sb300Series_EffAdj.Close()
        End If

        If chk400Series.Checked Then
            sb400Series_SysAvg.Flush()
            sb400Series_SysAvg.Close()
            sb400Series_EffAdj.Flush()
            sb400Series_EffAdj.Close()
        End If

        If chk500Series.Checked Then
            sb500Series_SysAvg.Flush()
            sb500Series_SysAvg.Close()
            sb500Series_EffAdj.Flush()
            sb500Series_EffAdj.Close()
        End If

        If chk600Series.Checked Then
            sb600Series_SysAvg.Flush()
            sb600Series_SysAvg.Close()
            sb600Series_EffAdj.Flush()
            sb600Series_EffAdj.Close()
        End If

        If chkSaveResults.Checked Then
            sbResults_SysAvg.Flush()
            sbResults_SysAvg.Close()
            sbResults_EffAdj.Flush()
            sbResults_EffAdj.Close()
        End If

        ' Write the CRPRES array to the file, if desired
        If chk_Save_CRPRESRecords.Checked Then
            For mLooper = 1 To 9
                mOutString = New StringBuilder
                mOutString.Append(mLooper.ToString & ",")
                For mLooper2 = 2 To 18
                    mOutString.Append(mCRPRESArray(mLooper, mLooper2).ToString & ",")
                Next
                mOutString.Append(mCRPRESArray(mLooper, 19).ToString)
                sbCRPRES.WriteLine(mOutString)
            Next

            sbCRPRES.Flush()
            sbCRPRES.Close()

        End If

        mEValuesUpdateStarts = Now()

        mOutString = New StringBuilder
        mOutString.Append("*** Update of Make-Whole Factors in Phase III Spreadsheet Started at " & mEValuesUpdateStarts.ToString("G"))
        sbLogFile.WriteLine()
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        ' Save the CRPRES array to the spreadsheet
        ' for each row...
        For mLooper = 1 To 9
            For mLooper2 = 66 To 83     ' for columns B thru S
                mMWFCRPRES.Cells(Chr(mLooper2) & (mLooper + 4).ToString).Value = mCRPRESArray(mLooper, mLooper2 - 64)
            Next
        Next

        ' Force the sheet to calculate
        mMWFWorkbook.WorkbookSet.Calculate()

        ' Save the MWF workbook
        mMWFWorkbookSet.Calculation = Calculation.Automatic
        mMWFWorkbook.Save()

        mOutString = New StringBuilder
        mOutString.Append("Update of Make-Whole Factors in Phase III Spreadsheet completed at " & Now.ToString("G") & ". " & Return_Elapsed_Time(mEValuesUpdateStarts, Now()))
        sbLogFile.WriteLine()
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        mPass3Starts = Now()

        mOutString = New StringBuilder
        mOutString.Append("*** Update of Make-Whole Factors in the EValues table started at " & mEValuesUpdateStarts.ToString("G"))
        sbLogFile.WriteLine()
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        ' Switch to the XMLValues sheet and save the values to the EValues table in SQL
        ' Write the value to the RailroadCostXML sheet as well
        For mLooper = 5 To 94

            ' Make sure we're on the right database
            gbl_Database_Name = Get_Database_Name_From_SQL(cmb_URCS_Year.Text, "EVALUES")
            OpenSQLConnection(gbl_Database_Name)

            ' build a SQL command using columns B, D, and F, adding an entry_dt value
            mStrSQL = "UPDATE " & Gbl_EValues_TableName & " SET " &
                        "VALUE = " & mMWFXMLValues.Cells("F" & mLooper.ToString).Value & ", " &
                        "ENTRY_DT = '" & DateTime.Now.ToString("G") & "' " &
                        "WHERE YEAR = " & cmb_URCS_Year.Text.ToString & " AND " &
                        "RR_ID = " & mMWFXMLValues.Cells("B" & mLooper.ToString).Value & " AND " &
                        "ECODE_ID = " & mMWFXMLValues.Cells("D" & mLooper.ToString).Value
            mSQLcmd = New SqlCommand
            mSQLcmd.Connection = gbl_SQLConnection
            mSQLcmd.CommandType = CommandType.Text
            mSQLcmd.CommandText = mStrSQL
            mSQLcmd.ExecuteNonQuery()

            ' Build the vlookup value to search for
            mVLookup = Get_Short_Name_By_RRID(mMWFXMLValues.Cells("B" & mLooper.ToString).Value) & "-"
            mVLookup = mVLookup & mMWFXMLValues.Cells("C" & mLooper.ToString).Value

            mP3CellRange = mP3RailroadUnitCostXMLSheet.Cells.Find(mVLookup,
                                                              mP3RailroadUnitCostXMLSheet.Cells("A1:A8758"),
                                                              FindLookIn.Values,
                                                              LookAt.Whole,
                                                              SearchOrder.ByRows,
                                                              SearchDirection.Next,
                                                              False)
            If IsNothing(mP3CellRange) Then
                ' The column doesn't exist
                ' For now we'll report this as an error
                ' In the future, we'll insert the column
                MsgBox("Vlookup value not found in P3XL spreadsheet!", vbOKOnly, "ERROR!")
                GoTo EndIt
            Else
                ' We found the vlookup cell, so we'll use the range row value and the target column to update the sheet
                mP3RailroadUnitCostXMLSheet.Cells(mXMLSheetColumn & (mP3CellRange.Row + 1).ToString).Value =
                    mMWFXMLValues.Cells("F" & mLooper.ToString).Value
            End If

        Next

        mOutString = New StringBuilder
        mOutString.Append("*** Update of Make-Whole Factors in EValues table completed at " & Now.ToString("G") & " " & Return_Elapsed_Time(mEValuesUpdateStarts, Now()))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        'Switch the connection to the Waybills database
        OpenSQLConnection("Waybills")

        mOutString = New StringBuilder
        mOutString.Append("*** Waybill Processing Statistics")
        sbLogFile.WriteLine()
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append("Number of Masked Waybill Records Read:")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mNumOfWaybillRecords.ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append("Number of Masked Segments Read:")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mNumOfMaskedSegments.ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        sbLogFile.WriteLine()

        mOutString = New StringBuilder
        mOutString.Append("Number of Costed Waybills Written:")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mNumOfCostedWaybills.ToString("N0"), 7))
        mOutString.Append(" (" & ((mNumOfCostedWaybills / mNumOfWaybillRecords) * 100).ToString("N2") &
                            "% of Total)")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append("Number of Costed Segments Written:")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mNumOfCostedSegments.ToString("N0"), 7))
        mOutString.Append(" (" & ((mNumOfCostedSegments / mNumOfMaskedSegments) * 100).ToString("N2") &
                            "% of Total)")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        sbLogFile.WriteLine()

        mOutString = New StringBuilder
        mOutString.Append("Railroads")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append("Segments")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        mNumOfCostedSegments = 0

        For mLooper2 = 1 To 9
            mOutString = New StringBuilder
            mOutString.Append(Get_Short_Name_By_RRID(mLooper2))
            mOutString.Append(Space(40 - mOutString.Length))
            mOutString.Append(Right_Justify(mCostedSegmentsByRailroad(mLooper2).ToString("N0"), 7))
            mNumOfCostedSegments = mNumOfCostedSegments + mCostedSegmentsByRailroad(mLooper2)
            sbLogFile.WriteLine(mOutString.ToString)
            sbLogFile.Flush()
        Next

        mOutString = New StringBuilder
        mOutString.Append("Total")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mNumOfCostedSegments.ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append("Rejections:")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append("Segments")
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append("Waybills")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        For mLooper2 = 1 To 7
            mOutString = New StringBuilder

            Select Case mLooper2
                Case 1
                    mOutString.Append("Bad Car Type")
                Case 2
                    mOutString.Append("Bad Tonnage")
                Case 3
                    mOutString.Append("Bad Revenue")
                Case 4
                    mOutString.Append("Bad Distance")
                Case 5
                    mOutString.Append("Bad Intermodal Specification")
                Case 6
                    mOutString.Append("Non US Segments")
                Case 7
                    mOutString.Append("Negative Cost")
            End Select

            mOutString.Append(Space(40 - mOutString.Length))
            mOutString.Append(Right_Justify(mRejectedSegments(mLooper2).ToString("N0"), 7))
            mOutString.Append(Space(52 - mOutString.Length))
            mOutString.Append(Right_Justify(mRejectedWaybills(mLooper2).ToString("N0"), 7))
            sbLogFile.WriteLine(mOutString.ToString)
            sbLogFile.Flush()
        Next

        mOutString = New StringBuilder
        mOutString.Append("Total Rejected Segments")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify((mRejectedSegments(BadCarType) +
            mRejectedSegments(BadTonnage) +
            mRejectedSegments(BadRevenue) +
            mRejectedSegments(BadDistance) +
            mRejectedSegments(BadIntermodalSpec) +
            mRejectedSegments(NonUSSegment) +
            mRejectedSegments(NegativeCost)).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify((mRejectedWaybills(BadCarType) +
            mRejectedWaybills(BadTonnage) +
            mRejectedWaybills(BadRevenue) +
            mRejectedWaybills(BadDistance) +
            mRejectedWaybills(BadIntermodalSpec) +
            mRejectedWaybills(NonUSSegment) +
            mRejectedWaybills(NegativeCost)).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        ' Header Line
        mOutString = New StringBuilder
        mOutString.Append("Costed As")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append("Segments")
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append("Waybills")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        ' Single Carload Data Line
        mOutString = New StringBuilder
        mOutString.Append("Single Carload")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedSegments(SingleMove).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedWaybills(SingleMove).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        'Multiple Carload Data Line
        mOutString = New StringBuilder
        mOutString.Append("Multiple Carload")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedSegments(MultiMove).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedWaybills(MultiMove).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        ' Unit Train Data line
        mOutString = New StringBuilder
        mOutString.Append("Unit Train")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedSegments(UnitMove).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedWaybills(UnitMove).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        ' Intermodal Data Line
        mOutString = New StringBuilder
        mOutString.Append("Intermodal")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedSegments(IntermodalMove).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedWaybills(IntermodalMove).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        ' Total Costed Data line
        mOutString = New StringBuilder
        mOutString.Append("Total Costed")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify((mCostedSegments(SingleMove) +
                          mCostedSegments(MultiMove) +
                          mCostedSegments(UnitMove) +
                          mCostedSegments(IntermodalMove)).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify((mCostedWaybills(SingleMove) +
                          mCostedWaybills(MultiMove) +
                          mCostedWaybills(UnitMove) +
                          mCostedWaybills(IntermodalMove)).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        ' Write out the results of the Variable Cost Breakdown
        mOutString = New StringBuilder
        mOutString.Append("Variable Cost Breakdown")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append(Space(12) & "System")
        mOutString.Append(Space(25 - mOutString.Length))
        mOutString.Append("Efficiency")
        mOutString.Append(Space(58 - mOutString.Length))
        mOutString.Append("Total")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append(Space(12) & "Average")
        mOutString.Append(Space(26 - mOutString.Length))
        mOutString.Append("Adjusted")
        mOutString.Append(Space(56 - mOutString.Length))
        mOutString.Append("Adjusted")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append(Space(13) & "Costs")
        mOutString.Append(Space(28 - mOutString.Length))
        mOutString.Append("Costs")
        mOutString.Append(Space(41 - mOutString.Length))
        mOutString.Append("Make Whole")
        mOutString.Append(Space(58 - mOutString.Length))
        mOutString.Append("Costs")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        For mLooper = 1 To 9
            mOutString = New StringBuilder
            mOutString.Append(Get_Short_Name_By_RRID(mLooper))
            mOutString.Append(Space(6 - mOutString.Length))
            mOutString.Append(Right_Justify(mCostBreakdown(mLooper, 1).ToString("N0"), 15))
            mCostBreakdownCol1Total = mCostBreakdownCol1Total + mCostBreakdown(mLooper, 1)

            mOutString.Append(Right_Justify(mCostBreakdown(mLooper, 2).ToString("N0"), 15))
            mCostBreakdownCol2Total = mCostBreakdownCol2Total + mCostBreakdown(mLooper, 2)

            mOutString.Append(Right_Justify(mCostBreakdown(mLooper, 3).ToString("N0"), 15))
            mCostBreakdownCol3Total = mCostBreakdownCol3Total + mCostBreakdown(mLooper, 3)

            mOutString.Append(Right_Justify((mCostBreakdown(mLooper, 2) + mCostBreakdown(mLooper, 3)).ToString("N0"), 15))
            mCostBreakdownCol4Total = mCostBreakdownCol4Total + mCostBreakdown(mLooper, 2) + mCostBreakdown(mLooper, 3)
            sbLogFile.WriteLine(mOutString.ToString)
            sbLogFile.Flush()
        Next

        sbLogFile.WriteLine()
        mOutString = New StringBuilder
        mOutString.Append("Total")
        mOutString.Append(Space(6 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostBreakdownCol1Total.ToString("N0"), 15))
        mOutString.Append(Right_Justify(mCostBreakdownCol2Total.ToString("N0"), 15))
        mOutString.Append(Right_Justify(mCostBreakdownCol3Total.ToString("N0"), 15))
        mOutString.Append(Right_Justify(mCostBreakdownCol4Total.ToString("N0"), 15))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        mCounter = 0
        mPass2Starts = Now

        mOutString = New StringBuilder
        mOutString.Append("*** Second pass thru Segments table to calculate costs including the Make-Whole Factors started at  " & DateTime.Now.ToString("G"))
        sbLogFile.WriteLine()
        sbLogFile.WriteLine(mOutString.ToString)

        mNumOfWaybillRecords = 0
        mNumOfCostedWaybills = 0
        mNumOfCostedRecords = 0
        mNumOfCostedSegments = 0
        mNumOfMaskedSegments = mDataTable.Rows.Count
        mNumOfRejectedSegments = 0
        mSerial_No = 0

        For mLooper = 1 To 9
            mCostedSegmentsByRailroad(mLooper) = 0
            mTotalVariableCosts(mLooper) = 0
        Next

        For mLooper = 1 To 7
            mRejectedSegments(mLooper) = 0
            mRejectedWaybills(mLooper) = 0
        Next

        For mLooper = 1 To 4
            mCostedSegments(mLooper) = 0
            mCostedWaybills(mLooper) = 0
        Next

        'Now we're ready to run through the datatable once more to get the L700 value for each segment
        For mLooper = 0 To mDataTable.Rows.Count - 1

            If mSerial_No <> mDataTable.Rows(mLooper)("Serial_No") Then
                mNumOfWaybillRecords = mNumOfWaybillRecords + 1
                mSerialChanged = True
            Else
                mSerialChanged = False
            End If

            mCounter = mCounter + 1
            If mCounter Mod 100 = 0 Then
                txtStatus.Text = "Pass 2 of 3 - Updating Segment record VC " & mCounter.ToString & " of " _
                    & mDataTable.Rows.Count.ToString & " - " & ((mLooper / mDataTable.Rows.Count) * 100).ToString("N1") & "%..."
                Refresh()
                Application.DoEvents()
            End If

            If mCounter Mod 100000 = 0 Then
                sbLogFile.WriteLine(“*** Processing record " & mCounter.ToString & " of the Segments table at " & DateTime.Now.ToString("G"))
            End If

            'Clean up any old data
            mP3CellRange = mP3RailroadCostProgramSheet.Cells("D6:D11")
            mP3CellRange.Value = ""
            mP3CellRange = mP3RailroadCostProgramSheet.Cells("E6:H14")
            mP3CellRange.Value = ""

            ' Set the Efficiency Adjustment option to "Y" for the first calculation
            mP3DetailedParametersSheet.Cells("C90").Value = "Y"

            ' Save this for the MakeWhole Factor Movement sheet
            mExp_Factor_Th = mDataTable.Rows(mLooper)("Exp_Factor_Th")

            'Start laying in the values to line 5 base zero of the spreadsheet
            mSerial_No = mDataTable.Rows(mLooper)("Serial_no")
            mSeg_No = mDataTable.Rows(mLooper)("Seg_no")

            ' Input Parameter: ID
            mP3BatchCostProgramSheet.Cells("A5").Value = mDataTable.Rows(mLooper)("Serial_No") & mDataTable.Rows(mLooper)("Seg_No").ToString

            ' Input Parameter: RR
            mP3BatchCostProgramSheet.Cells("B5").Value = mDataTable.Rows(mLooper)("RR_Num")

            ' Input Parameter: DIS
            mP3BatchCostProgramSheet.Cells("C5").Value = mDataTable.Rows(mLooper)("RR_Dist") * 0.1

            ' Input Parameter: SG
            mP3BatchCostProgramSheet.Cells("D5").Value = mDataTable.Rows(mLooper)("Seg_Type")

            ' Input Parameter: FC
            If rdo_Legacy.Checked = True Then
                ' Logic if running as Legacy 
                mTOFCMove = False
                Select Case mDataTable.Rows(mLooper)("STB_Car_Typ")
                    Case 46, 49, 52, 54
                        If (mDataTable.Rows(mLooper)("TOFC_Serv_Code") <> "") Or (mDataTable.Rows(mLooper)("Int_Eq_Flg") = 2) Then
                            mTOFCMove = True
                            mP3BatchCostProgramSheet.Cells("E5").Value = 46
                        Else
                            mP3BatchCostProgramSheet.Cells("E5").Value = mDataTable.Rows(mLooper)("STB_Car_typ")
                        End If
                    Case Else
                        mP3BatchCostProgramSheet.Cells("E5").Value = mDataTable.Rows(mLooper)("STB_Car_typ")
                End Select
            Else
                ' Check to see if we have a value for the Car Type.  If not, use default
                If String.IsNullOrEmpty(mDataTable.Rows(mLooper)("STB_Car_typ")) Then
                    mP3BatchCostProgramSheet.Cells("E5").Value = 52
                Else
                    mP3BatchCostProgramSheet.Cells("E5").Value = mDataTable.Rows(mLooper)("STB_Car_typ")
                End If

                mTOFCMove = False
                Select Case mDataTable.Rows(mLooper)("STB_Car_Typ")
                    Case 46, 48, 49, 52, 54
                        If (mDataTable.Rows(mLooper)("TOFC_Serv_Code") <> "") Or (mDataTable.Rows(mLooper)("Int_Eq_Flg") = 2) Then
                            ' If RoadRailer (Int_Eq_Flg = 2) set the cartyp to TOFC flat Car (46)
                            If mDataTable.Rows(mLooper)("Int_Eq_Flg") = 2 Then
                                mP3BatchCostProgramSheet.Cells("E5").Value = 46
                            End If
                            mTOFCMove = True
                        End If
                End Select
            End If

            ' Input Parameter: NC
            'Check for the larger of U_Cars or U_TC_Units
            If mDataTable.Rows(mLooper)("U_Cars") > mDataTable.Rows(mLooper)("U_TC_Units") Then
                mP3BatchCostProgramSheet.Cells("F5").Value = mDataTable.Rows(mLooper)("U_Cars")
            Else
                mP3BatchCostProgramSheet.Cells("F5").Value = mDataTable.Rows(mLooper)("U_TC_Units")
            End If

            ' Input Parameter: OWN
            If rdo_Legacy.Checked = True Then
                ' Logic if running as Legacy
                Select Case Trim(mDataTable.Rows(mLooper)("U_Car_init"))
                    Case "ABOX", "RBOX", "CSX", "CSXT", "GONX"
                        mP3BatchCostProgramSheet.Cells("G5").Value = "R"
                    Case Else
                        If Strings.Right(Trim(mDataTable.Rows(mLooper)("U_Car_Init")), 1) = "X" Then
                            mP3BatchCostProgramSheet.Cells("G5").Value = "P"
                        Else
                            mP3BatchCostProgramSheet.Cells("G5").Value = "R"
                        End If
                End Select
            Else
                Select Case CInt(cmb_URCS_Year.Text)
                    Case 2015, 2016, 2017
                        If Strings.Right(Trim(mDataTable.Rows(mLooper)("U_Car_Init")), 1) = "X" Then
                            mP3BatchCostProgramSheet.Cells("G5").Value = "P"
                        Else
                            mP3BatchCostProgramSheet.Cells("G5").Value = "R"
                        End If
                    Case Else
                        ' Use value in Car_Own field unless that is blank
                        Select Case Trim(mDataTable.Rows(mLooper)("Car_Own"))
                            Case "R", "P", "T"
                                mP3BatchCostProgramSheet.Cells("G5").Value = mDataTable.Rows(mLooper)("Car_Own")
                            Case Else
                                If Strings.Right(Trim(mDataTable.Rows(mLooper)("U_Car_Init")), 1) = "X" Then
                                    mP3BatchCostProgramSheet.Cells("G5").Value = "P"
                                Else
                                    mP3BatchCostProgramSheet.Cells("G5").Value = "R"
                                End If
                        End Select
                End Select
            End If

            ' Input Parameter: WT
            ' Calculate the tons per car or tons per TCU value
            If mTOFCMove Then
                If mDataTable.Rows(mLooper)("U_TC_Units") > 0 Then
                    mP3BatchCostProgramSheet.Cells("H5").Value = mDataTable.Rows(mLooper)("Bill_Wght_Tons") / mDataTable.Rows(mLooper)("U_TC_Units")
                Else
                    mP3BatchCostProgramSheet.Cells("H5").Value = 0
                End If
            Else
                If mDataTable.Rows(mLooper)("U_Cars") > 0 Then
                    mP3BatchCostProgramSheet.Cells("H5").Value = mDataTable.Rows(mLooper)("Bill_Wght_Tons") / mDataTable.Rows(mLooper)("U_Cars")
                Else
                    mP3BatchCostProgramSheet.Cells("H5").Value = 0
                End If
            End If

            ' Input Parameter: COM
            mP3BatchCostProgramSheet.Cells("I5").Value = "'" & Strings.Left(mDataTable.Rows(mLooper)("STCC"), 5)

            ' Input Parameter: SZ (Set the shipment size - legacy v. EP431 Sub 4)
            If mTOFCMove = True Then
                mP3BatchCostProgramSheet.Cells("J5").Value = "Intermodal"
            ElseIf mDataTable.Rows(mLooper)("U_Cars") <= 5 Then
                mP3BatchCostProgramSheet.Cells("J5").Value = "Single"
            ElseIf mDataTable.Rows(mLooper)("U_Cars") < mUnitTrainDef Then
                mP3BatchCostProgramSheet.Cells("J5").Value = "Multi"
            Else
                mP3BatchCostProgramSheet.Cells("J5").Value = "Unit"
            End If

            ' Input Parameter: L102 (Set the Circuity to 1)
            mP3BatchCostProgramSheet.Cells("K5").Value = 1

            ' Input Parameter: L569
            If mDataTable.Rows(mLooper)("Int_Eq_Flg") = 2 Then
                mP3BatchCostProgramSheet.Cells("L5").Value = "TCS"
            Else
                mP3BatchCostProgramSheet.Cells("L5").Value = Trim(mDataTable.Rows(mLooper)("TOFC_Serv_Code"))
            End If

            ' Input Parameter: TDIS
            mP3BatchCostProgramSheet.Cells("M5").Value = mDataTable.Rows(mLooper)("Total_Dist") * 0.1

            ' Input Parameter: ORR
            mP3BatchCostProgramSheet.Cells("N5").Value = mDataTable.Rows(mLooper)("ORR")

            ' Copy the formulas for P5:AC5
            With mP3BatchCostProgramSheet
                mP3CellRange = .Cells
                mP3CellRange("P1:AC1").Copy(.Cells("P5:AC5"))
            End With

            ' Implement Current Waybill Logic that excludes certain records from being costed
            mProcess = True

            ' Skip record if Freight Car (FC) equals 0
            If mP3BatchCostProgramSheet.Cells("E5").Value = 0 Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadCarType) = mRejectedWaybills(BadCarType) + 1
                End If
                mRejectedSegments(BadCarType) = mRejectedSegments(BadCarType) + 1
            End If

            ' Skip record if Weight (WT) is less than 1 ton
            If mP3BatchCostProgramSheet.Cells("H5").Value < 1 Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadTonnage) = mRejectedWaybills(BadTonnage) + 1
                End If
                mRejectedSegments(BadTonnage) = mRejectedSegments(BadTonnage) + 1
            End If

            ' Skip record if Total Revenue is zero
            If mDataTable.Rows(mLooper)("total_rev") = 0 Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadRevenue) = mRejectedWaybills(BadRevenue) + 1
                End If
                mRejectedSegments(BadRevenue) = mRejectedSegments(BadRevenue) + 1
            End If

            ' Skip record if 1 railroad (Total_Segs=1) and TDIS < 1.5 (rounded)
            If (mDataTable.Rows(mLooper)("total_Segs") = 1) And mP3BatchCostProgramSheet.Cells("M5").Value < 1.5 Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (mDataTable.Rows(mLooper)("total_Segs") > 1) And mP3BatchCostProgramSheet.Cells("M5").Value < 4.5 Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("ORR_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 0) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = False) And ((mDataTable.Rows(mLooper)("ORR_Dist") * 0.1) < 0.5) And
                    (mDataTable.Rows(mLooper)("JF") >= 0) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("TRR_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 1) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = False) And ((mDataTable.Rows(mLooper)("TRR_Dist") * 0.1) < 0.5) And
                (mDataTable.Rows(mLooper)("JF") >= 1) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR1_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 2) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = False) And ((mDataTable.Rows(mLooper)("JRR1_Dist") * 0.1) < 0.5) And
                (mDataTable.Rows(mLooper)("JF") >= 2) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR2_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 3) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = False) And
                   ((mDataTable.Rows(mLooper)("JRR2_Dist") * 0.1) < 0.5) And
                   (mDataTable.Rows(mLooper)("JF") >= 3) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR3_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 4) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = False) And
                   ((mDataTable.Rows(mLooper)("JRR3_Dist") * 0.1) < 0.5) And
                   (mDataTable.Rows(mLooper)("JF") >= 4) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR4_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 5) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = False) And
                ((mDataTable.Rows(mLooper)("JRR4_Dist") * 0.1) < 0.5) And (mDataTable.Rows(mLooper)("JF") >= 5) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR5_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 6) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = False) And
                   ((mDataTable.Rows(mLooper)("JRR5_Dist") * 0.1) < 0.5) And
                   (mDataTable.Rows(mLooper)("JF") >= 6) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = True) And
                   (Math.Round(mDataTable.Rows(mLooper)("JRR6_Dist") * 0.1, 0, MidpointRounding.AwayFromZero) = 0) And
                   (mDataTable.Rows(mLooper)("JF") >= 7) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            ElseIf (rdo_Legacy.Checked = False) And
                   ((mDataTable.Rows(mLooper)("JRR6_Dist") * 0.1) < 0.5) And
                   (mDataTable.Rows(mLooper)("JF") >= 7) Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadDistance) = mRejectedWaybills(BadDistance) + 1
                End If
                mRejectedSegments(BadDistance) = mRejectedSegments(BadDistance) + 1
            End If

            ' Skip record if FC = 46 and SZ not Intermodal
            If (mP3BatchCostProgramSheet.Cells("E5").Value = "46") And (mP3BatchCostProgramSheet.Cells("J5").Value <> "Intermodal") Then
                mProcess = False
                If mSerialChanged = True Then
                    mRejectedWaybills(BadIntermodalSpec) = mRejectedWaybills(BadIntermodalSpec) + 1
                End If
                mRejectedSegments(BadIntermodalSpec) = mRejectedSegments(BadIntermodalSpec) + 1
            End If

            If rdo_Legacy.Checked = False Then  ' Non-Legacy processing

                ' If selected cost non-US segment
                If (chk_Cost_All_Segments.Checked = False) And (mProcess = True) Then
                    If (mDataTable.Rows(mLooper)("RR_Cntry") <> "US") Then
                        mProcess = False
                        If mSerialChanged = True Then
                            mRejectedWaybills(NonUSSegment) = mRejectedWaybills(NonUSSegment) + 1
                        End If
                        mRejectedSegments(NonUSSegment) = mRejectedSegments(NonUSSegment) + 1
                    End If
                End If
            End If

            If mProcess = True Then

                If mPreviousSerialNo <> mSerial_No Then
                    mNumOfCostedWaybills = mNumOfCostedWaybills + 1
                    If mTOFCMove = True Then
                        mCostedWaybills(IntermodalMove) = mCostedWaybills(IntermodalMove) + 1
                    ElseIf mDataTable.Rows(mLooper)("U_Cars") <= 5 Then
                        mCostedWaybills(SingleMove) = mCostedWaybills(SingleMove) + 1
                    ElseIf mDataTable.Rows(mLooper)("U_Cars") < mUnitTrainDef Then
                        mCostedWaybills(MultiMove) = mCostedWaybills(MultiMove) + 1
                    Else
                        mCostedWaybills(UnitMove) = mCostedWaybills(UnitMove) + 1
                    End If
                    mPreviousSerialNo = mSerial_No
                End If

                ' Force the Excel sheet to "translate" the SQL data into user input data
                mP3Workbook.WorkbookSet.Calculate()

                mP3RailroadCostProgramSheet.Cells("D1").Value = cmb_URCS_Year.SelectedItem.ToString                 'Year
                mP3RailroadCostProgramSheet.Cells("D5").Value = mP3BatchCostProgramSheet.Cells("Q5").Value          'RR
                mP3RailroadCostProgramSheet.Cells("D6").Value = mP3BatchCostProgramSheet.Cells("R5").Value          'DIS
                mP3RailroadCostProgramSheet.Cells("D7").Value = mP3BatchCostProgramSheet.Cells("S5").Value          'SG
                mP3RailroadCostProgramSheet.Cells("D8").Value = mP3BatchCostProgramSheet.Cells("T5").Value          'FC
                mP3RailroadCostProgramSheet.Cells("D9").Value = mP3BatchCostProgramSheet.Cells("U5").Value         'NC
                mP3RailroadCostProgramSheet.Cells("D10").Value = mP3BatchCostProgramSheet.Cells("V5").Value         'OWN
                mP3RailroadCostProgramSheet.Cells("D11").Value = mP3BatchCostProgramSheet.Cells("W5").Value         'WT
                mP3RailroadCostProgramSheet.Cells("D12").Value = Replace(mP3BatchCostProgramSheet.Cells("X5").Value, "'", ";")         'COM
                mP3RailroadCostProgramSheet.Cells("D13").Value = mP3BatchCostProgramSheet.Cells("Y5").Value         'SZ

                ' Copy the OtherRRDist (if any) from the BatchCostProgram sheet to the RailroadCostProgram sheet
                mP3RailroadCostProgramSheet.Cells("H6").Value = mP3BatchCostProgramSheet.Cells("Z5").Value

                ' Copy the Circuity (if any) from the BatchCostProgram sheet to the DetailedParameters sheet
                mP3DetailedParametersSheet.Cells("C5").Value = mP3BatchCostProgramSheet.Cells("AA5").Value

                ' Copy the TOFC Plan (if any) from the BatchCostProgram sheet to the DetailedParameters sheet
                ' mP3DetailedParametersSheet.Cells("C14").Value = mP3BatchCostProgramSheet.Cells("L5").Value
                mP3DetailedParametersSheet.Cells("C14").Value = mP3BatchCostProgramSheet.Cells("AB5").Value

                ' Copy the ORR (if any) from the BatchCostProgram sheet to the DetailedParameters sheet
                ' mP3DetailedParametersSheet.Cells("J11").Value = mP3BatchCostProgramSheet.Cells("N5").Value
                mP3DetailedParametersSheet.Cells("J11").Value = mP3BatchCostProgramSheet.Cells("AC5").Value

                'Set the efficiency adjustment flag to Y in the DetailedParameters sheet
                mP3DetailedParametersSheet.Cells("C90").Value = "Y"

                'Calculate the costs for this segment
                mP3Workbook.WorkbookSet.Calculate()

                ' Copy P5:Y5 on the BatchcostProgram sheet to the BatchOutput sheet
                mP3CellRange = mP3BatchCostProgramSheet.Cells("P5:Y5")
                mP3CellRange.Copy(mP3BatchOutputSheet.Cells("A5"), PasteType.Values, PasteOperation.None, False, False)

                ' Copy K1:PQ1 on the BatchOutput sheet to K5:PQ5 on the same sheet
                mP3CellRange = mP3BatchOutputSheet.Cells("K1:PQ1")
                mP3CellRange.Copy(mP3BatchOutputSheet.Cells("K5"), PasteType.Values, PasteOperation.None, False, False)

                ' Produce the cost by multiplying the L700 value by teh Exp_Factor_Th value
                mThisCost = Math.Round(mP3BatchOutputSheet.Cells("PN5").Value * mExp_Factor_Th, 0, MidpointRounding.AwayFromZero)

                If chk_SaveToSQL.Checked = True Then
                    ' Store the cost to the segments record in SQL
                    mStrSQL = "UPDATE " & Gbl_Segments_TableName & " SET " &
                        "RR_VC = " & mThisCost & " " &
                        "WHERE Serial_No = '" & mDataTable.Rows(mLooper)("Serial_No").ToString & "' AND " &
                        "Seg_No = " & mDataTable.Rows(mLooper)("Seg_No")
                    OpenSQLConnection(Gbl_Waybill_Database_Name)
                    mSQLcmd = New SqlCommand
                    mSQLcmd.Connection = gbl_SQLConnection
                    mSQLcmd.CommandType = CommandType.Text
                    mSQLcmd.CommandText = mStrSQL
                    mSQLcmd.ExecuteNonQuery()
                End If

                ' Increment the processed segment counter for this railroad.
                mThisRoad = Get_RRID_By_Short_Name(mP3BatchOutputSheet.Cells("B5").Value)
                mCostedSegmentsByRailroad(mThisRoad) = mCostedSegmentsByRailroad(mThisRoad) + 1

                ' Add this segments cost to the total variable costs 
                mTotalVariableCosts(mThisRoad) = mTotalVariableCosts(mThisRoad) + mThisCost

                'Increment the Segments Costed counters
                Select Case mP3BatchCostProgramSheet.Cells("J5").Value
                    Case "Intermodal"
                        mCostedSegments(IntermodalMove) = mCostedSegments(IntermodalMove) + 1
                    Case "Single"
                        mCostedSegments(SingleMove) = mCostedSegments(SingleMove) + 1
                    Case "Multi"
                        mCostedSegments(MultiMove) = mCostedSegments(MultiMove) + 1
                    Case "Unit"
                        mCostedSegments(UnitMove) = mCostedSegments(UnitMove) + 1
                End Select

                'Increment the total number of Segments Costed
                mNumOfCostedSegments = mNumOfCostedSegments + 1

                If chk100Series.Checked = True Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get100Series(mP3BatchOutputSheet))
                    sb100Series_Costed.WriteLine(mOutString)
                End If

                If chk200Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get200Series(mP3BatchOutputSheet))
                    sb200Series_Costed.WriteLine(mOutString)
                End If

                If chk300Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get300Series(mP3BatchOutputSheet))
                    sb300Series_Costed.WriteLine(mOutString)
                End If

                If chk400Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get400Series(mP3BatchOutputSheet))
                    sb400Series_Costed.WriteLine(mOutString)
                End If

                If chk500Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get500Series(mP3BatchOutputSheet))
                    sb500Series_Costed.WriteLine(mOutString)
                End If

                If chk600Series.Checked Then
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_no") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_no") & ",")
                    mOutString.Append(Get600Series(mP3BatchOutputSheet))
                    sb600Series_Costed.WriteLine(mOutString)
                End If

                If chkSaveResults.Checked Then
                    mP3CellRange = mP3BatchCostProgramSheet.Cells
                    mOutString = New StringBuilder
                    mOutString.Append(mDataTable.Rows(mLooper)("Serial_No") & ",")
                    mOutString.Append(mDataTable.Rows(mLooper)("Seg_No") & ",")
                    For Each cell As IRange In mP3CellRange("Q5:Y5")
                        mVal = cell.Value
                        mOutString.Append(mVal.ToString & ",")
                    Next cell

                    ' Switch to the Batch Output sheet
                    mP3CellRange = mP3BatchOutputSheet.Cells

                    ' Get the results for L569 – the Intermodal Plan
                    mOutString.Append(mP3CellRange("KZ1").Value & ",")

                    ' Get the results for L587 – the total Make-Whole Costs
                    mOutString.Append(mP3CellRange("LR1").Value & ",")

                    ' Get the results for L696 – the total costs less the loss & damage costs
                    mOutString.Append(mP3CellRange("PJ1").Value & ",")

                    ' Get the results for L699 – the loss & damage costs
                    mOutString.Append(mP3CellRange("PM1").Value & ",")

                    ' Get the results for L700
                    mOutString.Append(mP3CellRange("PN1").Value & ",")

                    ' Get the Exp_Factor_Th field
                    mOutString.Append(mDataTable.Rows(mLooper)("Exp_Factor_Th") & ",")

                    ' Get the results for the ID
                    mOutString.Append(mP3CellRange("A5").Value)

                    sbResults_Costed.WriteLine(mOutString.ToString)

                End If
            End If
        Next

        mOutString = New StringBuilder
        mOutString.Append("*** Second pass thru Segments table to calculate costs including the Make-Whole Factors completed at  " & DateTime.Now.ToString("G") & ". " & Return_Elapsed_Time(mPass2Starts, Now()))
        sbLogFile.WriteLine()
        sbLogFile.WriteLine(mOutString.ToString)

        mOutString = New StringBuilder
        mOutString.Append("*** Waybill Processing Statistics")
        sbLogFile.WriteLine()
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append("Number of Masked Waybill Records Read:")
        mOutString.Append(Space(50 - mOutString.Length))
        mOutString.Append(Right_Justify(mNumOfWaybillRecords.ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append("Number of Masked Segments Read:")
        mOutString.Append(Space(50 - mOutString.Length))
        mOutString.Append(Right_Justify(mNumOfMaskedSegments.ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        sbLogFile.WriteLine()

        mOutString = New StringBuilder
        mOutString.Append("Number of Costed Waybills ")
        If chk_SaveToSQL.Checked = True Then
            mOutString.Append("Written:")
        Else
            mOutString.Append("Processed (Not Written):")
        End If
        mOutString.Append(Space(50 - mOutString.Length))
        mOutString.Append(Right_Justify(mNumOfCostedWaybills.ToString("N0"), 7))
        mOutString.Append(" (" & ((mNumOfCostedWaybills / mNumOfWaybillRecords) * 100).ToString("N2") &
                            "% of Total)")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append("Number of Costed Segments")
        If chk_SaveToSQL.Checked = True Then
            mOutString.Append("Written:")
        Else
            mOutString.Append("Processed (Not Written):")
        End If
        mOutString.Append(Space(50 - mOutString.Length))
        mOutString.Append(Right_Justify(mNumOfCostedSegments.ToString("N0"), 7))
        mOutString.Append(" (" & ((mNumOfCostedSegments / mNumOfMaskedSegments) * 100).ToString("N2") &
                            "% of Total)")
        sbLogFile.WriteLine(mOutString.ToString)

        sbLogFile.WriteLine()

        sbLogFile.WriteLine("Number of Segments With Bad Translations: " & mNumOfBadLookups.ToString("N0"))

        sbLogFile.WriteLine()
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append("Railroads")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append("Segments")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        mNumOfCostedSegments = 0

        For mLooper2 = 1 To 9
            mOutString = New StringBuilder
            mOutString.Append(Get_Short_Name_By_RRID(mLooper2))
            mOutString.Append(Space(40 - mOutString.Length))
            mOutString.Append(Right_Justify(mCostedSegmentsByRailroad(mLooper2).ToString("N0"), 7))
            mNumOfCostedSegments = mNumOfCostedSegments + mCostedSegmentsByRailroad(mLooper2)
            sbLogFile.WriteLine(mOutString.ToString)
            sbLogFile.Flush()
        Next

        mOutString = New StringBuilder
        mOutString.Append("Total")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mNumOfCostedSegments.ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        mOutString = New StringBuilder
        mOutString.Append("Rejections:")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append("Segments")
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append("Waybills")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        For mLooper2 = 1 To 7
            mOutString = New StringBuilder

            Select Case mLooper2
                Case 1
                    mOutString.Append("Bad Car Type")
                Case 2
                    mOutString.Append("Bad Tonnage")
                Case 3
                    mOutString.Append("Bad Revenue")
                Case 4
                    mOutString.Append("Bad Distance")
                Case 5
                    mOutString.Append("Bad Intermodal Specification")
                Case 6
                    mOutString.Append("Non US Segments")
                Case 7
                    mOutString.Append("Negative Cost")
            End Select

            mOutString.Append(Space(40 - mOutString.Length))
            mOutString.Append(Right_Justify(mRejectedSegments(mLooper2).ToString("N0"), 7))
            mOutString.Append(Space(52 - mOutString.Length))
            mOutString.Append(Right_Justify(mRejectedWaybills(mLooper2).ToString("N0"), 7))
            sbLogFile.WriteLine(mOutString.ToString)
            sbLogFile.Flush()
        Next

        mOutString = New StringBuilder
        mOutString.Append("Total Rejected Segments")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify((mRejectedSegments(BadCarType) +
            mRejectedSegments(BadTonnage) +
            mRejectedSegments(BadRevenue) +
            mRejectedSegments(BadDistance) +
            mRejectedSegments(BadIntermodalSpec) +
            mRejectedSegments(NonUSSegment) +
            mRejectedSegments(NegativeCost)).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify((mRejectedWaybills(BadCarType) +
            mRejectedWaybills(BadTonnage) +
            mRejectedWaybills(BadRevenue) +
            mRejectedWaybills(BadDistance) +
            mRejectedWaybills(BadIntermodalSpec) +
            mRejectedWaybills(NonUSSegment) +
            mRejectedWaybills(NegativeCost)).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        ' Header Line
        mOutString = New StringBuilder
        mOutString.Append("Costed As")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append("Segments")
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append("Waybills")
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        ' Single Carload Data Line
        mOutString = New StringBuilder
        mOutString.Append("Single Carload")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedSegments(SingleMove).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedWaybills(SingleMove).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        'Multiple Carload Data Line
        mOutString = New StringBuilder
        mOutString.Append("Multiple Carload")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedSegments(MultiMove).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedWaybills(MultiMove).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        ' Unit Train Data line
        mOutString = New StringBuilder
        mOutString.Append("Unit Train")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedSegments(UnitMove).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedWaybills(UnitMove).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        ' Intermodal Data Line
        mOutString = New StringBuilder
        mOutString.Append("Intermodal")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedSegments(IntermodalMove).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify(mCostedWaybills(IntermodalMove).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        ' Total Costed Data line
        mOutString = New StringBuilder
        mOutString.Append("Total Costed")
        mOutString.Append(Space(40 - mOutString.Length))
        mOutString.Append(Right_Justify((mCostedSegments(SingleMove) +
                          mCostedSegments(MultiMove) +
                          mCostedSegments(UnitMove) +
                          mCostedSegments(IntermodalMove)).ToString("N0"), 7))
        mOutString.Append(Space(52 - mOutString.Length))
        mOutString.Append(Right_Justify((mCostedWaybills(SingleMove) +
                          mCostedWaybills(MultiMove) +
                          mCostedWaybills(UnitMove) +
                          mCostedWaybills(IntermodalMove)).ToString("N0"), 7))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        ' Write out the result of the Total Variable Cost
        mOutString = New StringBuilder
        mOutString.Append(Right_Justify("Total", 21))
        sbLogFile.WriteLine(mOutString.ToString)
        mOutString = New StringBuilder
        mOutString.Append(Right_Justify("Variable", 21))
        sbLogFile.WriteLine(mOutString.ToString)
        mOutString = New StringBuilder
        mOutString.Append(Right_Justify("Costs", 21))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        For mLooper = 1 To 9
            mOutString = New StringBuilder
            mOutString.Append(Get_Short_Name_By_RRID(mLooper))
            mOutString.Append(Space(6 - mOutString.Length))
            mOutString.Append(Right_Justify(mTotalVariableCosts(mLooper).ToString("N0"), 15))
            sbLogFile.WriteLine(mOutString.ToString)
            sbLogFile.Flush()

            mGrandTotal = mGrandTotal + mTotalVariableCosts(mLooper)
        Next

        sbLogFile.WriteLine()
        mOutString = New StringBuilder
        mOutString.Append("Total")
        mOutString.Append(Space(6 - mOutString.Length))
        mOutString.Append(Right_Justify(mGrandTotal.ToString("N0"), 15))
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.WriteLine()
        sbLogFile.Flush()

        mPass3Starts = Now

pass3:
        mPass2Starts = Now()
        mOutString = New StringBuilder

        If chk_UseDifferentYear.Checked = True And chk_SaveToSQL.Checked = False Then
            mOutString.Append("*** As Requested, Updates NOT Saved to SQL")
        Else
            mOutString.Append("*** Update of Masked Waybill Table with Variable Costs from Segments table began at  " & DateTime.Now.ToString("G"))
        End If

        sbLogFile.WriteLine()
        sbLogFile.WriteLine(mOutString.ToString)
        sbLogFile.Flush()

        ' Now that we have the segments VC value updated, we'll get the masked table into the datatable
        mMaskedTable = Get_Masked_VC_Data(mThreesOnly)

        mCounter = 0

        If chk_SaveToSQL.Checked = True Or chk_UseDifferentYear.Checked = False Then

            Text = "Pass 3 of 3"
            txtStatus.Text = "Pass 3 of 3 - Updating Masked Waybill record " & mCounter.ToString & " of " _
                    & mDataTable.Rows.Count.ToString & " - " & ((mLooper / mDataTable.Rows.Count) * 100).ToString("N1") & "%..."
            Refresh()

            ' we'll now loop thru the masked table
            For mLooper = 0 To mMaskedTable.Rows.Count - 1
                mCounter = mCounter + 1
                If mCounter Mod 100 = 0 Then
                    txtStatus.Text = "Pass 3 of 3 - Updating Masked Waybill record " & mCounter.ToString & " of " _
                    & mMaskedTable.Rows.Count.ToString & " - " & ((mCounter / mMaskedTable.Rows.Count) * 100).ToString("N1") & "%..."
                    Refresh()
                    Application.DoEvents()
                End If

                If mCounter Mod 100000 = 0 Then
                    sbLogFile.WriteLine(“*** Processing record " & mCounter.ToString & " of the Segments table at " & DateTime.Now.ToString("G"))
                End If

                ' for each record in the masked table, get the segments for that record.
                mDataTable = Get_Segments_VC_Data(mMaskedTable.Rows(mLooper)("Serial_No"))
                mProcess = True

                For mLooper2 = 0 To mDataTable.Rows.Count - 1

                    ' if any of the segments has a negative value, set all segments in the waybill to zero
                    If mDataTable.Rows(mLooper2)("RR_VC") < 0 Then
                        ' Add this segment to the negative cost counter
                        mRejectedSegments(NegativeCost) = mRejectedSegments(NegativeCost) + 1
                        ' Set all VC fields in the Masked record to 0
                        mStrSQL = "Update " & Gbl_Masked_TableName & " SET Total_VC = 0, RR1_VC = 0, RR2_VC = 0, RR3_VC = 0, " &
                            "RR4_VC = 0, RR5_VC = 0, RR6_VC = 0, RR7_VC = 0, RR8_VC = 0 WHERE Serial_No = '" &
                            mMaskedTable.Rows(mLooper)("Serial_No") & "'"
                        mSQLcmd = New SqlCommand
                        mSQLcmd.Connection = gbl_SQLConnection
                        mSQLcmd.CommandType = CommandType.Text
                        mSQLcmd.CommandText = mStrSQL
                        mSQLcmd.ExecuteNonQuery()
                        mProcess = False

                        ' Set all Segments RR_VC fields to 0 for this serial number
                        mStrSQL = "Update " & Gbl_Segments_TableName & " SET RR_VC = 0 WHERE Serial_No = '" &
                            mMaskedTable.Rows(mLooper)("Serial_No") & "'"
                        mSQLcmd = New SqlCommand
                        mSQLcmd.Connection = gbl_SQLConnection
                        mSQLcmd.CommandType = CommandType.Text
                        mSQLcmd.CommandText = mStrSQL
                        mSQLcmd.ExecuteNonQuery()
                        Exit For
                    End If
                Next

                ' The waybill has a negative cost segment, so we have to add 1 to the rejected waybills counter array
                If mProcess = False Then
                    mRejectedWaybills(NegativeCost) = mRejectedWaybills(NegativeCost) + 1
                End If

                mThisCost = 0

                For mLooper2 = 0 To mDataTable.Rows.Count - 1

                    If mProcess = True Then
                        ' Update the VC fields in the Masked table
                        mStrSQL = "Update " & Gbl_Masked_TableName & " SET RR" & (mLooper2 + 1).ToString & "_VC = " &
                            mDataTable.Rows(mLooper2)("RR_VC").ToString & " WHERE Serial_No = '" &
                            mMaskedTable.Rows(mLooper)("Serial_No") & "'"
                        mSQLcmd = New SqlCommand
                        mSQLcmd.Connection = gbl_SQLConnection
                        mSQLcmd.CommandType = CommandType.Text
                        mSQLcmd.CommandText = mStrSQL
                        mSQLcmd.ExecuteNonQuery()
                        mThisCost = mThisCost + mDataTable.Rows(mLooper2)("RR_VC")
                    End If
                Next

                ' If we processed this, we have to update the total_VC value in the masked record
                If mProcess = True And chk_SaveToSQL.Checked = True Then
                    mStrSQL = "Update " & Gbl_Masked_TableName & " SET Total_VC = " &
                        mThisCost.ToString & " WHERE Serial_No = '" &
                        mMaskedTable.Rows(mLooper)("Serial_No") & "'"
                    mSQLcmd = New SqlCommand
                    mSQLcmd.Connection = gbl_SQLConnection
                    mSQLcmd.CommandType = CommandType.Text
                    mSQLcmd.CommandText = mStrSQL
                    mSQLcmd.ExecuteNonQuery()
                End If

            Next

            mOutString = New StringBuilder
            mOutString.Append("*** Update of Masked Waybill Table with Variable Costs from Segments table Completed at " & DateTime.Now.ToString("G") & ". " & Return_Elapsed_Time(mPass3Starts, Now()))
            sbLogFile.WriteLine(mOutString.ToString)

        Else
            If chk_SaveToSQL.Checked = False Then
                mOutString.Append("*** Update of Masked Waybill skipped as requested.")
            End If
        End If

        sbLogFile.Flush()

        'Housekeeping
        mP3WorkbookSet.Calculation = Calculation.Automatic
        mP3Workbook.Save()
        mP3Workbook.Close()

        mMWFWorkbook.Close()

        mDataTable = Nothing

        'close the CSV files if they were created.
        If chk100Series.Checked Then
            sb100Series_Costed.Flush()
            sb100Series_Costed.Close()
        End If

        If chk200Series.Checked Then
            sb200Series_Costed.Flush()
            sb200Series_Costed.Close()
        End If

        If chk300Series.Checked Then
            sb300Series_Costed.Flush()
            sb300Series_Costed.Close()
        End If

        If chk400Series.Checked Then
            sb400Series_Costed.Flush()
            sb400Series_Costed.Close()
        End If

        If chk500Series.Checked Then
            sb500Series_Costed.Flush()
            sb500Series_Costed.Close()
        End If

        If chk600Series.Checked Then
            sb600Series_Costed.Flush()
            sb600Series_Costed.Close()
        End If

        If chkSaveResults.Checked Then
            sbResults_Costed.Flush()
            sbResults_Costed.Close()
        End If

        If chk_Save_CRPRESRecords.Checked Then
            sbCRPRESRecord.Flush()
            sbCRPRESRecord.Close()
        End If

        sbLogFile.WriteLine()
        mWorkStr = "Run Completed at " & DateTime.Now.ToString("G") & ".  " & Return_Elapsed_Time(mPass1Starts, Now())
        sbLogFile.WriteLine(mWorkStr)
        sbLogFile.Flush()
        sbLogFile.Close()

        Text = "Done"

EndIt:
        txtStatus.Text = "Done! - Elapsed time: " & Return_Elapsed_Time(mPass1Starts, Now())
        btn_Return_To_MainMenu.Enabled = True
        btnExecute.Enabled = True

        Refresh()

    End Sub

    Function Get_Masked_VC_Data(ByVal mThreesOnly As Boolean) As DataTable
        Dim mStrSQL As String

        Get_Masked_VC_Data = New DataTable

        mStrSQL = "SELECT Serial_No, Total_VC, RR1_VC, RR2_VC, RR3_VC, RR4_VC, RR5_VC, RR6_VC, RR7_VC, RR8_VC " &
            " FROM " & Gbl_Masked_TableName

        If mThreesOnly = True Then
            'mStrSQL = mStrSQL & " where (right(" & Gbl_Masked_TableName & ".serial_no,1) = 3) AND " & Gbl_Masked_TableName & ".serial_no < 250000)"
            mStrSQL = mStrSQL & " where right(serial_no,2) = 33"
        End If

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        Using daAdapter As New SqlClient.SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(Get_Masked_VC_Data)
        End Using

    End Function

    Function Get_Segments_VC_Data(ByVal Serial_No As String) As DataTable
        Dim mStrSQL As String

        Get_Segments_VC_Data = New DataTable

        mStrSQL = "SELECT RR_VC FROM " & Gbl_Segments_TableName & " WHERE Serial_No = '" & Serial_No.ToString & "'"

        Using daAdapter As New SqlClient.SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(Get_Segments_VC_Data)
        End Using

    End Function

    Private Sub rdo_EP431Sub4_CheckedChanged(sender As Object, e As EventArgs) Handles rdo_Current.CheckedChanged
        If rdo_Current.Checked = True Then
            chk_Cost_All_Segments.Visible = True
        Else
            chk_Cost_All_Segments.Checked = False
            chk_Cost_All_Segments.Visible = False
        End If
    End Sub

    Private Sub chk_UseDifferentYear_CheckedChanged(sender As Object, e As EventArgs) Handles chk_UseDifferentYear.CheckedChanged
        If chk_UseDifferentYear.Checked = True Then
            cmb_Different_Year.Visible = True
            chk_SaveToSQL.Checked = False
        Else
            cmb_Different_Year.Visible = False
        End If
    End Sub

End Class
