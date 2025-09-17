Imports System.Threading
Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Public Class frmPhase2Main

#Region "Declarations"

    Private oDB As DBManager
    Private oSpreadSheet As ExcelSpreadsheet
    Private oPhase2InputData As Phase2InputData

    Private Delegate Sub DelegateUpdateStatus(Completed As Boolean, StatusText As String, RailRoadNumber As Integer)
    Private Delegate Sub DelegateErrorOccured(Message As String)
    Private Delegate Sub DelegateThreadDone()
    Private Event ThreadDone()

#End Region

#Region "Form Events"

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        mDataTable = Get_URCS_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            cmb_URCS_Year.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
        Next

        mDataTable = Nothing

        oDB = New DBManager

        AddHandler ThreadDone, AddressOf ThreadIsDone

        'Set up the checkbox container
        clbRailroads.DataSource = oDB.GetClass1RailList()
        clbRailroads.DisplayMember = "Name"

        txtFolder.Text = My.Settings.OutputDirectory
        cbLog.Checked = My.Settings.CreateLog

        cbAll.Checked = True
        EnableControls(True)

    End Sub

    Private Sub cbAll_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles cbAll.CheckedChanged

        For i As Integer = 0 To clbRailroads.Items.Count - 1
            clbRailroads.SetItemChecked(i, cbAll.Checked)
        Next i

        clbRailroads.ClearSelected()

    End Sub

    Private Sub btnReport_Click(sender As System.Object, e As System.EventArgs) Handles btnReport.Click



        If ValidateFormData() And ValidateFilesData() Then

            Dim oWorkerThread As Thread

            My.Settings.CurrentYear = cmb_URCS_Year.SelectedItem.ToString
            My.Settings.Save()

            'Create output report(s)
            oWorkerThread = New Thread(New ThreadStart(AddressOf CreateReports))
            oWorkerThread.Start()

            clbRailroads.SelectedItem = Nothing
            EnableControls(False)


        End If

    End Sub

    Private Sub btnEValues_Click(sender As System.Object, e As System.EventArgs) Handles btnEValues.Click

        Dim oWorkerThread As Thread

        'Create output report(s)
        oWorkerThread = New Thread(New ThreadStart(AddressOf CreateEValues))
        oWorkerThread.Start()

        clbRailroads.SelectedItem = Nothing
        EnableControls(False)

    End Sub

    Private Sub btnLaunch_Click(sender As System.Object, e As System.EventArgs) Handles btnLaunch.Click



        If ValidateFormData() And ValidateFilesData() Then

            Dim oWorkerThread As Thread

            SaveSettings(cmb_URCS_Year.Text)

            'Create output report(s)
            oWorkerThread = New Thread(New ThreadStart(AddressOf LaunchProcess))
            oWorkerThread.Start()

            EnableControls(False)
        End If

    End Sub

    Private Sub btnFolder_Click(sender As System.Object, e As System.EventArgs) Handles btnFolder.Click

        Dim oFolderBrowserDialog As FolderBrowserDialog

        oFolderBrowserDialog = New FolderBrowserDialog
        oFolderBrowserDialog.ShowDialog()

        If oFolderBrowserDialog.SelectedPath.Length > 0 Then
            txtFolder.Text = oFolderBrowserDialog.SelectedPath
        End If

    End Sub

    Private Sub rbStepByStep_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbStepByStep.CheckedChanged

        EnableControls(True)

    End Sub

    Private Sub btn_Return_To_MainMenu_Click(sender As System.Object, e As System.EventArgs) Handles btn_Return_To_MainMenu.Click
        ' Open the Main Menu Form
        Dim frmNew As New frm_MainMenu()
        frmNew.Show()
        ' Close this Menu
        Me.Close()
    End Sub

#End Region

#Region "Delegates handlers"

    Private Sub ThreadIsDone()

        Me.Invoke(New DelegateThreadDone(AddressOf EnableControls))

    End Sub

    Private Sub SpreadSheetStatusUpdated(RailRoadNumber As Integer, StatusText As String, Completed As Boolean)

        Me.Invoke(New DelegateUpdateStatus(AddressOf UpdateStatus), Completed, StatusText, RailRoadNumber)

    End Sub

    Private Sub PrepareDataStatusUpdated(StatusText As String, Completed As Boolean)

        Me.Invoke(New DelegateUpdateStatus(AddressOf UpdateStatus), Completed, StatusText, -1)

    End Sub

    Private Sub ThreadErrorOccured(Message As String)

        Me.Invoke(New DelegateErrorOccured(AddressOf ErrorOccured), Message)

    End Sub

#End Region

#Region "Helper functions"

    Private Sub CreateReports(Optional RaiseThreadDoneEvent As Boolean = True)

        oPhase2InputData = New Phase2InputData(My.Settings.CurrentYear.ToString, oDB)
        AddHandler oPhase2InputData.StatusUpdated, AddressOf PrepareDataStatusUpdated
        AddHandler oPhase2InputData.ErrorOccured, AddressOf ThreadErrorOccured

        Logger.Log("Preparing data for Phase 2 reports for year " & My.Settings.CurrentYear.ToString)
        If oPhase2InputData.Prepare() Then

            'Create output files
            For Each oDataRow As DataRowView In clbRailroads.CheckedItems
                oSpreadSheet = New ExcelSpreadsheet(txtFolder.Text, My.Settings.CurrentYear, oDataRow("RR_ID"), oDataRow("RRICC"), oDataRow("SHORT_NAME"),
                                                                               oDataRow("NAME"), "n/a", 5, False, False, True, False, False, False, cbLog.Checked, oDB)

                AddHandler oSpreadSheet.StatusUpdated, AddressOf SpreadSheetStatusUpdated
                oSpreadSheet.Create()
                oSpreadSheet = Nothing
            Next
        End If

        oPhase2InputData = Nothing

        If RaiseThreadDoneEvent Then RaiseEvent ThreadDone()

    End Sub

    Private Sub CreateEValues()

        Dim oExcelApp As New Excel.Application
        Dim oWorkbook As Excel.Workbook
        Dim oWorksheet As Excel.Worksheet
        Dim oRange As Excel.Range
        Dim sYear As String = ""
        Dim sRRid As String
        Dim oValues As Array
        Dim oRailroads As Integer() = Nothing

        Dim oDirectory As DirectoryInfo
        Dim oFiles As FileInfo()

        'SaveSettings(My.Settings.CurrentYear)

        oDirectory = New IO.DirectoryInfo(txtFolder.Text)
        oFiles = oDirectory.GetFiles("*" & My.Settings.CurrentYear & ".xlsx")

        If oFiles.Count > 0 Then
            For Each oFile As FileInfo In oFiles
                oFile.Refresh()
                If oFile.Exists Then
                    UpdateStatus(False, "Loading file " & oFile.Name)
                    oWorkbook = oExcelApp.Workbooks.Open(oFile.FullName, , True)
                    oWorksheet = oWorkbook.Sheets("E Summary")
                    oRange = oWorksheet.UsedRange

                    'Retrieve year and railroad id
                    sYear = CType(oRange.Cells(8, 1), Excel.Range).Value.ToString
                    sRRid = CType(oRange.Cells(8, 3), Excel.Range).Value.ToString

                    'Retrieve eCode values
                    oValues = CType(oWorksheet.Range("F8:G980"), Excel.Range).Value

                    UpdateStatus(False, "Inserting values into the database for " & oFile.Name)
                    oDB.InsertSubstitutions(sYear, sRRid, oValues)

                    If oRailroads Is Nothing Then
                        ReDim Preserve oRailroads(0)
                    Else
                        ReDim Preserve oRailroads(oRailroads.Length)
                    End If

                    oRailroads(oRailroads.Length - 1) = sRRid
                End If
            Next

            oDB.RunSubstitutions(sYear, oRailroads)
            UpdateStatus(False, "All files are loaded into the database")

            If ValidateFormData(ReportInputs:=False) Then
                oDB.CreateEValues(My.Settings.CurrentYear)
                UpdateStatus(True, "All E-Values for " & sYear & " are loaded My.Settings.CurrentYearinto the database")

                'Only create XML file if all roads are processed.
                If cbAll.Checked Then
                    UpdateStatus(True, "Writing XML File")
                    CreateXML(My.Settings.CurrentYear.ToString)

                    UpdateStatus(True, "Finished")
                End If
            End If
        Else
            MessageBox.Show("The output folder does not contain required files.")
        End If

        RaiseEvent ThreadDone()

    End Sub

    Private Sub LaunchProcess()

        CreateReports(False)
        CreateEValues()

        'Only create XML file if all roads are processed.
        If cbAll.Checked Then
            UpdateStatus(True, "Writing XML File")
            CreateXML(My.Settings.CurrentYear.ToString)

            UpdateStatus(True, "Finished")
        End If

        EnableControls(True)

    End Sub

    Private Function ValidateFilesData() As Boolean

        'Check if output folder exists
        If Not Directory.Exists(txtFolder.Text) Then
            MessageBox.Show("The output folder does not exist. Please, choose a valid folder.")
            Return False
        End If

        'Check if files from previous run exist
        If Directory.GetFiles(txtFolder.Text).Count > 0 Then
            If MessageBox.Show("The output folder contains files which will be overriden during the run. Would you like to continue?", "Warning", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                Return False
            End If
        End If

        'Create log folder
        If cbLog.Checked And Not Directory.Exists(txtFolder.Text & "\Log") Then
            Directory.CreateDirectory(txtFolder.Text & "\Log")
        End If

        Return True

    End Function

    Private Function ValidateFormData(Optional ReportInputs As Boolean = True) As Boolean

        If ReportInputs Then
            If clbRailroads.CheckedItems.Count = 0 Then
                MessageBox.Show("Please, choose at least one railroad to process.")
                Return False
            End If
        End If

        If Integer.Parse(My.Settings.CurrentYear) < 1990 Then
            MessageBox.Show("Please enter a valid 4 digit year greater than 1990.")
            Return False
        End If

        If Integer.Parse(My.Settings.CurrentYear) > Now.Year Then
            MessageBox.Show("The year entered is invalid as it is in the future.")
            Return False
        End If

        Return True
    End Function

    Private Sub UpdateStatus(Completed As Boolean, StatusText As String, Optional RailRoadNumber As Integer = -1)

        If RailRoadNumber >= 0 Then
            For i As Integer = 0 To clbRailroads.Items.Count - 1
                If clbRailroads.Items(i)("RR_ID") = RailRoadNumber Then
                    tssLabel.Text = clbRailroads.Items(i)("NAME") & ": " & StatusText
                    Exit For
                End If
            Next
        Else
            Control.CheckForIllegalCrossThreadCalls = False
            TSSLabel1.Text = StatusText
            Control.CheckForIllegalCrossThreadCalls = True
        End If
    End Sub

    Private Sub ErrorOccured(Message As String)

        MessageBox.Show(Message)

    End Sub

    Private Sub EnableControls(Optional Status As Boolean = True)

        grpParameters.Enabled = Status
        grpRun.Enabled = Status
        btn_Return_To_MainMenu.Enabled = Status

        If Status Then
            btnLaunch.Enabled = rbAllSteps.Checked
            btnReport.Enabled = rbStepByStep.Checked
            btnEValues.Enabled = rbStepByStep.Checked
        End If

    End Sub

    Private Sub SaveSettings(ByVal mYear As String)

        With My.Settings
            .OutputDirectory = txtFolder.Text

            If mYear.Length = 4 Then
                .CurrentYear = mYear.ToString
            End If

            .CreateLog = cbLog.Checked
            .Save()
        End With

    End Sub

    Private Sub CreateXML(ByVal Year As String)

        Dim cmdCommand As SqlCommand
        Dim oWriter As System.IO.StreamWriter
        Dim oDirectory As DirectoryInfo
        Dim oFiles As FileInfo()
        Dim sFileName As String, mWorkString As String

        OpenSQLConnection(Get_Database_Name_From_SQL(My.Settings.CurrentYear.ToString, "EVALUES"))

        cmdCommand = New SqlClient.SqlCommand
        cmdCommand.CommandType = CommandType.StoredProcedure
        cmdCommand.CommandText = "usp_GenerateEValuesXML"
        cmdCommand.Connection = Global_Variables.gbl_SQLConnection

        Dim inputValue As SqlParameter = New SqlParameter("@Year", SqlDbType.Char, 4)
        inputValue.Value = Year
        inputValue.Direction = ParameterDirection.Input
        cmdCommand.Parameters.Add(inputValue)

        Dim returnValue As SqlParameter = New SqlParameter("@Output", SqlDbType.Xml)
        returnValue.Direction = ParameterDirection.Output
        cmdCommand.Parameters.Add(returnValue)

        cmdCommand.ExecuteReader()
        cmdCommand.CommandTimeout = 60

        'If file already exists, delete the existing
        sFileName = My.Settings.FileName.Replace("%year%", Year.ToString)

        oDirectory = New IO.DirectoryInfo(txtFolder.Text)
        oFiles = oDirectory.GetFiles("*.xml")
        If oFiles.Count > 0 Then
            For Each oFile As FileInfo In oFiles
                If oFile.Name = sFileName Then
                    oFile.Delete()
                End If
            Next
        End If

        'Create XML file
        oWriter = New System.IO.StreamWriter(txtFolder.Text + "\" + sFileName, True, System.Text.Encoding.Unicode)
        mWorkString = "<?xml version=""1.0"" encoding=""utf-16"" standalone=""yes""?>"
        oWriter.WriteLine(mWorkString)
        mWorkString = returnValue.Value
        mWorkString = Replace(mWorkString, "<R", "  <R")
        mWorkString = Replace(mWorkString, "</R", "  </R")
        mWorkString = Replace(mWorkString, "<E", "    <E")
        mWorkString = Replace(mWorkString, "a>", "a>" & vbCrLf)
        mWorkString = Replace(mWorkString, "d>", "d>" & vbCrLf)
        mWorkString = Replace(mWorkString, """>", """>" & vbCrLf)
        mWorkString = Replace(mWorkString, "-->", "-->" & vbCrLf)
        mWorkString = Replace(mWorkString, "/>", "/>" & vbCrLf)

        oWriter.WriteLine(mWorkString)
        oWriter.Close()


    End Sub

#End Region

End Class
