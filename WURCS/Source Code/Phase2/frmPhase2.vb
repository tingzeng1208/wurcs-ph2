Imports System.IO
Imports System.Threading
Imports System.Xml
Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Public Class frmPhase2

#Region "Declarations"

    Private oDB As DBManager
    Private oSpreadSheet As ExcelSpreadsheet
    Private oPhase2InputData As Phase2InputData

    Private Delegate Sub DelegateUpdateStatus(ByVal Completed As Boolean, ByVal StatusText As String, ByVal RailRoadNumber As Integer)
    Private Delegate Sub DelegateErrorOccured(ByVal Message As String)
    Private Delegate Sub DelegateThreadDone()
    Private Event ThreadDone()

#End Region

#Region "Form Events"

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cnConnection As SqlConnection
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As DataSet
        Dim mTable_Name As String
        Dim mDatabase_Name As String
        Dim mConnStr As String

        ' MRS - Load the year values to the year combobox
        'Get the database_name and table_name for the URCS_Years value from the database
        mDatabase_Name = Get_Database_Name(1, "URCS_YEARS")
        mTable_Name = Get_Table_Name(1, "URCS_YEARS")

        ' Load the Year combobox from the SQL database
        ' Open the SQL connection using the global variable holding the connection string
        mConnStr = BuildConnectionString(Global_Variables.gbl_Server_Name, mDatabase_Name)
        cnConnection = New SqlConnection(mConnStr)
        cnConnection.Open()

        ' Execute the query
        daAdapter = New SqlDataAdapter("SELECT urcs_year FROM " & mTable_Name, cnConnection)
        cnConnection.Close()

        dsDataSet = New DataSet
        daAdapter.Fill(dsDataSet)

        ' Load the Year values into the combobox
        For i = 0 To dsDataSet.Tables(0).Rows.Count - 1
            Me.cmbYear.Items.Add(dsDataSet.Tables(0).Rows(i).Item(0).ToString)
        Next

        daAdapter.Dispose()
        dsDataSet.Dispose()
        cnConnection.Dispose()

        oDB = New DBManager

        AddHandler ThreadDone, AddressOf ThreadIsDone

        clbRailroads.DataSource = oDB.GetClass1RailList()
        clbRailroads.DisplayMember = "Name"

        txtFolder.Text = My.Settings.OutputDirectory
        cbLog.Checked = My.Settings.CreateLog

        cbAll.Checked = True
        EnableControls(True)

    End Sub

    Private Sub cbAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAll.CheckedChanged

        For i As Integer = 0 To clbRailroads.Items.Count - 1
            clbRailroads.SetItemChecked(i, cbAll.Checked)
        Next i

        clbRailroads.ClearSelected()

    End Sub

    Private Sub btnReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReport.Click

        If ValidateFormData() And ValidateFilesData() Then

            Dim oWorkerThread As Thread

            'Create output report(s)
            oWorkerThread = New Thread(New ThreadStart(AddressOf CreateReports))
            oWorkerThread.Start()

            clbRailroads.SelectedItem = Nothing
            EnableControls(False)
        End If

    End Sub

    Private Sub btnEValues_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEValues.Click

        Dim oWorkerThread As Thread

        'Create output report(s)
        oWorkerThread = New Thread(New ThreadStart(AddressOf CreateEValues))
        oWorkerThread.Start()

        clbRailroads.SelectedItem = Nothing
        EnableControls(False)

    End Sub

    Private Sub btnLaunch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLaunch.Click

        If ValidateFormData() And ValidateFilesData() Then

            Dim oWorkerThread As Thread

            'Create output report(s)
            oWorkerThread = New Thread(New ThreadStart(AddressOf LaunchProcess))
            oWorkerThread.Start()

            EnableControls(False)
        End If

    End Sub

    Private Sub btnFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFolder.Click

        Dim oFolderBrowserDialog As FolderBrowserDialog

        oFolderBrowserDialog = New FolderBrowserDialog
        oFolderBrowserDialog.ShowDialog()

        If oFolderBrowserDialog.SelectedPath.Length > 0 Then
            txtFolder.Text = oFolderBrowserDialog.SelectedPath
        End If

    End Sub

    Private Sub rbStepByStep_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbStepByStep.CheckedChanged

        EnableControls(True)

    End Sub

#End Region

#Region "Delegates handlers"

    Private Sub ThreadIsDone()

        Me.Invoke(New DelegateThreadDone(AddressOf EnableControls))

    End Sub

    Private Sub SpreadSheetStatusUpdated(ByVal RailRoadNumber As Integer, ByVal StatusText As String, ByVal Completed As Boolean)

        Me.Invoke(New DelegateUpdateStatus(AddressOf UpdateStatus), Completed, StatusText, RailRoadNumber)

    End Sub

    Private Sub PrepareDataStatusUpdated(ByVal StatusText As String, ByVal Completed As Boolean)

        Me.Invoke(New DelegateUpdateStatus(AddressOf UpdateStatus), Completed, StatusText, -1)

    End Sub

    Private Sub ThreadErrorOccured(ByVal Message As String)

        Me.Invoke(New DelegateErrorOccured(AddressOf ErrorOccured), Message)

    End Sub

#End Region


#Region "Helper functions"

    Private Sub CreateReports(Optional ByVal RaiseThreadDoneEvent As Boolean = True)

        SaveSettings()

        oPhase2InputData = New Phase2InputData(cmbYear.Text, oDB)
        AddHandler oPhase2InputData.StatusUpdated, AddressOf PrepareDataStatusUpdated
        AddHandler oPhase2InputData.ErrorOccured, AddressOf ThreadErrorOccured

        If oPhase2InputData.Prepare() Then

            'Create output files
            For Each oDataRow As DataRowView In clbRailroads.CheckedItems
                oSpreadSheet = New ExcelSpreadsheet(txtFolder.Text, cmbYear.Text, oDataRow("RR_ID"), oDataRow("RRICC"), oDataRow("SHORT_NAME"), _
                                                                                oDataRow("NAME"), "n/a", 5, False, False, True, False, False, cbLog.Checked, oDB)

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

        SaveSettings()

        oDirectory = New IO.DirectoryInfo(txtFolder.Text)
        oFiles = oDirectory.GetFiles("*" & cmbYear.Text & ".xlsx")

        If oFiles.Count > 0 Then
            For Each oFile As FileInfo In oFiles
                oFile.Refresh()
                If oFile.Exists Then
                    UpdateStatus(False, "Loading file " & oFile.Name)
                    oWorkbook = oExcelApp.Workbooks.Open(oFile.FullName, , True)
                    oWorksheet = oWorkbook.Sheets(148)
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
                oDB.CreateEValues(cmbYear.Text)
                UpdateStatus(True, "All E-Values for " & sYear & " are loaded into the database")
            End If
        Else
            MessageBox.Show("The output folder does not contain required files.")
        End If

        RaiseEvent ThreadDone()

    End Sub

    Private Sub LaunchProcess()

        CreateReports(False)
        CreateEValues()

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

    Private Function ValidateFormData(Optional ByVal ReportInputs As Boolean = True) As Boolean

        If ReportInputs Then
            If clbRailroads.CheckedItems.Count = 0 Then
                MessageBox.Show("Please, choose at least one railroad to process.")
                Return False
            End If
        End If

        Try
            Integer.Parse(cmbYear.Text)
        Catch
            MessageBox.Show("The year entered does not appear to be a valid value.")
            Return False
        End Try

        If Integer.Parse(cmbYear.Text) < 1990 Then
            MessageBox.Show("Please enter a valid 4 digit year greater than 1990")
            Return False
        End If

        If Integer.Parse(cmbYear.Text) > Now.Year Then
            MessageBox.Show("The year entered is invalid as it is in the future")
            Return False
        End If

        Return True
    End Function

    Private Sub UpdateStatus(ByVal Completed As Boolean, ByVal StatusText As String, Optional ByVal RailRoadNumber As Integer = -1)

        If RailRoadNumber >= 0 Then
            For i As Integer = 0 To clbRailroads.Items.Count - 1
                If clbRailroads.Items(i)("RR_ID") = RailRoadNumber Then
                    tssLabel.Text = clbRailroads.Items(i)("NAME") & ": " & StatusText
                    Exit For
                End If
            Next
        Else
            tssLabel.Text = StatusText
        End If
    End Sub

    Private Sub ErrorOccured(ByVal Message As String)

        MessageBox.Show(Message)

    End Sub

    Private Sub EnableControls(Optional ByVal Status As Boolean = True)

        grpParameters.Enabled = Status
        grpRun.Enabled = Status

        If Status Then
            btnLaunch.Enabled = rbAllSteps.Checked
            btnReport.Enabled = rbStepByStep.Checked
            btnEValues.Enabled = rbStepByStep.Checked
        End If

    End Sub

    Private Sub SaveSettings()

        With My.Settings
            .OutputDirectory = txtFolder.Text
            .CurrentYear = cmbYear.Text
            .CreateLog = cbLog.Checked
            .Save()
        End With

    End Sub

#End Region

End Class
