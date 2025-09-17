Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class ProductionErrorLog

    Dim ConnString As String
    Dim dsErrors As DataTable

    Private Sub ProductionErrorLog_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

        'Get the connection string from the app.config
        ConnString = New String(ConfigurationManager.AppSettings("ConnectionString"))

        'get the error table from SQL
        dsErrors = New DataTable
        dsErrors = getErrorTable()

        'bind the gridview to the table
        gv_Errors.DataSource = dsErrors

        'delete row index column
        gv_Errors.RowHeadersVisible = False
        'for appearance
        gv_Errors.RowsDefaultCellStyle.BackColor = Color.Gray
        gv_Errors.AutoResizeColumns( _
            DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader)
        'display datagridview in form
        gv_Errors.Visible = True

    End Sub

    Private Sub btn_exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_exit.Click
        Dim _Production As New Production
        _Production.Show()

        Me.Close()

    End Sub

    ''' <summary>
    ''' Gets a local copy of the error table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getErrorTable() As DataTable

        Dim cnConnection As New SqlConnection
        Dim cmdCommand As SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try

            cnConnection.ConnectionString = ConnString
            cnConnection.Open()

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "Select * from U_ERRORS order by error_timestamp" 'Simple select query for Railroad table
            cmdCommand.Connection = cnConnection

            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand

            daAdapter.Fill(dsDataSet, "U_ERRORS")

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "Errors"

        Catch ex As Exception
            'if we get an error toss it to the web app
            Throw (ex)
        Finally
            'if we failed or succedded, close the sql connection
            If cnConnection.State = ConnectionState.Open Then
                cnConnection.Close()
            End If
        End Try

        Return dsDataSet.Tables(0)

        'clean up
        cnConnection.Dispose()
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

End Class