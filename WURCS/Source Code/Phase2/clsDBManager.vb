Imports System.Data.SqlClient
Imports System.Text

Public Class DBManager

    Dim sConnection As String

    Public Sub New()

        sConnection = My.Settings.ConnectionString

    End Sub

    Public ReadOnly Property Connection As String
        Get
            Return sConnection
        End Get
    End Property

    Public Function GetClass1RailList() As DataTable

        Dim cmdCommand As SqlCommand
        Dim daAdapter As New SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try

            ' Open the SQL connection to the Controls database
            Gbl_Class1RailList_DatabaseName = Get_Database_Name_From_SQL("1", "CLASS1RAILLIST")
            Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "CLASS1RAILLIST")
            OpenSQLConnection(Gbl_Class1RailList_DatabaseName)

            cmdCommand = New SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "SELECT * FROM " & Gbl_Class1RailList_TableName & " WHERE RR_ID<>0 ORDER BY RR_ID"
            Logger.Log("GetClass1RailList: " & cmdCommand.CommandText) ' <--- Add this line

            cmdCommand.Connection = gbl_SQLConnection

            'fill the dataset
            daAdapter = New SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "Class1RailList"

        Catch SqlEx As SqlException
            Throw New System.Exception("Error when retrieving CLASS1RAILLIST table values", SqlEx)
        End Try

        GetClass1RailList = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function GetClass1RailDataToPrepare(ByVal CurrentYear As String) As DataTable

        Dim cmdCommand As SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try

            ' Open the SQL connection using the global variable holding the connection string
            Gbl_Class1RailList_DatabaseName = Get_Database_Name_From_SQL("1", "CLASS1RAILLIST")
            Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "CLASS1RAILLIST")
            OpenSQLConnection(Gbl_Class1RailList_DatabaseName)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "SELECT *, (SELECT TOP 1 RRICC FROM " & Gbl_Class1RailList_TableName & " WHERE REGION_ID = R.REGION_ID And RRICC > 900000) RegionRRICC," &
                                               "(Select TOP 1 RRICC FROM " & Gbl_Class1RailList_TableName & " WHERE REGION_ID = 0 And RRICC > 900000) NationRRICC " &
                                     "FROM " & Gbl_Class1RailList_TableName & " R WHERE EFFECTIVE_YEAR <=" & CurrentYear & " And EXPIRATION_YEAR > " & CurrentYear
            cmdCommand.Connection = gbl_SQLConnection

            Logger.Log("GetClass1RailDataToPrepare : " + cmdCommand.CommandText)
            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "Class1"

        Catch SqlEx As SqlException
            Throw New System.Exception("Error When retrieving CLASS1RAILLIST table values", SqlEx)
        End Try

        GetClass1RailDataToPrepare = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function GetCustomData(ByVal TableName As String) As DataTable

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        cmdCommand.CommandText = Nothing

        Try

            'Get the database_name and table_name for the referenced/passed value from the database
            gbl_Table_Name = Get_Table_Name_From_SQL("1", TableName)
            gbl_Database_Name = Get_Database_Name_From_SQL("1", TableName)

            OpenSQLConnection(gbl_Database_Name)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "Select * FROM [" & gbl_Table_Name & "]"
            Logger.Log("GetCustomData: " & cmdCommand.CommandText) ' <--- Add this line
            cmdCommand.Connection = gbl_SQLConnection

            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = TableName

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL Error When executing " & cmdCommand.CommandText, SqlEx)
        End Try

        GetCustomData = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    ''' <summary>
    ''' Gets the U_AVALUES table from SQL
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAValue(ByVal RailRoadNumber As Integer) As DataTable

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try
            'Get the database_name and table_name for the referenced/passed value from the database
            Gbl_AValues_TableName = Get_Table_Name_From_SQL(My.Settings.CurrentYear, "AValues")
            Gbl_AValues_DatabaseName = Get_Database_Name_From_SQL(My.Settings.CurrentYear, "AValues")
            Gbl_ACode_TableName = Get_Table_Name_From_SQL("1", "ACodes")
            Gbl_ACode_DatabaseName = Get_Database_Name_From_SQL("1", "ACodes")
            OpenSQLConnection(Gbl_AValues_DatabaseName)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "Select at.[Year],at.[aCode_id],at.[Value],ac.[aLine], ac.[Rpt_sheet], ac.[aColumn] FROM [" &
                Gbl_AValues_TableName & "] at (NOLOCK) JOIN [" & Gbl_ACode_DatabaseName & "].[dbo].[" &
                Gbl_ACode_TableName & "] ac (NOLOCK) On at.aCode_id = ac.aCode_id WHERE RR_Id = " &
                RailRoadNumber & " ORDER BY acode_id"
            cmdCommand.Connection = gbl_SQLConnection

            Logger.Log("GetAValue : " + cmdCommand.CommandText)

            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "AValues"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL Error When retrieving U_AVALUES. SQL statement: " & cmdCommand.CommandText, SqlEx)
        End Try

        GetAValue = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    ''' <summary>
    ''' Gets the U_AVALUES table from SQL where RR_ID = 0
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAValue0_RR() As DataTable

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try
            'Get the database_name and table_name for the referenced/passed value from the database
            Gbl_AValues_TableName = Get_Table_Name_From_SQL(My.Settings.CurrentYear, "AValues")
            Gbl_AValues_DatabaseName = Get_Database_Name_From_SQL(My.Settings.CurrentYear, "AValues")
            Gbl_ACode_TableName = Get_Table_Name_From_SQL("1", "ACodes")
            Gbl_ACode_DatabaseName = Get_Database_Name_From_SQL("1", "ACodes")
            OpenSQLConnection(Gbl_AValues_DatabaseName)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "Select at.[Year],at.[aCode_id],at.[Value],ac.[aLine], ac.[Rpt_sheet] FROM [" & Gbl_AValues_TableName & "] at " &
                                     "(NOLOCK) JOIN [" & Gbl_Controls_Database_Name & "].[dbo].[" & Gbl_ACode_TableName &
                                     "] ac (NOLOCK) on at.aCode_id = ac.aCode_id WHERE RR_Id = 0 ORDER BY acode_id"
            cmdCommand.Connection = gbl_SQLConnection
            Logger.Log("GetAValue0_RR: " & cmdCommand.CommandText) ' <--- Add this line
            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "aValues"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when retrieving U_AVALUES.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        End Try

        GetAValue0_RR = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function GetAValueRegion_RR(ByVal RailRoadNumber As Integer) As DataTable

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try

            ' Open the SQL connection using the global variable holding the connection string
            OpenSQLConnection(Get_Database_Name_From_SQL(My.Settings.CurrentYear.ToString, "AVALUES"))

            'Get the database_name and table_name for the referenced/passed value from the database
            'Get the database_name and table_name for the referenced/passed value from the database
            Gbl_AValues_TableName = Get_Table_Name_From_SQL(My.Settings.CurrentYear, "AValues")
            Gbl_AValues_DatabaseName = Get_Database_Name_From_SQL(My.Settings.CurrentYear, "AValues")
            Gbl_ACode_TableName = Get_Table_Name_From_SQL("1", "ACodes")
            Gbl_ACode_DatabaseName = Get_Database_Name_From_SQL("1", "ACodes")
            Gbl_Region_TableName = Get_Table_Name_From_SQL("1", "Region") 'this is in the Controls database
            Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "Class1RailList") 'this is in the Controls database
            OpenSQLConnection(Gbl_AValues_DatabaseName)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "SELECT at.[Year],at.[aCode_id],at.[Value],ac.[aLine], ac.[Rpt_sheet], ac.[Code] " &
                                     "FROM [" & Gbl_AValues_TableName & "] at (NOLOCK) " &
                                     "JOIN [" & Gbl_ACode_DatabaseName & "].[dbo].[" & Gbl_ACode_TableName & "] ac (NOLOCK) on at.aCode_id = ac.aCode_id " &
                                     "WHERE RR_Id = (SELECT TOP 1 RR_ID FROM [" & Gbl_Controls_Database_Name & "].[dbo]." & Gbl_Class1RailList_TableName & " " &
                                                    "WHERE SHORT_NAME = (SELECT TOP 1 Description " &
                                                         "FROM [" & Gbl_Controls_Database_Name & "].[dbo]." & Gbl_Region_TableName & " reg INNER JOIN [" &
                                                         Gbl_Controls_Database_Name & "].[dbo]." & Gbl_Class1RailList_TableName & " rr ON reg.id = rr.REGION_ID " &
                                                         "WHERE RR_ID = " & RailRoadNumber & ")) ORDER BY acode_id"
            cmdCommand.Connection = gbl_SQLConnection
            Logger.Log("GetAValueRegion_RR: " & cmdCommand.CommandText) ' <--- Add this line
            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "aValues"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when retrieving U_AVALUES.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        End Try

        GetAValueRegion_RR = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function GetPriceIndexes(ByVal RailRoadNumber As Integer, ByVal Year As String) As DataTable

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try

            ' Open the SQL connection using the global variable holding the connection string
            Gbl_Price_Index_TableName = Get_Table_Name_From_SQL("1", "Index") 'this is in the Controls database
            Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "Class1RailList") 'this is in the Controls database
            OpenSQLConnection(Gbl_Controls_Database_Name)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "SELECT * " &
                                     "FROM [" & Gbl_Price_Index_TableName & "] p " &
                                     "INNER JOIN [" & Gbl_Class1RailList_TableName & "] rr ON p.Region = rr.REGION_ID " &
                                     "WHERE rr.RR_ID = " & RailRoadNumber & " AND YEAR = " & Year
            cmdCommand.Connection = gbl_SQLConnection
            Logger.Log("GetPriceIndexes: " & cmdCommand.CommandText) ' <--- Add this line
            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "PriceIndex"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when retrieving U_PRICE_INDEX.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        End Try

        GetPriceIndexes = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function GetCarTypeStatistics(ByVal RailRoadNumber As Integer) As DataTable

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try

            Gbl_Op_Stats_By_Car_Type_TableName = Get_Table_Name_From_SQL("1", "Op_Stats_By_Car_Type") 'this is in the Controls database
            Gbl_Region_TableName = Get_Table_Name_From_SQL("1", "Region") 'this is in the Controls database
            Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "Class1RailList") 'this is in the Controls database
            OpenSQLConnection(Gbl_Controls_Database_Name)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "SELECT [Line],[C1],[C2],[C3],[C4],[C5],[C6],[C7],[C8],[C9],[C10],[C11] FROM [" &
                                     Gbl_Op_Stats_By_Car_Type_TableName & "] c " &
                                     "INNER JOIN [" & Gbl_Region_TableName & "] r ON c.Region = r.description " &
                                     "INNER JOIN [" & Gbl_Class1RailList_TableName & "] rr ON r.id = rr.REGION_ID " &
                                     "WHERE rr.RR_ID = " & RailRoadNumber
            cmdCommand.Connection = gbl_SQLConnection
            Logger.Log("GetCarTypeStatistics: " & cmdCommand.CommandText) ' <--- Add this line
            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "CarType"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when retrieving R_OP_STATS_BY_CAR_TYPE.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        End Try

        GetCarTypeStatistics = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function GetCarTypeStatisticsPart2(ByVal RailRoadNumber As Integer) As DataTable

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try

            ' Open the SQL connection using the global variable holding the connection string

            Gbl_Op_Stats_By_Car_Type2_TableName = Get_Table_Name_From_SQL("1", "Op_Stats_By_Car_Type_2") 'this is in the Controls database
            Gbl_Region_TableName = Get_Table_Name_From_SQL("1", "Region") 'this is in the Controls database
            Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "Class1RailList") 'this is in the Controls database
            OpenSQLConnection(Gbl_Controls_Database_Name)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "SELECT [Line],[C1],[C2],[C3],[C4],[C5],[C6],[C7],[C8],[C9],[C10],[C11],[C12],[C13],[C14] " &
                                     "FROM [" & Gbl_Op_Stats_By_Car_Type2_TableName & "] c " &
                                     "INNER JOIN [" & Gbl_Region_TableName & "] r ON c.Region = r.description " &
                                     "INNER JOIN [" & Gbl_Class1RailList_TableName & "] rr ON r.id = rr.REGION_ID " &
                                     "WHERE rr.RR_ID = " & RailRoadNumber

            cmdCommand.Connection = gbl_SQLConnection
            Logger.Log("GetCarTypeStatisticsPart2: " & cmdCommand.CommandText) ' <--- Add this line

            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "CarTypePart2"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when retrieving R_OP_STATS_BY_CAR_TYPE_2.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        End Try

        GetCarTypeStatisticsPart2 = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function GetCarTypeStatisticsPart3() As DataTable

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try

            Gbl_Op_Stats_By_Car_Type3_TableName = Get_Table_Name_From_SQL("1", "Op_Stats_By_Car_Type_3") 'this is in the Controls database
            OpenSQLConnection(Gbl_Controls_Database_Name)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "SELECT [Line],[C1],[C2],[C3],[C4],[C5] FROM [" &
                Gbl_Op_Stats_By_Car_Type3_TableName & "] ORDER BY Line"
            cmdCommand.Connection = gbl_SQLConnection
            Logger.Log("GetCarTypeStatisticsPart3: " & cmdCommand.CommandText) ' <--- Add this line
            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "CarTypePart3"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when retrieving R_OP_STATS_BY_CAR_TYPE_3.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        End Try

        GetCarTypeStatisticsPart3 = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function GetLineSourceText() As DataTable

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try
            Gbl_Line_Source_Text_TableName = Get_Table_Name_From_SQL("1", "Line_Source_Text") 'this is in the Controls database
            OpenSQLConnection(Gbl_Controls_Database_Name)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "SELECT * FROM " & Gbl_Line_Source_Text_TableName & " ORDER BY Rpt_sheet,Line"
            cmdCommand.Connection = gbl_SQLConnection

            Logger.Log("GetLineSourceText: " & cmdCommand.CommandText) ' <--- Add this line

            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "aCode"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when retrieving U_LINE_SOURCE_TEXT.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        End Try

        GetLineSourceText = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function GetDataDictionary(ByVal CurrentYear As Integer) As DataTable

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try
            OpenSQLConnection(Get_Database_Name_From_SQL("1", "Data_Dictionary"))

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "SELECT SUBSTRING(WTALL,1,6) Line, MAX(AnnPeriod) AnnPeriod " &
                                     "FROM " & Get_Table_Name_From_SQL("1", "Data_Dictionary") & " " &
                                     "WHERE Effective_Dt <= " & CurrentYear.ToString() & " " &
                                     "AND Expiration_Dt > " & CurrentYear.ToString() & " " &
                                     "GROUP BY SUBSTRING(WTALL,1,6) ORDER BY 1"

            cmdCommand.Connection = gbl_SQLConnection
            Logger.Log("GetDataDictionary (Integer): " & cmdCommand.CommandText) ' <--- Add this line
            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "DataDictionary"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when retrieving DATA_DICTIONARY.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        End Try

        GetDataDictionary = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function GetEcodes() As DataTable

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try
            OpenSQLConnection(Get_Database_Name_From_SQL("1", "ECODES"))

            Gbl_ECode_TableName = Get_Table_Name_From_SQL("1", "ECodes")

            dsDataSet.Clear()

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "Select * from " & Gbl_ECode_TableName ' simple select query for ecode data
            cmdCommand.Connection = gbl_SQLConnection

            Logger.Log("GetEcodes: " & cmdCommand.CommandText) ' <--- Add this line
            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "EcodeData"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when retrieving UR_ECODE. SQL statement: " & cmdCommand.CommandText, SqlEx)
        End Try

        GetEcodes = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    ''' <summary>
    ''' Handles error trapping and writes the message and stack trace back to SQL
    ''' </summary>
    ''' <param name="message"></param>
    ''' <param name="stack"></param>
    ''' <param name="location"></param>
    ''' <remarks></remarks>
    Public Sub HandleError(ByVal CurrentYear As String, ByVal data As String, ByVal message As String, ByVal stack As String, ByVal location As String)

        'write the error to SQL
        Dim mSQLstr As String
        Dim Timestamp As String

        Timestamp = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fffffff tt")
        message = message.Replace("'", "").Replace("""", "")
        stack = stack.Replace("'", "").Replace("""", "")

        gbl_Database_Name = Get_Database_Name_From_SQL(My.Settings.CurrentYear.ToString, "ERRORS")
        Gbl_Errors_TableName = Get_Table_Name_From_SQL(My.Settings.CurrentYear.ToString, "ERRORS")
        'OpenADOConnection(gbl_Database_Name)

        'mSQLstr = "INSERT INTO " & Gbl_Errors_TableName & " VALUES ('" & data.ToString & "','" & Timestamp.ToString & "','" & message.ToString & "','" & stack.ToString & "','" & location.ToString & "')"
        'gbl_ADOConnection.Execute(mSQLstr)

    End Sub

    Public Sub ClearSubstitutions(ByVal Year As String, ByVal RRid As String)

        Dim mSQLStr As String

        ' Open the SQL connection to the work year's database
        gbl_Database_Name = Get_Database_Name_From_SQL(My.Settings.CurrentYear.ToString, "SUBSTITUTIONS")
        Gbl_Substitutions_TableName = Get_Table_Name_From_SQL(My.Settings.CurrentYear.ToString, "SUBSTITUTIONS")
        OpenADOConnection(gbl_Database_Name)

        mSQLStr = "DELETE FROM " & Gbl_Substitutions_TableName & " WHERE Year = '" & Year & "' AND RR_ID = " & RRid
        Logger.Log("ClearSubstitutions: " & mSQLStr) ' <--- Add this line

        'Comment out deletion
        'gbl_ADOConnection.Execute(mSQLStr)

    End Sub

    Public Sub InsertSubstitutions(ByVal Year As String, ByVal RRid As String, ByVal Values As System.Array)

        Dim cnConnection As New SqlConnection
        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim eCode, Value As String

        Try
            ' Clear/delete the substitutions records for previous runs
            ClearSubstitutions(Year, RRid)

            ' Get the table names for the Substitutions and ECode tables
            Gbl_Substitutions_TableName = Get_Table_Name_From_SQL(Year.ToString, "SUBSTITUTIONS")
            Gbl_ECode_TableName = Get_Table_Name_From_SQL("1", "ECODES")

            ' Open the SQL connection using the global variable holding the connection string
            gbl_Database_Name = Get_Database_Name_From_SQL(My.Settings.CurrentYear.ToString, "SUBSTITUTIONS")
            OpenSQLConnection(gbl_Database_Name)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.Connection = gbl_SQLConnection

            daAdapter = New SqlClient.SqlDataAdapter

            For i = 1 To Values.GetUpperBound(0)

                eCode = Values(i, 1)
                Value = Values(i, 2)

                cmdCommand.CommandText = "INSERT INTO " & Gbl_Substitutions_TableName &
                                         " (Year, RR_id, eCode_id, Value, Final_Value, entry_dt) VALUES ('" &
                                         Year & "', " & RRid & ", (SELECT TOP 1 eCode_id FROM " &
                                         Gbl_Controls_Database_Name & ".dbo." & Gbl_ECode_TableName & " WHERE eCode = '" &
                                         eCode & "'), " &
                                         IIf(Value = "n/a", 0, Value) & "," & IIf(Value = "n/a", 0, Value) & ", '" &
                                         DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") & "')"
                Logger.Log("InsertSubstitutions: " & cmdCommand.CommandText) ' <--- Add this line
                daAdapter.InsertCommand = cmdCommand
                daAdapter.InsertCommand.CommandTimeout = 60
                'Comment out insertion
                'daAdapter.InsertCommand.ExecuteNonQuery()
            Next

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when inserting SUBSTITUTIONS records.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        Finally
            'if we failed or succedded, close the sql connection
            If cnConnection.State = ConnectionState.Open Then
                cnConnection.Close()
            End If
        End Try

    End Sub

    Public Sub RunSubstitutions(ByVal Year As String, ByVal RRIDs As System.Array)

        Dim cmdCommand As SqlCommand
        Dim ThisCmdString As New StringBuilder

        Try

            ' Get the table names for the Substitutions and ECode tables
            Gbl_Substitutions_TableName = Get_Table_Name_From_SQL(Year.ToString, "SUBSTITUTIONS")
            Gbl_ECode_TableName = Get_Table_Name_From_SQL("1", "ECODES")

            ' Open the SQL connection using the global variable holding the connection string
            OpenSQLConnection(Get_Database_Name_From_SQL(My.Settings.CurrentYear.ToString, "SUBSTITUTIONS"))

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.Connection = gbl_SQLConnection

            Array.Sort(RRIDs)
            Array.Reverse(RRIDs)
            For i = 0 To RRIDs.GetUpperBound(0)
                'ThisCmdString.Append("DECLARE @NotesText CHAR(55) = 'No substitution - initial value within reasonable range';")
                'ThisCmdString.Append("DECLARE @RegionID TINYINT;")
                'ThisCmdString.Append("DECLARE @RR_NAME VARCHAR(5);")
                'ThisCmdString.Append("DECLARE @Region_NAME VARCHAR(5);")
                'ThisCmdString.Append("SELECT @RegionID = CASE WHEN RR.RR_ID = RG.RR_ID THEN RG.Other_RR_ID ELSE RG.RR_ID END ")
                'ThisCmdString.Append("FROM URCS_Controls.dbo.R_Class1_Rail_List RR INNER JOIN ")
                'ThisCmdString.Append("(SELECT Region_ID, RR_ID, (SELECT TOP 1 RR_ID FROM URCS_Controls.dbo.R_Class1_Rail_List WHERE RRICC > 900000 AND RR_ID > 0 AND RR_ID <> S.RR_ID) Other_RR_ID ")
                'ThisCmdString.Append("FROM URCS_Controls.dbo.R_Class1_Rail_List S WHERE RRICC > 900000 AND RR_ID > 0) ")
                'ThisCmdString.Append("RG ON RR.REGION_ID = RG.REGION_ID WHERE RR.RR_ID = @RR_ID" & vbCrLf)
                'ThisCmdString.Append("SELECT TOP 1 @RR_NAME = SHORT_NAME FROM URCS_Controls.dbo.R_Class1_Rail_List WHERE RR_ID = @RR_ID" & vbCrLf)
                'ThisCmdString.Append("SELECT TOP 1 @Region_NAME = SHORT_NAME FROM URCS_Controls.dbo.R_Class1_Rail_List WHERE RR_ID = @RegionID" & vbCrLf)
                '' Run rules 1 through 13
                'ThisCmdString.Append("UPDATE OriginalValues SET OriginalValues.Final_Value = ReplaceValues.Final_Value, ")
                'ThisCmdString.Append("Notes = 'Updated with ' + ReplaceCodes.eCode + ' from ' + @RR_NAME + ' (RR_ID = ' + CAST(ReplaceValues.RR_id as VARCHAR(10)) + ')' ")
                'ThisCmdString.Append("FROM (U_Substitutions OriginalValues INNER JOIN URCS_Controls.dbo.R_ECodes OriginalCodes ")
                'ThisCmdString.Append("ON OriginalValues.eCode_id = OriginalCodes.eCode_ID AND OriginalCodes.eLine=201) INNER JOIN ")
                'ThisCmdString.Append("(U_Substitutions ReplaceValues INNER JOIN URCS_Controls.dbo.R_ECodes ReplaceCodes ")
                'ThisCmdString.Append("ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID AND ReplaceCodes.eLine=202) ")
                'ThisCmdString.Append("ON OriginalValues.RR_id = ReplaceValues.RR_id AND OriginalValues.Year = ReplaceValues.Year AND OriginalCodes.eCode = ")
                'ThisCmdString.Append("REPLACE(ReplaceCodes.eCode,'2C','1C') ")
                'ThisCmdString.Append("WHERE OriginalValues.RR_id = @RR_ID AND OriginalValues.Year = @Year AND OriginalCodes.ePart = 'E1'" & vbCrLf)
                '' Run rules 14 through 17
                'ThisCmdString.Append("UPDATE OriginalValues SET OriginalValues.Final_Value = ReplaceValues.Final_Value, ")
                'ThisCmdString.Append("Notes = 'Updated with ' + ReplaceCodes.eCode + ' from ' + @RR_NAME + ' (RR_ID = ' + CAST(ReplaceValues.RR_id as VARCHAR(10)) + ')' ")
                'ThisCmdString.Append("FROM (U_Substitutions OriginalValues INNER JOIN URCS_Controls.dbo.R_ECodes OriginalCodes ")
                'ThisCmdString.Append("ON OriginalValues.eCode_id = OriginalCodes.eCode_ID AND OriginalCodes.eLine=101) INNER JOIN ")
                'ThisCmdString.Append("(U_Substitutions ReplaceValues INNER JOIN URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID AND ReplaceCodes.eLine=102) ")
                'ThisCmdString.Append("ON OriginalValues.RR_id = ReplaceValues.RR_id AND OriginalValues.Year = ReplaceValues.Year AND OriginalCodes.eCode = REPLACE(ReplaceCodes.eCode,'2C','1C') ")
                'ThisCmdString.Append("WHERE OriginalValues.RR_id = @RR_ID AND OriginalValues.Year = @Year AND OriginalCodes.ePart = 'E2' AND OriginalCodes.eColumn IN (2,3,4,24)" & vbCrLf)
                '' Run rules 18 through 38
                'ThisCmdString.Append("UPDATE OriginalValues SET OriginalValues.Final_Value = CASE WHEN OriginalValues.Final_Value <= 0 THEN ReplaceValues.Final_Value ELSE OriginalValues.Final_Value END, ")
                'ThisCmdString.Append("Notes = CASE WHEN OriginalValues.Final_Value <= 0 THEN 'Updated with ' + ReplaceCodes.eCode + ' from ' + @Region_NAME + ' ")
                'ThisCmdString.Append("(RR_ID = ' + CAST(ReplaceValues.RR_id as VARCHAR(10)) + ')' ")
                'ThisCmdString.Append("ELSE CASE WHEN OriginalValues.Notes IS NULL THEN @NotesText ELSE OriginalValues.Notes END END ")
                'ThisCmdString.Append("FROM (U_Substitutions OriginalValues INNER JOIN URCS_Controls.dbo.R_ECodes OriginalCodes ON OriginalValues.eCode_id = OriginalCodes.eCode_ID) INNER JOIN ")
                'ThisCmdString.Append("(U_Substitutions ReplaceValues INNER JOIN URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID) ")
                'ThisCmdString.Append("ON OriginalValues.Year = ReplaceValues.Year AND OriginalCodes.eCode_id = ReplaceCodes.eCode_id ")
                'ThisCmdString.Append("WHERE OriginalValues.RR_id = @RR_ID AND ReplaceValues.RR_id = @RegionID AND OriginalValues.Year = @Year ")
                'ThisCmdString.Append("AND OriginalCodes.eLine BETWEEN 200 AND 300 AND OriginalCodes.eColumn = 13 AND OriginalCodes.ePart = 'E1'" & vbCrLf)
                '' Run rules 39 through 56
                'ThisCmdString.Append("UPDATE OriginalValues SET OriginalValues.Final_Value = ")
                'ThisCmdString.Append("CASE WHEN OriginalValues.Final_Value < 1 OR OriginalValues.Final_Value > 3 ")
                'ThisCmdString.Append("THEN ReplaceValues.Final_Value ")
                'ThisCmdString.Append("ELSE OriginalValues.Final_Value END, ")
                'ThisCmdString.Append("Notes = CASE WHEN OriginalValues.Final_Value < 1 OR OriginalValues.Final_Value > 3 ")
                'ThisCmdString.Append("THEN 'Updated with ' + ReplaceCodes.eCode + ' from ' + @Region_NAME + ' (RR_ID = ' + CAST(ReplaceValues.RR_id as VARCHAR(10)) + ')' ")
                'ThisCmdString.Append("ELSE CASE WHEN OriginalValues.Notes IS NULL THEN @NotesText ELSE OriginalValues.Notes END END ")
                'ThisCmdString.Append("FROM (U_Substitutions OriginalValues INNER JOIN URCS_Controls.dbo.R_ECodes OriginalCodes ON OriginalValues.eCode_id = OriginalCodes.eCode_ID) INNER JOIN ")
                'ThisCmdString.Append("(U_Substitutions ReplaceValues INNER JOIN URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID) ")
                'ThisCmdString.Append("ON OriginalValues.Year = ReplaceValues.Year AND OriginalCodes.eCode_id = ReplaceCodes.eCode_id ")
                'ThisCmdString.Append("WHERE OriginalValues.RR_id = @RR_ID AND ReplaceValues.RR_id = @RegionID AND OriginalValues.Year = @Year ")
                'ThisCmdString.Append("AND OriginalCodes.eLine BETWEEN 100 AND 200 AND OriginalCodes.eColumn = 3 AND OriginalCodes.ePart = 'E2' " & vbCrLf)
                '' Run rules 58 through 59
                'ThisCmdString.Append("UPDATE OriginalValues SET OriginalValues.Final_Value = ReplaceValues.Final_Value, ")
                'ThisCmdString.Append("Notes = 'Updated with ' + ReplaceCodes.eCode + ' from ' + @RR_NAME + ' (RR_ID = ' + CAST(ReplaceValues.RR_id as VARCHAR(10)) + ')' ")
                'ThisCmdString.Append("FROM (U_Substitutions OriginalValues INNER JOIN ")
                'ThisCmdString.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON OriginalValues.eCode_id = OriginalCodes.eCode_ID AND OriginalCodes.eColumn = 2) INNER JOIN ")
                'ThisCmdString.Append("(U_Substitutions ReplaceValues INNER JOIN ")
                'ThisCmdString.Append("URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID AND ReplaceCodes.eColumn = 13 ) ")
                'ThisCmdString.Append("ON OriginalValues.Year = ReplaceValues.Year AND OriginalValues.RR_id = ReplaceValues.RR_id AND OriginalCodes.eLine = ReplaceCodes.eLine ")
                'ThisCmdString.Append("WHERE OriginalValues.RR_id = @RR_ID AND OriginalValues.Year = @Year ")
                'ThisCmdString.Append("AND OriginalCodes.eLine IN (215,216) AND OriginalCodes.ePart = 'E1'" & vbCrLf)
                '' Run rules 61 through 113 (excluding 87)
                'ThisCmdString.Append("UPDATE OriginalValues SET OriginalValues.Final_Value = ReplaceValues.Final_Value, ")
                'ThisCmdString.Append("Notes = 'Updated with ' + ReplaceCodes.eCode + ' from ' + @RR_NAME + ' (RR_ID = ' + CAST(ReplaceValues.RR_id as VARCHAR(10)) + ')' ")
                'ThisCmdString.Append("FROM (U_Substitutions OriginalValues INNER JOIN ")
                'ThisCmdString.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON OriginalValues.eCode_id = OriginalCodes.eCode_ID AND OriginalCodes.eLine IN (115,116)) INNER JOIN ")
                'ThisCmdString.Append("(U_Substitutions ReplaceValues INNER JOIN ")
                'ThisCmdString.Append("URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID AND ReplaceCodes.eLine = 117) ")
                'ThisCmdString.Append("ON OriginalValues.Year = ReplaceValues.Year AND OriginalValues.RR_id = ReplaceValues.RR_id AND ")
                'ThisCmdString.Append("OriginalCodes.eColumn = ReplaceCodes.eColumn And OriginalCodes.EPart = ReplaceCodes.EPart ")
                'ThisCmdString.Append("WHERE OriginalValues.RR_id = @RR_ID AND OriginalValues.Year = @Year AND OriginalCodes.ePart = 'E2' ")
                'ThisCmdString.Append("AND (OriginalCodes.eColumn = 2 OR OriginalCodes.eColumn BETWEEN 5 AND 29) " & vbCrLf)
                '' Run rule 114
                'ThisCmdString.Append("UPDATE OriginalValues SET OriginalValues.Final_Value = ")
                'ThisCmdString.Append("CASE WHEN OriginalValues.Final_Value = 1 THEN ReplaceValues.Final_Value ELSE OriginalValues.Final_Value END, ")
                'ThisCmdString.Append("Notes = CASE WHEN OriginalValues.Final_Value = 1 ")
                'ThisCmdString.Append("THEN 'Updated with ' + ReplaceCodes.eCode + ' from ' + @RR_NAME + ' (RR_ID = ' + CAST(ReplaceValues.RR_id as VARCHAR(10)) + ')' ")
                'ThisCmdString.Append("ELSE CASE WHEN OriginalValues.Notes IS NULL THEN @NotesText ELSE OriginalValues.Notes END END ")
                'ThisCmdString.Append("FROM (U_Substitutions OriginalValues INNER JOIN ")
                'ThisCmdString.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON OriginalValues.eCode_id = OriginalCodes.eCode_ID AND OriginalCodes.eColumn = 8) INNER JOIN ")
                'ThisCmdString.Append("(U_Substitutions ReplaceValues INNER JOIN ")
                'ThisCmdString.Append("URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID AND ReplaceCodes.eColumn = 4) ")
                'ThisCmdString.Append("ON OriginalValues.Year = ReplaceValues.Year AND OriginalValues.RR_id = ReplaceValues.RR_id AND OriginalCodes.eLine = ReplaceCodes.eLine ")
                'ThisCmdString.Append("WHERE OriginalValues.RR_id = @RR_ID AND OriginalValues.Year = @Year AND OriginalCodes.ePart = 'E2' AND OriginalCodes.eLine = 111" & vbCrLf)
                '' Run rule 115
                'ThisCmdString.Append("UPDATE OriginalValues SET OriginalValues.Final_Value = CASE WHEN OriginalValues.Final_Value <> 4162 THEN 4162 ELSE OriginalValues.Final_Value END, ")
                'ThisCmdString.Append("Notes = CASE WHEN OriginalValues.Final_Value <> 4162 THEN 'Updated with value 4162' ELSE @NotesText END ")
                'ThisCmdString.Append("FROM U_Substitutions OriginalValues INNER JOIN ")
                'ThisCmdString.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON OriginalValues.eCode_id = OriginalCodes.eCode_ID AND OriginalCodes.eColumn = 23 ")
                'ThisCmdString.Append("WHERE OriginalValues.RR_id = @RR_ID AND OriginalValues.Year = @Year AND OriginalCodes.ePart = 'E2' AND OriginalCodes.eLine = 111" & vbCrLf)
                'cmdCommand.CommandText = ThisCmdString.ToString
                cmdCommand.CommandText = "EXECUTE usp_RunSubstitutions '" & Year & "'," & RRIDs(i)
                Logger.Log("RunSubstitutions: " & cmdCommand.CommandText) ' <--- Add this line
                'Comment out substitution
                'cmdCommand.ExecuteNonQuery()
                cmdCommand.CommandTimeout = 60
            Next i

        Catch ex As System.Exception
            'if we get an error toss it to the web app
            Throw (ex)
        Finally
            'if we failed or succedded, close the sql connection
            If gbl_SQLConnection.State = ConnectionState.Open Then
                gbl_SQLConnection.Close()
            End If
        End Try

    End Sub

    Public Sub CreateEValues(ByVal Year As String)

        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter

        Try
            ' Have to delete the old ones first
            ClearEValues(Year)

            ' Get the table names for the Substitutions and ECode tables
            Gbl_Substitutions_TableName = Get_Table_Name_From_SQL(Year.ToString, "SUBSTITUTIONS")
            Gbl_EValues_TableName = Get_Table_Name_From_SQL(Year.ToString, "EVALUES")

            ' Open the SQL connection using the global variable holding the connection string
            gbl_Database_Name = Get_Database_Name_From_SQL(My.Settings.CurrentYear.ToString, "EVALUES")
            OpenSQLConnection(gbl_Database_Name)

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.Connection = gbl_SQLConnection

            daAdapter = New SqlClient.SqlDataAdapter

            cmdCommand.CommandText = "INSERT INTO " & Gbl_EValues_TableName & " " &
                                     "SELECT Year, RR_Id, eCode_id, Final_Value, entry_dt " &
                                     "FROM " & Gbl_Substitutions_TableName & " WHERE Year = '" & Year & "'"

            daAdapter.InsertCommand = cmdCommand
            Logger.Log("CreateEValues: " & cmdCommand.CommandText) ' <--- Add this line
            'Comment out insertion
            daAdapter.InsertCommand.CommandTimeout = 60
            'daAdapter.InsertCommand.ExecuteNonQuery()

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when writing E_VALUES.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        Finally
            'if we failed or succedded, close the sql connection
            If gbl_SQLConnection.State = ConnectionState.Open Then
                gbl_SQLConnection.Close()
            End If
        End Try

        daAdapter.Dispose()
        cmdCommand.Dispose()

    End Sub

    Public Sub ClearEValues(ByVal Year As String)

        Dim mStrSQL As String

        Try
            ' Get the table name
            Gbl_EValues_TableName = Get_Table_Name_From_SQL(Year.ToString, "EVALUES")

            ' Open the SQL connection using the global variable holding the connection string
            gbl_Database_Name = Get_Database_Name_From_SQL(Year, "EVALUES")
            OpenADOConnection(gbl_Database_Name)

            mStrSQL = "DELETE FROM " & Gbl_EValues_TableName & " WHERE Year = '" & Year & "'"
            Logger.Log("ClearEValues: " & mStrSQL) ' <--- Add this line

            'Comment out deletion
            'gbl_ADOConnection.Execute(mStrSQL)

        Catch SqlEx As SqlException
            Throw New System.Exception("Error when deleting records from U_EVALUES table", SqlEx)
        End Try

    End Sub

    Public Function GetTransTable(ByVal CurrentYear As String) As DataTable

        Dim cnConnection As New SqlConnection
        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try

            ' Locate the Trans database
            Gbl_Trans_DatabaseName = Get_Database_Name_From_SQL("1", "TRANS")
            Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "TRANS")

            Logger.Log("Gbl_Trans_Database = " + Gbl_Trans_DatabaseName + "; Gbl_Trans_TableName = " + Gbl_Trans_TableName)
            ' Open the SQL connection
            cnConnection = New SqlConnection(LoadSQLConnectionStringForDB(Gbl_Trans_DatabaseName))
            cnConnection.Open()

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "SELECT * FROM " & Gbl_Trans_TableName & " WHERE Year between " &
                                        (Integer.Parse(CurrentYear) - 4).ToString() & " And " & CurrentYear
            cmdCommand.Connection = cnConnection

            Logger.Log("GetTransTable: " & cmdCommand.CommandText) ' <--- Add this line
            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "TRANS"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when retrieving TRANS records.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        Finally
            'if we failed or succedded, close the sql connection
            If cnConnection.State = ConnectionState.Open Then
                cnConnection.Close()
            End If
        End Try

        GetTransTable = dsDataSet.Tables(0)

        'clean up
        cnConnection.Dispose()
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function GetDataDictionary(ByVal CurrentYear As String) As DataTable

        Dim cnConnection As New SqlConnection
        Dim cmdCommand As New SqlCommand
        Dim daAdapter As SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try

            Gbl_Data_Dictionary_TableName = Get_Table_Name_From_SQL("1", "DATA_DICTIONARY")

            ' Open the SQL connection - the Data Dictionary is always in the Controls database
            cnConnection = New SqlConnection(LoadSQLConnectionStringForDB(Gbl_Controls_Database_Name))
            cnConnection.Open()

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = "SELECT * FROM " & Gbl_Data_Dictionary_TableName & " WHERE Effective_Dt <= " &
                                        CurrentYear & " And Expiration_Dt > " & CurrentYear &
                                        " ORDER BY URCSID"
            cmdCommand.Connection = cnConnection

            Logger.Log("GetDataDictionary (String): " & cmdCommand.CommandText) ' <--- Add this line

            'Call the proc and fill the dataset
            daAdapter = New SqlClient.SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "DICTIONARY"

        Catch SqlEx As SqlException
            Throw New System.Exception("SQL error when retrieving DATA_DICTIONARY records.  SQL statement: " & cmdCommand.CommandText, SqlEx)
        Finally
            'if we failed or succedded, close the sql connection
            If cnConnection.State = ConnectionState.Open Then
                cnConnection.Close()
            End If
        End Try

        GetDataDictionary = dsDataSet.Tables(0)

        'clean up
        cnConnection.Dispose()
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Sub AdjustUTransValues(ByVal col As Integer, ByVal newVal As Double, ByVal Year As String, ByVal RRICC As String, ByVal SCH As String, ByVal LINE As String)

        Dim cmdCommand As SqlCommand
        Dim daAdapter As SqlDataAdapter

        Try
            OpenSQLConnection(Get_Database_Name_From_SQL("1", "TRANS"))

            cmdCommand = New SqlClient.SqlCommand
            cmdCommand.CommandType = CommandType.Text

            'cmdCommand.CommandText = "UPDATE " & Get_Table_Name_From_SQL("1", "TRANS") & " SET C" & col.ToString() & " = " & newVal.ToString() &
            '                      "WHERE Year = " & Year & " AND SCH = " & SCH & "  AND LINE = " & LINE

            ' RRICC added to update only one railroad. 5/3/2021
            cmdCommand.CommandText = "UPDATE " & Get_Table_Name_From_SQL("1", "TRANS") & " SET C" & col.ToString() & " = " & newVal.ToString() &
                                     "WHERE Year = " & Year & " AND RRICC = " & RRICC & " AND SCH = " & SCH & "  AND LINE = " & LINE

            cmdCommand.Connection = gbl_SQLConnection

            Logger.Log("AdjustUTransValues: " & cmdCommand.CommandText) ' <--- Add this line

            daAdapter = New SqlClient.SqlDataAdapter

            daAdapter.UpdateCommand = cmdCommand
            daAdapter.UpdateCommand.CommandTimeout = 60
            daAdapter.UpdateCommand.ExecuteNonQuery()

        Catch ex As System.Exception
            Throw (ex)
        End Try

    End Sub

    Public Sub WriteURAcodeData(ByVal ACode As String, ByVal AColumn As String, ByVal ALine As String, ByVal LineA As String, ByVal APart As String, ByVal Code As String, ByVal RptSheet As String)

        Dim mSQLStr As String

        gbl_Database_Name = Get_Database_Name_From_SQL("1", "ACODES")
        Gbl_ACode_TableName = Get_Table_Name_From_SQL("1", "ACODES")
        OpenADOConnection(gbl_Database_Name)

        mSQLStr = "INSERT INTO " & Gbl_ACode_TableName & " (aCode, aColumn, aLine, LineA, APart, Code, Rpt_sheet) VALUES ('" &
                                            ACode & "', " & AColumn & ", " & ALine & ", '" &
                                            LineA & "', '" & APart & "', '" & Code & "', '" & RptSheet & "')"
        Logger.Log("WriteURAcodeData: " & mSQLStr) ' <--- Add this line

        'Comment out insertion
        'gbl_ADOConnection.Execute(mSQLStr)

    End Sub

    Public Sub writeUAValues(ByVal Year As String, ByVal RRid As String, ByVal ACodeID As String, ByVal Value As String)

        Dim mSQLStr As String

        gbl_Database_Name = Get_Database_Name_From_SQL(My.Settings.CurrentYear.ToString, "AVALUES")
        Gbl_AValues_TableName = Get_Table_Name_From_SQL(My.Settings.CurrentYear.ToString.ToString, "AVALUES")
        OpenADOConnection(gbl_Database_Name)

        mSQLStr = "INSERT INTO " & Gbl_AValues_TableName & " (Year, RR_id, aCode_id, Value, entry_dt) VALUES ('" &
                                        Year & "', " & RRid & ", " & ACodeID & ", " & Value & ", '" &
                                        DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") & "')"
        Logger.Log("writeUAValues: " & mSQLStr) ' <--- Add this line

        'Comment out insertion
        'gbl_ADOConnection.Execute(mSQLStr)

    End Sub

    Public Sub TruncateATables()

        Dim mSQLStr As String

        gbl_Database_Name = Get_Database_Name_From_SQL("1", "ACODES")
        Gbl_ACode_TableName = Get_Table_Name_From_SQL("1", "ACODES")
        OpenADOConnection(gbl_Database_Name)

        mSQLStr = "DELETE " & Gbl_ACode_TableName

        Logger.Log(" TruncateATables (1) : " + mSQLStr)

        'Comment out deletion
        'gbl_ADOConnection.Execute(mSQLStr)

        mSQLStr = "DBCC CHECKIDENT('" & Gbl_ACode_TableName & "', RESEED, 0)"

        Logger.Log(" TruncateATables (2): " + mSQLStr)
        'Comment out reseeding
        'gbl_ADOConnection.Execute(mSQLStr)

        gbl_Database_Name = Get_Database_Name_From_SQL(My.Settings.CurrentYear.ToString, "AVALUES")
        Gbl_AValues_TableName = Get_Table_Name_From_SQL(My.Settings.CurrentYear.ToString, "AVALUES")
        OpenADOConnection(gbl_Database_Name)

        mSQLStr = "DELETE " & Gbl_AValues_TableName
        Logger.Log(" TruncateATables : " + mSQLStr)
        'Comment out deletion
        'gbl_ADOConnection.Execute(mSQLStr)


    End Sub

End Class
