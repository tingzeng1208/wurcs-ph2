Imports System.Data.SqlClient
Imports System.Text
Imports Microsoft.SqlServer
Module SQL_Routines

    Function LoadSQLConnectionStringForDB(ByVal mDatabase As String) As String

        ' Purpose: Formats a valid connection string for SQLClient to a particular Database Name
        ' Returns: Formatted string
        ' Author:   Michael Sanders

        LoadSQLConnectionStringForDB = "Server=" & gbl_Server_Name & ";Database=" & mDatabase.ToString & ";Trusted_Connection = True;Connection Timeout=5"
        Logger.Log(LoadSQLConnectionStringForDB)
    End Function

    Function LoadSQLConnectionStringForServer(ByVal mServer As String) As String

        ' Purpose: Formats a valid connection string for SQLClient to a particular Server
        ' Returns: Formatted string
        ' Author:   Michael Sanders

        LoadSQLConnectionStringForServer = "Server=" & mServer & ";Database=" & Gbl_Controls_Database_Name & ";Integrated Security=True;"
        Logger.Log(LoadSQLConnectionStringForServer)

    End Function

    Function LoadADOConnectionStringForDB(ByVal mDatabase As String) As String

        ' Purpose: Formats a valid connection string for SQLClient
        ' Returns: Formatted string
        ' Author:   Michael Sanders

        ' This is slated to be removed as to not utiliza ADODB

        LoadADOConnectionStringForDB = "Provider=SQLNCLI11;Server=" & gbl_Server_Name & ";Database=" & mDatabase.ToString & ";Trusted_Connection=yes;"
        Logger.Log(LoadADOConnectionStringForDB)
    End Function

    Public Sub OpenSQLConnection(ByRef mDatabaseName As String)

        'Purpose:       Checks the state of the SQL connection and closes/opens to specified database.
        'Assumptions:   
        'Affects:       
        'Inputs:        Database name to open connection to
        'Returns:       
        'Author:        Michael Sanders

        'Does the connection exist?
        If gbl_SQLConnection Is Nothing Then
            gbl_SQLConnection = New SqlConnection
            gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(mDatabaseName)
            gbl_SQLConnection.Open()
        End If

        ' Is the connection broken?
        If gbl_SQLConnection.State = ConnectionState.Broken Then
            gbl_SQLConnection.Close()
        End If

        ' If it is opened to the wrong database, close it
        If (gbl_SQLConnection.State = ConnectionState.Open) And
            (gbl_SQLConnection.Database.ToString <> mDatabaseName) Then
            gbl_SQLConnection.Close()
        End If

        ' Open the database if it is closed
        If gbl_SQLConnection.State = ConnectionState.Closed Then
            gbl_SQLConnection = New SqlConnection(LoadSQLConnectionStringForDB(mDatabaseName))
            gbl_SQLConnection.Open()
        End If

    End Sub

    Public Sub OpenADOConnection(ByRef mDatabaseName As String)

        'Purpose:       Checks the state of the ADO connection and closes/opens to specified database.
        'Assumptions:   
        'Affects:       
        'Inputs:        Database name to open connection to
        'Returns:       
        'Author:        Michael Sanders

        ' This is slated to be removed as to not utiliza ADODB

        ' Does the connection exist?
        If gbl_ADOConnection Is Nothing Then
            gbl_ADOConnection = New ADODB.Connection
            gbl_ADOConnString = LoadADOConnectionStringForDB(mDatabaseName)
            gbl_ADOConnection.Open(gbl_ADOConnString)
        End If

        ' Is the connection broken?
        If gbl_ADOConnection.State = ConnectionState.Broken Then
            Global_Variables.gbl_ADOConnection.Close()
        End If

        ' If it is opened to the wrong database, close it
        If (Global_Variables.gbl_ADOConnection.State = ConnectionState.Open) Then
            If (Global_Variables.gbl_ADOConnection.DefaultDatabase.ToString <> mDatabaseName) Then
                Global_Variables.gbl_ADOConnection.Close()
            End If
        Else
            'Do nothing
        End If

        ' Open the database if it is closed
        If Global_Variables.gbl_ADOConnection.State = ConnectionState.Closed Then
            Global_Variables.gbl_ADOConnection = New ADODB.Connection
            Global_Variables.gbl_ADOConnString = LoadADOConnectionStringForDB(mDatabaseName)
            Global_Variables.gbl_ADOConnection.Open(Global_Variables.gbl_ADOConnString)
        End If

    End Sub

    Function SetRST() As ADODB.Recordset

        'Purpose:       Initializes an ADO recordset with optimal settings
        'Assumptions:   
        'Affects:       
        'Inputs:        
        'Returns:       Initialized empty recordset
        'Author:        Michael Sanders

        SetRST = New ADODB.Recordset
        SetRST.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        SetRST.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        SetRST.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic

    End Function
    Function Verify_SQL_Running(ByRef mServer As String) As Boolean
        Dim connectString As String = ""
        Dim objConn As New SqlConnection

        connectString = LoadSQLConnectionStringForServer(mServer)

        Try
            objConn.ConnectionString = connectString
            objConn.Open()
            objConn.Close()
            Verify_SQL_Running = True
        Catch ex As Exception
            Verify_SQL_Running = False
        End Try

    End Function

    Function Get_Table_Name_From_SQL(
                       ByRef mYear As String,
                       ByRef mData_Type As String) As String

        'Purpose:       Gets the table name value from Table Locator table in URCS Controls database.
        'Assumptions:   Location of controls database from application settings
        'Affects:       Global variable for table name
        'Inputs:        Year value and data type value
        'Returns:       Table Name
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Dim mDataTable As DataTable
        Dim mSQLstr As String

        ' Set the database name to the Controls database value
        Gbl_Controls_Database_Name = My.Settings.Controls_DB

        ' Open the SQL connection
        OpenSQLConnection(Gbl_Controls_Database_Name)

        ' Build the SQL statement
        mSQLstr = "SELECT Table_Name FROM U_TABLE_LOCATOR WHERE Year = " & mYear & " AND Data_Type = '" & UCase(mData_Type) & "'"

        Logger.Log("Get_Table_Name_From_SQL: " + mSQLstr)

        mDataTable = New DataTable

        ' Fill the datatable from SQL - this should be a unique record
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        'Check to make sure we have something to return.  If not display the error message
        If mDataTable.Rows.Count = 0 Then
            MsgBox("No entry found for year " & mYear & ", Data_Type " & mData_Type, vbOKOnly, "ERROR!")
            Application.Exit()
        End If

        ' Return the value
        Get_Table_Name_From_SQL = mDataTable.Rows(0)("table_name").ToString

        ' housekeeping
        mDataTable = Nothing

    End Function

    Function Get_Database_Name_From_SQL(
                       ByRef mYear As String,
                       ByRef mData_Type As String) As String

        'Purpose:       Gets the database name value from Table Locator table in URCS Controls database.
        'Assumptions:   Location of controls database from application settings
        'Affects:       Global variable for database name
        'Inputs:        Year value and data type value
        'Returns:       Database name
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Dim mDataTable As DataTable
        Dim mSQLstr As String

        ' Set the database name to the Controls database value
        Gbl_Controls_Database_Name = My.Settings.Controls_DB

        OpenSQLConnection(Gbl_Controls_Database_Name)

        ' Build the SQL statement
        mSQLstr = "SELECT Database_Name FROM U_TABLE_LOCATOR WHERE Year = " & mYear & " AND Data_Type = '" & UCase(mData_Type) & "'"

        Logger.Log("Get_Database_Name_From_SQL: " + mSQLstr)
        mDataTable = New DataTable

        ' Fill the datatable from SQL - this should be a unique record
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        ' Return the value
        Get_Database_Name_From_SQL = mDataTable.Rows(0)("database_name").ToString

        ' housekeeping
        mDataTable = Nothing

    End Function

    Function Get_URCS_Years_Table() As DataTable

        'Purpose:       Gets all URCS years values from table in URCS Controls database.
        'Assumptions:   
        'Affects:       
        'Inputs:        
        'Returns:       Datatable
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Dim mSQLstr As String

        Get_URCS_Years_Table = New DataTable

        ' Get the table name from the Table Locator table
        Gbl_URCS_Years_TableName = Trim(Get_Table_Name_From_SQL("1", "URCS_YEARS"))

        OpenSQLConnection(Gbl_Controls_Database_Name)

        ' Build the SQL statement
        mSQLstr = "select URCS_YEAR from " & Gbl_URCS_Years_TableName

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(Get_URCS_Years_Table)
        End Using

    End Function

    Function Get_CLASS1RAILLIST_Table() As DataTable

        'Purpose:       Gets all the records in the Class1RailList table
        'Assumptions:   
        'Affects:       
        'Inputs:        
        'Returns:       Datatable
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Dim mSQLstr As String

        Get_CLASS1RAILLIST_Table = New DataTable

        ' Get the table and database name from the Table Locator table
        Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "CLASS1RAILLIST")
        Gbl_Controls_Database_Name = Get_Database_Name_From_SQL("1", "CLASS1RAILLIST")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        ' Build the SQL statement
        mSQLstr = "select * from " & Gbl_Class1RailList_TableName

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(Get_CLASS1RAILLIST_Table)
        End Using

    End Function

    Function Get_WayRRR_Table() As DataTable

        'Purpose:       Gets all the records in the WAYRRR table in SQL, sorted ascending.
        'Assumptions:   
        'Affects:       
        'Inputs:        
        'Returns:       Datatable
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Dim mSQLstr As String

        Get_WayRRR_Table = New DataTable

        ' Get the table and database name from the Table Locator table
        Gbl_URCS_WAYRRR_TableName = Get_Table_Name_From_SQL("1", "WAYRRR")
        Gbl_Controls_Database_Name = Get_Database_Name_From_SQL("1", "WAYRRR")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        ' Build the SQL statement
        mSQLstr = "SELECT * FROM " & Gbl_URCS_WAYRRR_TableName & " WHERE RR_ID > 0 ORDER BY RR_ID ASC"

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(Get_WayRRR_Table)
        End Using

    End Function

    Function Get_SQL_DataTable(ByRef mConnection As SqlConnection, ByRef mSQLStr As String) As DataTable

        'Purpose:       Executes a SQL statement and returns the values in a datatable
        'Assumptions:   
        'Affects:       
        'Inputs:        Correctly formatted SQL statement
        'Returns:       Datatable
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Get_SQL_DataTable = New DataTable

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLStr, mConnection)
            daAdapter.Fill(Get_SQL_DataTable)
        End Using

    End Function

    Function Get_RRID_By_Short_Name(ByVal mShort_Name As String) As Integer

        'Purpose:       Gets the RRID for a railroad from the Class1RailList table
        'Assumptions:   
        'Affects:       
        'Inputs:        Valid Class 1 Short Name (BNSF, CSX, UP, etc.)
        'Returns:       Integer value of the RRID
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Dim mSQLstr As String
        Dim mDataTable As New DataTable

        Get_RRID_By_Short_Name = 0

        ' Get the table and database name from the Table Locator table
        Gbl_Railroads_TableName = Get_Table_Name_From_SQL("1", "CLASS1RAILLIST")
        Gbl_Controls_Database_Name = Get_Database_Name_From_SQL("1", "CLASS1RAILLIST")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        ' Build the SQL statement
        mSQLstr = "select RR_ID from " & Gbl_Railroads_TableName & " WHERE SHORT_NAME = '" & mShort_Name & "'"

        ' Fill the datatable from SQL - this will be a single record
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        ' Return the record's single value
        If mDataTable.Rows.Count > 0 Then
            Get_RRID_By_Short_Name = mDataTable.Rows(0).Item(0)
        End If

    End Function

    Function Get_States_Table() As DataTable

        'Purpose:       Gets all of the state codes (AL, CA, ND, etc.)
        'Assumptions:   
        'Affects:       
        'Inputs:        
        'Returns:       Datatable
        'Author:        Michael Sanders
        'Date:          10/26/2017
        '
        Dim mSQLstr As String

        Get_States_Table = New DataTable

        ' Get the table and database name from the Table Locator table
        Gbl_State_Codes_TableName = Get_Table_Name_From_SQL("1", "State_Codes")
        Gbl_Controls_Database_Name = Get_Database_Name_From_SQL("1", "State_Codes")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        ' Build the SQL statement
        mSQLstr = "select ch_state from " & Gbl_State_Codes_TableName

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(Get_States_Table)
        End Using

    End Function

    Function Get_AAR_Indexes_Table(ByRef mYear) As DataTable

        'Purpose:       Gets all of the AAR Indexes values for a particular year.
        'Assumptions:   
        'Affects:       
        'Inputs:        Which year to select all records.
        'Returns:       Datatable
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Dim mSQLstr As String

        Get_AAR_Indexes_Table = New DataTable

        ' Get the table and database name from the Table Locator table
        Gbl_URCS_AARIndex_TableName = Get_Table_Name_From_SQL("1", "URCS_AAR_INDEXES")
        Gbl_Controls_Database_Name = Get_Database_Name_From_SQL("1", "URCS_AAR_INDEXES")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        ' Build the SQL statement
        mSQLstr = "select * from " & Gbl_URCS_AARIndex_TableName & " WHERE year = " & mYear

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(Get_AAR_Indexes_Table)
        End Using

    End Function

    Function Get_Waybill_Years_Table() As DataTable

        'Purpose:       Gets all of the values from the WB Years SQL table
        'Assumptions:   
        'Affects:       
        'Inputs:        
        'Returns:       Datatable
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Dim mSQLstr As String

        Get_Waybill_Years_Table = New DataTable

        ' Get the table and database name from the Table Locator table
        Gbl_WB_Years_TableName = Get_Table_Name_From_SQL("1", "WB_YEARS")
        Gbl_Controls_Database_Name = Get_Database_Name_From_SQL("1", "WB_YEARS")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        ' Build the SQL statement
        mSQLstr = "select wb_year from " & Gbl_WB_Years_TableName & " ORDER BY wb_year DESC"

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(Get_Waybill_Years_Table)
        End Using

    End Function

    Function Build_Count_Trans_SQL_Statement(
                                  ByVal Year As Integer,
                                  ByVal RRICC As Decimal,
                                  ByVal Sch As Integer,
                                  ByVal StartLine As Integer,
                                  Optional ByVal EndLine As Object = Nothing) As String

        'Purpose:       This routine returns the SQL statement to count select records in the Trans table 
        'Assumptions:   
        'Affects:       
        'Inputs:        Which year, RRICC code, Schedule, starting line number and optional ending line.
        'Returns:       formatted SQL statement
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "Trans")

        If IsNothing(EndLine) Then
            Build_Count_Trans_SQL_Statement = "Select count(*) from " & Gbl_Trans_TableName & " WHERE " &
                "(YEAR = " & Year.ToString & " " &
                "AND RRICC = " & RRICC.ToString & " " &
                "AND SCH = " & Sch.ToString & " " &
                "AND LINE = " & StartLine.ToString & ")"
        Else
            Build_Count_Trans_SQL_Statement = "Select count(*) from " & Gbl_Trans_TableName & " WHERE " &
                "(YEAR = " & Year.ToString & " " &
                "AND RRICC = " & RRICC.ToString & " " &
                "AND SCH = " & Sch.ToString & " " &
                "AND LINE BETWEEN " & StartLine.ToString & " " &
                "AND " & EndLine.ToString & ")"
        End If

    End Function

    Function Build_Update_Trans_SQL_Record_Statement(
                                   ByVal mYear As Integer,
                                   ByVal mRRICC As Decimal,
                                   ByVal mSch As Integer,
                                   ByVal mLine As Integer,
                                   ByVal mC1 As Object,
                                   Optional ByVal mC2 As Object = 0,
                                   Optional ByVal mC3 As Object = 0,
                                   Optional ByVal mC4 As Object = 0,
                                   Optional ByVal mC5 As Object = 0,
                                   Optional ByVal mC6 As Object = 0,
                                   Optional ByVal mC7 As Object = 0,
                                   Optional ByVal mC8 As Object = 0,
                                   Optional ByVal mC9 As Object = 0,
                                   Optional ByVal mC10 As Object = 0,
                                   Optional ByVal mC11 As Object = 0,
                                   Optional ByVal mC12 As Object = 0,
                                   Optional ByVal mC13 As Object = 0,
                                   Optional ByVal mC14 As Object = 0,
                                   Optional ByVal mC15 As Object = 0) As String

        'Purpose:       This routine returns the SQL statement to update a record in the Trans table 
        'Assumptions:   
        'Affects:       
        'Inputs:        Year, RRICC Code, Schedule, Line and column value(s) to update.
        'Returns:       formatted SQL statement
        'Author:        Michael Sanders
        'Date:          10/25/2017

        Dim mComma As String
        Dim mStringb As New StringBuilder
        Dim mDate As Date

        mComma = ", "
        mDate = Now().ToString("MM/dd/yyyy")

        If IsNothing(mC2) Then mC2 = 0
        If IsNothing(mC3) Then mC3 = 0
        If IsNothing(mC4) Then mC4 = 0
        If IsNothing(mC5) Then mC5 = 0
        If IsNothing(mC6) Then mC6 = 0
        If IsNothing(mC6) Then mC6 = 0
        If IsNothing(mC7) Then mC7 = 0
        If IsNothing(mC8) Then mC8 = 0
        If IsNothing(mC9) Then mC9 = 0
        If IsNothing(mC10) Then mC10 = 0
        If IsNothing(mC11) Then mC11 = 0
        If IsNothing(mC12) Then mC12 = 0
        If IsNothing(mC13) Then mC13 = 0
        If IsNothing(mC14) Then mC14 = 0
        If IsNothing(mC15) Then mC15 = 0

        mStringb = New StringBuilder
        With mStringb
            .Append("UPDATE " & Trim(Gbl_Trans_TableName) & " Set " &
                    "C1 = " & CStr(mC1) & mComma)
            .Append("C2 = " & CStr(mC2) & mComma)
            .Append("C3 = " & CStr(mC3) & mComma)
            .Append("C4 = " & CStr(mC4) & mComma)
            .Append("C5 = " & CStr(mC5) & mComma)
            .Append("C6 = " & CStr(mC6) & mComma)
            .Append("C7 = " & CStr(mC7) & mComma)
            .Append("C8 = " & CStr(mC8) & mComma)
            .Append("C9 = " & CStr(mC9) & mComma)
            .Append("C10 = " & CStr(mC10) & mComma)
            .Append("C11 = " & CStr(mC11) & mComma)
            .Append("C12 = " & CStr(mC12) & mComma)
            .Append("C13 = " & CStr(mC13) & mComma)
            .Append("C14 = " & CStr(mC14) & mComma)
            .Append("C15 = " & CStr(mC15) & mComma)
            .Append("UpdatedDateTime = CAST('" & mDate.ToString & "' AS date)" & mComma)
            .Append("UpdatedBy = '" & CStr(My.User.Name) & "' ")
            .Append("WHERE YEAR = " & CStr(mYear) & " ")
            .Append("And RRICC = " & CStr(mRRICC) & " ")
            .Append("And SCH = " & CStr(mSch) & " ")
            .Append("And LINE = " & CStr(mLine))
        End With

        Build_Update_Trans_SQL_Record_Statement = mStringb.ToString

    End Function

    Function Build_Insert_Trans_SQL_Record_Statement(ByVal mYear As Integer,
                                    ByVal mRRICC As Decimal,
                                    ByVal mSch As Integer,
                                    ByVal mLine As Integer,
                                    ByVal mC1 As Decimal,
                                    Optional ByVal mC2 As Object = 0,
                                    Optional ByVal mC3 As Object = 0,
                                    Optional ByVal mC4 As Object = 0,
                                    Optional ByVal mC5 As Object = 0,
                                    Optional ByVal mC6 As Object = 0,
                                    Optional ByVal mC7 As Object = 0,
                                    Optional ByVal mC8 As Object = 0,
                                    Optional ByVal mC9 As Object = 0,
                                    Optional ByVal mC10 As Object = 0,
                                    Optional ByVal mC11 As Object = 0,
                                    Optional ByVal mC12 As Object = 0,
                                    Optional ByVal mC13 As Object = 0,
                                    Optional ByVal mC14 As Object = 0,
                                    Optional ByVal mC15 As Object = 0) As String

        'Purpose:       This routine returns the SQL statement to insert a record in the Trans table 
        'Assumptions:   
        'Affects:       
        'Inputs:        Year, RRICC Code, Schedule and Line
        'Returns:       formatted SQL statement
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Dim mComma As String
        Dim mStringB As StringBuilder
        Dim mDate As Date

        mComma = ", "
        mDate = Now().ToString("MM/dd/yyyy")

        mStringB = New StringBuilder

        If IsNothing(mC2) Then mC2 = 0
        If IsNothing(mC3) Then mC3 = 0
        If IsNothing(mC4) Then mC4 = 0
        If IsNothing(mC5) Then mC5 = 0
        If IsNothing(mC6) Then mC6 = 0
        If IsNothing(mC6) Then mC6 = 0
        If IsNothing(mC7) Then mC7 = 0
        If IsNothing(mC8) Then mC8 = 0
        If IsNothing(mC9) Then mC9 = 0
        If IsNothing(mC10) Then mC10 = 0
        If IsNothing(mC11) Then mC11 = 0
        If IsNothing(mC12) Then mC12 = 0
        If IsNothing(mC13) Then mC13 = 0
        If IsNothing(mC14) Then mC14 = 0
        If IsNothing(mC15) Then mC15 = 0

        With mStringB
            .Append("INSERT INTO " & Trim(Gbl_Trans_TableName) & " (" &
            "Year, RRICC, Sch, Line, C1, C2, C3, C4, C5, " &
            "C6, C7, C8, C9, C10, C11, C12, C13, C14, C15, UpdatedDateTime, UpdatedBy " &
            ") VALUES (")
            .Append(CStr(mYear) & mComma)
            .Append(CStr(mRRICC) & mComma)
            .Append(CStr(mSch) & mComma)
            .Append(CStr(mLine) & mComma)
            .Append(CStr(mC1) & mComma)
            .Append(CStr(mC2) & mComma)
            .Append(CStr(mC3) & mComma)
            .Append(CStr(mC4) & mComma)
            .Append(CStr(mC5) & mComma)
            .Append(CStr(mC6) & mComma)
            .Append(CStr(mC7) & mComma)
            .Append(CStr(mC8) & mComma)
            .Append(CStr(mC9) & mComma)
            .Append(CStr(mC10) & mComma)
            .Append(CStr(mC11) & mComma)
            .Append(CStr(mC12) & mComma)
            .Append(CStr(mC13) & mComma)
            .Append(CStr(mC14) & mComma)
            .Append(CStr(mC15) & mComma)
            .Append("CAST('" & mDate.ToString & "' AS date)" & mComma)
            .Append("'" & CStr(My.User.Name) & "')")
        End With

        Build_Insert_Trans_SQL_Record_Statement = mStringB.ToString

    End Function

    Function Build_Insert_Trans_SQL_Field_Statement(ByVal mYear As Integer,
                                    ByVal mRRICC As Decimal,
                                    ByVal mSch As Integer,
                                    ByVal mLine As Integer,
                                    ByVal mColumn As Integer,
                                    ByVal mValue As Decimal) As String

        'Purpose:       This routine returns the SQL statement to insert a record's field in the Trans table 
        'Assumptions:   
        'Affects:       
        'Inputs:        Year, RRICC Code, Schedule and Line
        'Returns:       formatted SQL statement
        'Author:        Michael Sanders
        'Date:          10/26/2017

        Dim mComma As String
        Dim mStringB As StringBuilder
        Dim mDate As Date

        mComma = ", "
        mDate = Now().ToString("MM/dd/yyyy")

        mStringB = New StringBuilder

        With mStringB
            .Append("INSERT INTO " & Trim(Gbl_Trans_TableName) & " (" &
            "Year, RRICC, Sch, Line, C" & mColumn & ", UpdatedDateTime, UpdatedBy " &
            ") VALUES (")
            .Append(CStr(mYear) & mComma)
            .Append(CStr(mRRICC) & mComma)
            .Append(CStr(mSch) & mComma)
            .Append(CStr(mLine) & mComma)
            .Append(CStr(mValue) & mComma)
            .Append("CAST('" & mDate.ToString & "' AS date)" & mComma)
            .Append("'" & CStr(My.User.Name) & "')")
        End With

        Build_Insert_Trans_SQL_Field_Statement = mStringB.ToString

    End Function

    Function Build_Update_Trans_SQL_Field_Statement(
                                   ByVal mYear As Integer,
                                   ByVal mRRICC As Decimal,
                                   ByVal mSch As Integer,
                                   ByVal mLine As Integer,
                                   ByVal mCol As Integer,
                                   ByVal mValue As Object) As String

        'Purpose:       This routine returns the SQL statement to update a record in the Trans table 
        'Assumptions:   
        'Affects:       
        'Inputs:        Year, RRICC Code, Schedule, Line and column value(s) to update.
        'Returns:       formatted SQL statement
        'Author:        Michael Sanders
        'Date:          10/25/2017

        Dim mComma As String
        Dim mStringb As New StringBuilder
        Dim mDate As Date

        mComma = ", "
        mDate = Now().ToString("MM/dd/yyyy")

        mStringb = New StringBuilder
        With mStringb
            .Append("UPDATE " & Trim(Gbl_Trans_TableName) & " Set " &
                    "C" & mCol.ToString & " = " & CStr(mValue) & mComma)
            .Append("UpdatedDateTime = CAST('" & mDate.ToString & "' AS date)" & mComma)
            .Append("UpdatedBy = '" & CStr(My.User.Name) & "' ")
            .Append("WHERE YEAR = " & CStr(mYear) & " ")
            .Append("And RRICC = " & CStr(mRRICC) & " ")
            .Append("And SCH = " & CStr(mSch) & " ")
            .Append("And LINE = " & CStr(mLine) & " ")
        End With

        Build_Update_Trans_SQL_Field_Statement = mStringb.ToString

    End Function

    Function Build_Delete_Trans_SQL_Statement(
                                   ByVal mYear As Integer,
                                   ByVal mSch As Integer,
                                   ByVal mLine As Integer) As String

        'Purpose:       This routine returns the SQL statement to delete a record in the Trans table
        'Assumptions:   
        'Affects:       
        'Inputs:        Year, Schedule, and Line Number to delete
        'Returns:       formatted SQL statement
        'Author:        Michael Sanders
        'Date:          10/25/2017

        Dim mComma As String

        mComma = ", "

        Build_Delete_Trans_SQL_Statement = "DELETE FROM " & Get_Table_Name_From_SQL("1", "TRANS") & " WHERE " &
            "SCH = " & CStr(mSch) & " And " &
            "LINE = " & CStr(mLine)

    End Function

    Function Build_Select_AARIndex_By_Region_SQL_Statement(ByVal mRegion As Integer, ByVal mYear As Integer) As String

        'Purpose:       This routine returns the SQL statement to select record(s) in the AAR Index table
        'Assumptions:   
        'Affects:       
        'Inputs:        Region and Year
        'Returns:       formatted SQL statement
        'Author:        Michael Sanders
        'Date:          10/24/2017

        Build_Select_AARIndex_By_Region_SQL_Statement = "Select * FROM " & Get_Table_Name_From_SQL("1", "URCS_AAR_Indexes") & " WHERE Region = " &
            CStr(mRegion) & " And Year = " & CStr(mYear)

    End Function

    Function Build_Insert_AARIndex_SQL_statement(
        ByVal mTableName As String,
        ByVal mRegion As String,
        ByVal mYear As String,
        ByVal mFuel As String,
        ByVal mMS As String,
        ByVal mPS As String,
        ByVal mWage As String,
        ByVal mMP As String) As StringBuilder

        'Purpose:       This routine returns the SQL statement to insert a specific record in the AAR Index table
        'Assumptions:   
        'Affects:       
        'Inputs:        Region and Year
        'Returns:       formatted SQL statement
        'Author:        Michael Sanders
        'Date:          10/24/2017

        Dim mComma As String

        ' Get the table name from the Table Locator table
        Gbl_WB_Years_TableName = Get_Table_Name_From_SQL("1", "WB_YEARS")

        mComma = ", "

        Build_Insert_AARIndex_SQL_statement = New StringBuilder
        Build_Insert_AARIndex_SQL_statement.Append("INSERT INTO " & mTableName & " (" &
            "Region, Year, Fuel, MS, PS, Wage, MP " &
            ") VALUES (")
        Build_Insert_AARIndex_SQL_statement.Append(mRegion & mComma)
        Build_Insert_AARIndex_SQL_statement.Append(mYear & mComma)
        Build_Insert_AARIndex_SQL_statement.Append(mFuel & mComma)
        Build_Insert_AARIndex_SQL_statement.Append(mMS & mComma)
        Build_Insert_AARIndex_SQL_statement.Append(mPS & mComma)
        Build_Insert_AARIndex_SQL_statement.Append(mWage & mComma)
        Build_Insert_AARIndex_SQL_statement.Append(mMP & ")")

    End Function

    Function Build_Update_AARIndex_SQL_Statement(
        ByVal mTableName As String,
        ByVal mRegion As String,
        ByVal mYear As String,
        ByVal mFuel As String,
        ByVal mMS As String,
        ByVal mPS As String,
        ByVal mWage As String,
        ByVal mMP As String) As String
        '
        ' This routine returns the SQL statement to update a specific record in the index table
        ' Created by: Michael Sanders
        ' Date: 10/24/2017
        '
        Dim mComma As String

        mComma = ", "

        Build_Update_AARIndex_SQL_Statement = "UPDATE " & mTableName & " Set Fuel = " & mFuel & mComma &
            "MS = " & mMS & mComma &
            "PS = " & mPS & mComma &
            "Wage = " & mWage & mComma &
            "MP = " & mMP & " " &
            "WHERE Region = " & mRegion & " And Year = " & mYear

    End Function

    Function Count_Price_Index_Records(
        ByVal mYear As String,
        ByVal mIndex As String,
        ByVal mRegion As String) As Integer
        '
        ' This routine returns the count of specific records in the index table
        ' Created by: Michael Sanders
        ' Date: 10/24/2017
        '
        Dim mComma As String
        Dim mSQLStr As String
        Dim mDataTable As New DataTable

        Gbl_Price_Index_DatabaseName = Get_Database_Name_From_SQL("1", "INDEX")
        Gbl_Price_Index_TableName = Get_Table_Name_From_SQL("1", "INDEX")

        mComma = ", "
        Count_Price_Index_Records = 0

        mSQLStr = "Select * FROM " & Gbl_Price_Index_TableName &
            " WHERE Year = " & mYear &
            " And [Index] = " & mIndex &
            " And Region = " & mRegion

        OpenSQLConnection(Gbl_Price_Index_DatabaseName)

        Using daAdapter As New SqlDataAdapter(mSQLStr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        If mDataTable.Rows.Count > 0 Then
            ' Return the number of records found
            Count_Price_Index_Records = mDataTable.Rows.Count
        End If

        mDataTable = Nothing

    End Function

    Public Sub Insert_Price_Index_Record(ByVal mYear As Integer,
        ByVal mIndex As Integer,
        ByVal mRegion As Integer,
        ByVal mCurrent_Year As Single,
        ByVal mCurrent_Year_Minus_1 As Single,
        ByVal mCurrent_Year_Minus_2 As Single,
        ByVal mCurrent_Year_Minus_3 As Single,
        ByVal mCurrent_Year_Minus_4 As Single)

        Dim mSQLCommand As SqlCommand
        Dim mStrSQL As StringBuilder
        Dim mComma As String = ", "

        mStrSQL = New StringBuilder
        mStrSQL.Append("INSERT INTO " & Get_Table_Name_From_SQL("1", "INDEX") & " (" &
            "Year, [Index], Region, Current_Year, Current_Year_Minus_1, Current_Year_Minus_2, Current_Year_Minus_3, Current_Year_Minus_4 " &
            ") VALUES (")
        mStrSQL.Append(mYear & mComma)
        mStrSQL.Append(mIndex & mComma)
        mStrSQL.Append(mRegion & mComma)
        mStrSQL.Append(mCurrent_Year & mComma)
        mStrSQL.Append(mCurrent_Year_Minus_1 & mComma)
        mStrSQL.Append(mCurrent_Year_Minus_2 & mComma)
        mStrSQL.Append(mCurrent_Year_Minus_3 & mComma)
        mStrSQL.Append(mCurrent_Year_Minus_4 & ")")

        Gbl_Price_Index_DatabaseName = Get_Database_Name_From_SQL("1", "INDEX")
        Gbl_Price_Index_TableName = Get_Table_Name_From_SQL("1", "INDEX")

        ' Open the SQL connection
        OpenSQLConnection(Gbl_Price_Index_DatabaseName)

        ' execute the command
        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandText = mStrSQL.ToString
        mSQLCommand.ExecuteNonQuery()

    End Sub

    Public Sub Update_Price_Index_Record(ByVal mYear As Integer,
        ByVal mIndex As Integer,
        ByVal mRegion As Integer,
        ByVal mField_Name As String,
        ByVal mCurrent_Year_Value As Single)

        Dim mSQLCommand As SqlCommand
        Dim mStrSQL As StringBuilder

        mStrSQL = New StringBuilder
        mStrSQL.Append("UPDATE " & Get_Table_Name_From_SQL("1", "INDEX") & " Set " & mField_Name & " = " & mCurrent_Year_Value.ToString &
            " WHERE Year = " & mYear.ToString & " And" &
            " [Index] = " & mIndex.ToString & " And" &
            " Region = " & mRegion.ToString)

        Gbl_Price_Index_TableName = Get_Table_Name_From_SQL("1", "INDEX")

        ' Open the SQL connection
        OpenSQLConnection(My.Settings.Controls_DB)

        ' execute the command
        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandText = mStrSQL.ToString
        mSQLCommand.ExecuteNonQuery()

    End Sub

    Function Build_Select_Trans_SQL_Statement(ByVal mYear As String,
    ByVal mRRICC As String,
    ByVal mSch As String,
    ByVal mStartLine As String,
    Optional ByVal mEndLine As String = "0") As String
        '
        ' This routine returns the SQL statement to select specific record(s) from the Trans table
        ' Created by: Michael Sanders
        ' Date: 10/23/2017
        '
        Dim mComma As String

        mComma = ", "

        If mEndLine = "0" Then
            Build_Select_Trans_SQL_Statement = "Select * from " & Get_Table_Name_From_SQL("1", "TRANS") & " WHERE " &
                "YEAR = " & mYear & " " &
                "And RRICC = " & mRRICC & " " &
                "And SCH = " & mSch & " " &
                "And LINE = " & mStartLine
        Else
            Build_Select_Trans_SQL_Statement = "Select * from " & Get_Table_Name_From_SQL("1", "TRANS") & " WHERE " &
                "YEAR = " & mYear & " " &
                "And RRICC = " & mRRICC & " " &
                "And SCH = " & mSch & " " &
                "And LINE BETWEEN " & mStartLine & " " &
                "And " & mEndLine
        End If

    End Function

    Function Build_Simple_Count_Records_SQL_Statement(ByVal mTableName As String) As String
        '
        ' This routine returns the SQL statement to count the number of records in any table
        ' Created by: Michael Sanders
        ' Date: 10/23/2017
        '

        Build_Simple_Count_Records_SQL_Statement = "Select Count(*) from " & mTableName

    End Function

    Function Count_Waybills(ByVal mYear As String) As Integer
        '
        ' This routine returns the record count of a year's Waybills table
        ' Created by: Michael Sanders
        ' Date: 10/23/2017
        '
        Dim mDatatable As DataTable
        Dim mSQLstr As String

        mDatatable = New DataTable

        ' Get the Table and database for the year's waybill sample
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(mYear.ToString, "MASKED")
        Gbl_Waybill_Database_Name = Get_Database_Name_From_SQL(mYear.ToString, "MASKED")

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        mSQLstr = "Select Count(*) As MyCount from " & Gbl_WB_Years_TableName

        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        ' Return the number of records found
        Count_Waybills = mDatatable.Rows(0)("MyCount")

        mDatatable = Nothing

    End Function

    Function Count_Waybills_By_Railroad(ByVal mYear As Integer, ByVal mRailroad As String) As String
        '
        ' This routine returns the count of Waybills for a year for one railroad.
        ' Created by: Michael Sanders
        ' Date: 10/23/2017
        '

        Dim mDatatable As DataTable
        Dim mSQLstr As String
        Dim mstrRR As Integer

        mDatatable = New DataTable

        ' Get the AARID from the database
        mstrRR = Get_AARID(mYear, mRailroad.ToString)

        ' Get the Table and database for the year's waybill sample
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(mYear.ToString, "MASKED")
        Gbl_Waybill_Database_Name = Get_Database_Name_From_SQL(mYear.ToString, "MASKED")

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        mSQLstr = "Select Count(*) As MyCount from " & Gbl_Masked_TableName & " WHERE (report_rr = " & mstrRR & ")"

        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        ' Return the number of records found
        Count_Waybills_By_Railroad = mDatatable.Rows(0)("MyCount")

        mDatatable = Nothing

    End Function

    Function Count_Waybills_By_State(ByVal mYear As Integer, ByVal mState As String) As String

        Dim mStrSQL As String
        Dim mSt As String
        Dim mDatatable As DataTable

        mDatatable = New DataTable

        ' Get the table and database information for the Waybill table
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(mYear.ToString, "MASKED")
        Gbl_Waybill_Database_Name = Get_Database_Name_From_SQL(mYear.ToString, "MASKED")

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        mSt = Trim(mState)

        mStrSQL = "Select count(*) FROM " & Gbl_Masked_TableName & " As MyCount WHERE " &
            mSt & "_Flg = 1 Or (O_ST = '" & mSt & "' Or JCT1_ST ='" & mSt &
            "' Or JCT2_ST = '" & mSt & "' Or JCT3_ST = '" & mSt & "' Or JCT4_ST = '" & mSt &
            "' Or JCT5_ST = '" & mSt & "' Or JCT6_ST = '" & mSt & "' Or JCT7_ST = '" & mSt &
            "' Or T_ST = '" & mSt & "')"

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        ' Return the number of records found
        Count_Waybills_By_State = mDatatable.Rows(0)("MyCount")

        mDatatable = Nothing

    End Function

    Function Count_States_Waybills(ByVal mYear As Integer, ByVal mSelectedStates As ListBox.SelectedObjectCollection) As String

        Dim mStrSQL As String
        Dim mLooper As Integer
        Dim mDatatable As DataTable

        mDatatable = New DataTable

        Count_States_Waybills = 0

        ' Exit (returns 0) if no values for mStates has been passed
        If mSelectedStates.Count = 0 Then Exit Function

        ' Get the table and database information for the Waybill table
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(mYear.ToString, "MASKED")
        Gbl_Waybill_Database_Name = Get_Database_Name_From_SQL(mYear.ToString, "MASKED")

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        'start bulding the SQL statement.  Load the state flg(s)
        mStrSQL = "Select count(*) FROM " & Gbl_Masked_TableName & " As MyCount WHERE "

        For mLooper = 1 To mSelectedStates.Count                                                    'Count is 1 based
            mStrSQL = mStrSQL & Trim(mSelectedStates.Item(mLooper - 1).ToString) & "_Flg = 1 OR "   'Position is zero based
        Next mLooper

        ' That takes care of the state flgs.  Now to address the origin, termination and junction states
        ' Origin states
        For mLooper = 1 To mSelectedStates.Count
            mStrSQL = mStrSQL & "O_ST = '" & Trim(mSelectedStates.Item(mLooper - 1).ToString) & "' OR "
        Next mLooper

        ' Termination States
        For mLooper = 1 To mSelectedStates.Count
            mStrSQL = mStrSQL & "T_ST = '" & Trim(mSelectedStates.Item(mLooper - 1).ToString) & "' OR "
        Next

        ' Junction states
        For mLooper = 1 To mSelectedStates.Count

            For mJunctionLooper = 1 To 7
                mStrSQL = mStrSQL & "JCT" & mJunctionLooper.ToString & "_ST = '" & Trim(mSelectedStates.Item(mLooper - 1).ToString) & "'"
                If mJunctionLooper <> 7 Then
                    mStrSQL = mStrSQL & " OR "
                End If
            Next mJunctionLooper

            If mLooper <> mSelectedStates.Count Then
                mStrSQL = mStrSQL & " OR "
            End If

        Next mLooper

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        ' Return the number of records found
        Count_States_Waybills = mDatatable.Rows(0)(0)

        mDatatable = Nothing

    End Function

    Function Count_Marks_Records(ByVal mRoadMark As String, ByVal mLastMaintStamp As String) As String
        '
        ' This routine returns the count of Marks records for a Road_Mark and a Last_Maint_Stamp.
        ' Created by: Michael Sanders
        ' Date: 12/6/2017
        '
        Dim mDatatable As DataTable
        Dim mSQLstr As String

        mDatatable = New DataTable

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        mSQLstr = "Select Count(*) as MyCount from " & Gbl_Marks_Tablename &
            " WHERE (Road_Mark = '" & mRoadMark & "' AND Last_Maint_Stamp = '" & mLastMaintStamp & "')"

        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        ' Return the number of records found
        Count_Marks_Records = mDatatable.Rows(0)("MyCount")

        mDatatable = Nothing

    End Function

    Function Count_CSM_Records(ByVal mRoadMark As String, ByVal mFSAC As String, ByVal mLastMaintStamp As String) As String
        '
        ' This routine returns the count of Marks records for a Road_Mark and a Last_Maint_Stamp.
        ' Created by: Michael Sanders
        ' Date: 12/6/2017
        '
        Dim mDatatable As DataTable
        Dim mSQLstr As String

        mDatatable = New DataTable

        OpenSQLConnection(Gbl_Controls_Database_Name)

        mSQLstr = "Select Count(*) as MyCount from " & Gbl_CSM_TableName &
            " WHERE (Road_Mark = '" & mRoadMark & "' AND FSAC = '" & mFSAC & "' AND Last_Maint_Stamp = '" & mLastMaintStamp & "')"

        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        ' Return the number of records found
        Count_CSM_Records = mDatatable.Rows(0)("MyCount")

        mDatatable = Nothing

    End Function

    Function Count_Productivity_Records(ByVal mYear As Integer,
                                        ByVal mLength_Of_Haul_Stratum As Integer,
                                        ByVal mCar_Type_Stratum As Integer,
                                        ByVal mLading_Weight_Stratum As Integer,
                                        ByVal mCars_Stratum As Integer) As Integer
        Dim mDatatable As DataTable
        Dim mSQLstr As String

        mDatatable = New DataTable

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        mSQLstr = "Select Count(*) as MyCount from " & Gbl_Productivity_TableName & " " &
            "WHERE (Year = " & mYear.ToString & " AND " &
            "Length_Of_Haul_Stratum = " & mLength_Of_Haul_Stratum.ToString & " AND " &
            "Car_Type_Stratum = " & mCar_Type_Stratum.ToString & " AND " &
            "Lading_Weight_Stratum = " & mLading_Weight_Stratum.ToString & " AND " &
            "Cars_Stratum = " & mCars_Stratum.ToString & ")"

        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        ' Return the number of records found
        Count_Productivity_Records = mDatatable.Rows(0)("MyCount")

        mDatatable = Nothing

    End Function

    Function Build_Update_FCS_SQL(ByVal mYear As Decimal,
                                    ByVal mAARID As Decimal,
                                    ByVal mLocal_Traffic_Cars As Decimal,
                                    ByVal mForwarded_Traffic_Cars As Decimal,
                                    ByVal mReceived_Traffic_Cars As Decimal,
                                    ByVal mBridged_Traffic_Cars As Decimal) As String

        Dim mComma As String

        mComma = ", "

        Build_Update_FCS_SQL = "UPDATE " & Gbl_URCS_FCS_TableName & " Set " &
            "Local_Traffic_Cars = " & CStr(mLocal_Traffic_Cars) & mComma &
            "Forwarded_Traffic_Cars = " & CStr(mForwarded_Traffic_Cars) & mComma &
            "Received_Traffic_Cars = " & CStr(mReceived_Traffic_Cars) & mComma &
            "Bridged_Traffic_Cars = " & CStr(mBridged_Traffic_Cars) & " " &
            "WHERE Year = " & CStr(mYear) & " " &
            "And AARID = " & CStr(mAARID)

    End Function

    Function Build_Insert_FCS_SQL(ByVal mYear As Decimal,
        ByVal mAARID As Decimal,
        ByVal mLocal_Traffic_Cars As Decimal,
        ByVal mForwarded_Traffic_Cars As Decimal,
        ByVal mReceived_Traffic_Cars As Decimal,
        ByVal mBridged_Traffic_Cars As Decimal) As String

        Dim mComma As String

        mComma = ", "

        Build_Insert_FCS_SQL = "INSERT INTO " & Gbl_URCS_FCS_TableName & " (" &
            "Year, " &
            "AARID, " &
            "Local_Traffic_Cars" & mComma &
            "Forwarded_Traffic_Cars" & mComma &
            "Received_Traffic_Cars" & mComma &
            "Bridged_Traffic_Cars" &
            ") VALUES (" &
            CStr(mYear) & mComma &
            CStr(mAARID) & mComma &
            CStr(mLocal_Traffic_Cars) & mComma &
            CStr(mForwarded_Traffic_Cars) & mComma &
            CStr(mReceived_Traffic_Cars) & mComma &
            CStr(mBridged_Traffic_Cars) & ")"

    End Function

    Function Build_Select_RRInfo_SQL(ByVal mAARID As Integer) As String

        Gbl_URCS_WAYRRR_TableName = Get_Table_Name_From_SQL("1", "WAYRRR")

        Build_Select_RRInfo_SQL = "Select * from " & Gbl_URCS_WAYRRR_TableName & " WHERE " &
            "AARID = " & CStr(mAARID)

    End Function

    Function Build_Count_Railroads_SQL() As String

        Build_Count_Railroads_SQL = "Select Count(*) As MyCount from " & Get_Table_Name_From_SQL("1", "R_URCS_Railroads")

    End Function

    Function Build_Select_Railroads_SQL() As String

        Build_Select_Railroads_SQL = "Select * from " & Get_Table_Name_From_SQL("1", "R_URCS_Railroads")

    End Function

    Function Build_Select_Railroads_By_URCSCode_SQL() As String

        Build_Select_Railroads_By_URCSCode_SQL = "Select * from " & Get_Table_Name_From_SQL("1", "R_URCS_Railroads") & " ORDER BY URCSCODE"

    End Function

    Function Build_Select_Dictionary_SQL(Optional ByVal mArguement As String = "") As String

        Build_Select_Dictionary_SQL = ""

        If mArguement = "" Then
            Build_Select_Dictionary_SQL = "Select * FROM " & Get_Table_Name_From_SQL("1", "data_dictionary") & " WHERE Expiration_Dt = 9999 ORDER BY sortcode"
        Else
            If IsNumeric(mArguement) Then
                Select Case Val(mArguement)
                    Case Is < 900000
                        'Retrieve a single railroad set
                        Build_Select_Dictionary_SQL = "Select * FROM " & Get_Table_Name_From_SQL("1", "data_dictionary") & " " &
                            "WHERE (loadcode = 0 Or loadcode = 3) And Expiration_Dt = 9999 ORDER BY urcsid, sortcode"
                    Case 900004, 900007
                        'Retrieve a region set
                        Build_Select_Dictionary_SQL = "Select * FROM " & Get_Table_Name_From_SQL("1", "data_dictionary") & " " &
                            "WHERE (loadcode = 1 Or loadcode = 3 Or loadcode = 4 Or loadcode = 5) And Expiration_Dt = 9999 ORDER BY urcsid, sortcode"
                    Case 900099
                        'Retrieve the National set
                        Build_Select_Dictionary_SQL = "Select * FROM " & Get_Table_Name_From_SQL("1", "data_dictionary") & " " &
                            "WHERE (loadcode = 2 Or loadcode = 3) And Expiration_Dt = 9999 ORDER BY urcsid, sortcode"
                End Select
            Else
                Build_Select_Dictionary_SQL = "Select * FROM " & Get_Table_Name_From_SQL("1", "data_dictionary") & " "
                Build_Select_Dictionary_SQL = Build_Select_Dictionary_SQL & "WHERE (" & mArguement & ") And Expiration_Dt = 9999 ORDER BY sortcode"
            End If
        End If
EndIt:

    End Function

    Function Select_Scaled_Trans_Value(ByVal Year As Integer,
    ByVal RRICC As Decimal,
    ByVal Sch As Integer,
    ByVal Line As Integer,
    ByVal Col As Integer,
    ByVal Scaler As Integer) As Double

        Dim mDatatable As DataTable
        Dim mStrSQL As String

        mDatatable = New DataTable

        ' Get the table and database information for the Waybill table
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(Year.ToString, "MASKED")
        Gbl_Controls_Database_Name = Get_Database_Name_From_SQL(Year.ToString, "MASKED")

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        Select_Scaled_Trans_Value = 0

        mStrSQL = Build_Count_Trans_SQL_Statement(Year, RRICC, Sch, Line)

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        If mDatatable.Rows(0)("MyCount") > 0 Then

            mStrSQL = "Select * FROM " & Gbl_Trans_TableName & " " &
            "WHERE year = " & Str(Year) & " And " &
            "RRICC = " & Str(RRICC) & " And " &
            "SCH = " & Str(Sch) & " And " &
            "LINE = " & Str(Line)

            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDatatable)
            End Using

            If Not DBNull.Value.Equals(mDatatable.Rows(0)("c" & CStr(Col)).Value) Then
                Select_Scaled_Trans_Value = mDatatable.Rows(0)("c" & CStr(Col)).Value
            End If

            If Select_Scaled_Trans_Value > 0 Then
                Select Case Scaler
                    Case 1
                        Select_Scaled_Trans_Value = Select_Scaled_Trans_Value / 10
                    Case 2
                        Select_Scaled_Trans_Value = Select_Scaled_Trans_Value / 100
                    Case 3
                        Select_Scaled_Trans_Value = Select_Scaled_Trans_Value / 1000
                    Case 4
                        Select_Scaled_Trans_Value = Select_Scaled_Trans_Value / 10000
                    Case 5
                        Select_Scaled_Trans_Value = Select_Scaled_Trans_Value / 100000
                    Case 6
                        Select_Scaled_Trans_Value = Select_Scaled_Trans_Value / 1000000
                End Select
            End If
        End If

        mDatatable = Nothing

    End Function

    Public Function Build445SQL(
                               ByVal mYearStr As String,
                               ByVal mWBTableName As String,
                               ByVal strInline As String) As String

        Dim mSQLStr As New StringBuilder
        Dim workstr As String
        Dim mSTCC_W49 As String
        Dim mDate As Date
        Dim mDate_Str As String
        Dim mYear As Integer
        Dim mWaybill As Class_Waybill
        Dim rand As New Random()

        mWaybill = New Class_Waybill

        'Get the values needed to do the masking routine

        ' STCC_w49
        workstr = Trim(Mid(strInline, 54, 7))
        Do While Len(workstr) < 7
            workstr = "0" & workstr
        Loop
        mSTCC_W49 = workstr

        ' Date
        workstr = Mid(strInline, 13, 2) & "/" & Mid(strInline, 15, 2) & "/" &
                Mid(strInline, 17, 2)
        workstr = Replace(workstr, " ", "0")
        mDate_Str = workstr
        mDate = CDate(workstr)

        ' Year
        mYear = Val(mYearStr)

        mSQLStr = New StringBuilder

        mSQLStr.Append("INSERT INTO [dbo].[" & mWBTableName & "] (")
        mSQLStr.Append("[serial_no],")
        mSQLStr.Append("[wb_num],")
        mSQLStr.Append("[wb_date],")
        mSQLStr.Append("[acct_period],")
        mSQLStr.Append("[u_cars],")
        mSQLStr.Append("[u_car_init],")
        mSQLStr.Append("[u_car_num],")
        mSQLStr.Append("[tofc_serv_code],")
        mSQLStr.Append("[u_tc_units],")
        mSQLStr.Append("[u_tc_init],")
        mSQLStr.Append("[u_tc_num],")
        mSQLStr.Append("[stcc_w49],")
        mSQLStr.Append("[bill_wght],")
        mSQLStr.Append("[act_wght],")
        mSQLStr.Append("[u_rev],")
        mSQLStr.Append("[tran_chrg],")
        mSQLStr.Append("[misc_chrg],")
        mSQLStr.Append("[intra_state_code],")
        mSQLStr.Append("[transit_code],")
        mSQLStr.Append("[all_rail_code],")
        mSQLStr.Append("[type_move],")
        mSQLStr.Append("[move_via_water],")
        mSQLStr.Append("[truck_for_rail],")
        mSQLStr.Append("[Shortline_Miles],")
        mSQLStr.Append("[rebill],")
        mSQLStr.Append("[stratum],")
        mSQLStr.Append("[subsample],")
        mSQLStr.Append("[transborder_flg],")
        mSQLStr.Append("[rate_flg],")
        mSQLStr.Append("[wb_id],")
        mSQLStr.Append("[report_rr],")
        mSQLStr.Append("[o_fsac],")
        mSQLStr.Append("[orr],")
        mSQLStr.Append("[jct1],")
        mSQLStr.Append("[jrr1],")
        mSQLStr.Append("[jct2],")
        mSQLStr.Append("[jrr2],")
        mSQLStr.Append("[jct3],")
        mSQLStr.Append("[jrr3],")
        mSQLStr.Append("[jct4],")
        mSQLStr.Append("[jrr4],")
        mSQLStr.Append("[jct5],")
        mSQLStr.Append("[jrr5],")
        mSQLStr.Append("[jct6],")
        mSQLStr.Append("[jrr6],")
        mSQLStr.Append("[jct7],")
        mSQLStr.Append("[jrr7],")
        mSQLStr.Append("[jct8],")
        mSQLStr.Append("[jrr8],")
        mSQLStr.Append("[jct9],")
        mSQLStr.Append("[trr],")
        mSQLStr.Append("[t_fsac],")
        mSQLStr.Append("[pop_cnt],")
        mSQLStr.Append("[stratum_cnt],")
        mSQLStr.Append("[report_period],")
        mSQLStr.Append("[car_own_mark],")
        mSQLStr.Append("[car_lessee_mark],")
        mSQLStr.Append("[car_cap],")
        mSQLStr.Append("[nom_car_cap],")
        mSQLStr.Append("[tare],")
        mSQLStr.Append("[outside_l],")
        mSQLStr.Append("[outside_w],")
        mSQLStr.Append("[outside_h],")
        mSQLStr.Append("[ex_outside_h],")
        mSQLStr.Append("[type_wheel],")
        mSQLStr.Append("[no_axles],")
        mSQLStr.Append("[draft_gear],")
        mSQLStr.Append("[art_units],")
        mSQLStr.Append("[pool_code],")
        mSQLStr.Append("[car_typ],")
        mSQLStr.Append("[mech],")
        mSQLStr.Append("[lic_st],")
        mSQLStr.Append("[mx_wght_rail],")
        mSQLStr.Append("[o_splc],")
        mSQLStr.Append("[t_splc],")
        mSQLStr.Append("[u_fuel_surchg],")
        mSQLStr.Append("[err_code1],")
        mSQLStr.Append("[err_code2],")
        mSQLStr.Append("[err_code3],")
        mSQLStr.Append("[err_code4],")
        mSQLStr.Append("[err_code5],")
        mSQLStr.Append("[err_code6],")
        mSQLStr.Append("[err_code7],")
        mSQLStr.Append("[err_code8],")
        mSQLStr.Append("[err_code9],")
        mSQLStr.Append("[err_code10],")
        mSQLStr.Append("[err_code11],")
        mSQLStr.Append("[err_code12],")
        mSQLStr.Append("[err_code13],")
        mSQLStr.Append("[err_code14],")
        mSQLStr.Append("[err_code15],")
        mSQLStr.Append("[err_code16],")
        mSQLStr.Append("[err_code17],")
        mSQLStr.Append("[car_own],")
        mSQLStr.Append("[tofc_unit_type],")
        mSQLStr.Append("[alk_flg],")
        mSQLStr.Append("[Tracking_No]")

        mSQLStr.Append(") VALUES (")            'do not delete

        ' serial_no
        mSQLStr.Append(Mid(strInline, 1, 6) & ",")
        ' wb_no
        mSQLStr.Append(Mid(strInline, 7, 6) & ",")
        ' wb_date
        mSQLStr.Append("'" & mDate_Str & "', ")
        ' acct_period
        ' This gets the month and year value
        workstr = Trim(Mid(strInline, 19, 4))
        workstr = Replace(workstr, " ", "0")
        Do While Len(workstr) < 4
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' u_cars
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 23, 4)) & ",")
        ' u_car_init
        mSQLStr.Append("'" & Trim(Mid(strInline, 27, 4)) & "', ")
        ' u_car_num
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 31, 6)) & ",")
        ' tofc_serv_code
        mSQLStr.Append("'" & Trim(Mid(strInline, 37, 3)) & "', ")
        ' u_tc_units
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 40, 4)) & ",")
        ' u_tc_init
        mSQLStr.Append("'" & Trim(Mid(strInline, 44, 4)) & "', ")
        ' u_tc_num
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 48, 6)) & ",")
        ' stcc_w49
        mSQLStr.Append("'" & mSTCC_W49 & "', ")
        ' bill_wght
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 61, 9)) & ",")
        ' act_wght
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 70, 9)) & ",")
        ' u_rev
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 79, 9)) & ",")
        ' tran_chrg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 88, 9)) & ",")
        ' misc_chrg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 97, 9)) & ",")
        ' intra_state_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 106, 1)) & ",")
        ' transit_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 107, 1)) & ",")
        ' all_rail_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 108, 1)) & ",")
        ' type_move
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 109, 1)) & ",")
        ' move_via_water
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 110, 1)) & ",")
        ' truck_for_rail
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 111, 1)) & ",")
        ' Shortline_Miles
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 112, 4)) & ",")
        ' rebill
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 116, 1)) & ",")
        ' stratum
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 117, 1)) & ",")
        ' subsample
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 118, 1)) & ",")
        ' tramsborder_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 119, 1)) & ",")
        ' rate_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 120, 1)) & ",")
        ' wb_id
        mSQLStr.Append("'" & Trim(Mid(strInline, 121, 25)) & "', ")
        ' report_rr
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 146, 3)) & ",")
        ' o_fsac
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 149, 5)) & ",")
        ' orr
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 154, 3)) & ",")
        ' jct1
        mSQLStr.Append("'" & Trim(Mid(strInline, 157, 5)) & "', ")
        ' jrr1
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 162, 3)) & ",")
        ' jct2
        mSQLStr.Append("'" & Trim(Mid(strInline, 165, 5)) & "', ")
        ' jrr2
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 170, 3)) & ",")
        ' jct3
        mSQLStr.Append("'" & Trim(Mid(strInline, 173, 5)) & "', ")
        ' jrr3
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 178, 3)) & ",")
        ' jct4
        mSQLStr.Append("'" & Trim(Mid(strInline, 181, 5)) & "', ")
        ' jrr4
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 186, 3)) & ",")
        ' jct5
        mSQLStr.Append("'" & Trim(Mid(strInline, 189, 5)) & "', ")
        ' jrr5
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 194, 3)) & ",")
        ' jct6
        mSQLStr.Append("'" & Trim(Mid(strInline, 197, 5)) & "', ")
        ' jrr6
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 202, 3)) & ",")
        ' jct7
        mSQLStr.Append("'" & Trim(Mid(strInline, 205, 5)) & "', ")
        ' jrr7
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 210, 3)) & ",")
        ' jct8
        mSQLStr.Append("'" & Trim(Mid(strInline, 213, 5)) & "', ")
        ' jrr8
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 218, 3)) & ",")
        ' jct9
        mSQLStr.Append("'" & Trim(Mid(strInline, 221, 5)) & "', ")
        ' trr
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 226, 3)) & ",")
        ' t_fsac
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 229, 5)) & ",")
        ' pop_cnt
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 234, 8)) & ",")
        ' stratum_cnt
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 242, 6)) & ",")
        ' report_period
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 248, 1)) & ",")
        ' car_own_mark
        mSQLStr.Append("'" & Trim(Mid(strInline, 249, 4)) & "', ")
        ' car_lessee_mark
        mSQLStr.Append("'" & Trim(Mid(strInline, 253, 4)) & "', ")
        ' car_cap
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 257, 5)) & ",")
        ' nom_car_cap
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 262, 3)) & ",")
        ' tare
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 265, 4)) & ",")
        ' outside_l
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 269, 5)) & ",")
        ' outside_w
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 274, 4)) & ",")
        ' outside_h
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 278, 4)) & ",")
        ' ex_outside_h
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 282, 4)) & ",")
        ' type_wheel
        mSQLStr.Append("'" & Trim(Mid(strInline, 286, 1)) & "', ")
        ' no_axles
        mSQLStr.Append("'" & Trim(Mid(strInline, 287, 1)) & "', ")
        ' draft_gear
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 288, 2)) & ",")
        ' art_units
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 290, 1)) & ",")
        ' pool_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 291, 7)) & ",")
        ' car_typ
        mSQLStr.Append("'" & Trim(Mid(strInline, 298, 4)) & "', ")
        ' mech
        mSQLStr.Append("'" & Trim(Mid(strInline, 302, 4)) & "', ")
        ' lic_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 306, 2)) & "', ")
        ' mx_wght_rail
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 308, 3)) & ",")
        ' o_splc
        workstr = Trim(Mid(strInline, 311, 6))
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' t_splc
        workstr = Trim(Mid(strInline, 317, 6))
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' u_fuel_surchg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 323, 9)) & ",")
        ' err_code1
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 383, 2)) & ",")
        ' err_code2
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 385, 2)) & ",")
        ' err_code3
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 387, 2)) & ",")
        ' err_code4
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 389, 2)) & ",")
        ' err_code5
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 391, 2)) & ",")
        ' err_code6
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 393, 2)) & ",")
        ' err_code7
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 395, 2)) & ",")
        ' err_code8
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 397, 2)) & ",")
        ' err_code9
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 399, 2)) & ",")
        ' err_code10
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 401, 2)) & ",")
        ' err_code11
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 403, 2)) & ",")
        ' err_code12
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 405, 2)) & ",")
        ' err_code13
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 407, 2)) & ",")
        ' err_code14
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 409, 2)) & ",")
        ' err_code15
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 411, 2)) & ",")
        ' err_code16
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 413, 2)) & ",")
        ' err_code17
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 415, 2)) & ",")
        'car_own
        mSQLStr.Append("'" & Trim(Mid(strInline, 417, 1)) & "', ")
        'tofc_unit_type
        mSQLStr.Append("'" & Mid(strInline, 419, 4) & "', ")
        'alk_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 432, 1)) & ",")
        'Tracking_No
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 433, 13)))

        mSQLStr.Append(")")
        Build445SQL = mSQLStr.ToString

    End Function

    'Updated to address 913 waybilll record.  11/16/2020
    Public Function Build913SQL(
                               ByVal mYearStr As String,
                               ByVal mWBTableName As String,
                               ByVal strInline As String) As String

        Dim mSQLStr As New StringBuilder
        Dim workstr As String
        Dim mRate_Flg As Integer
        Dim mSTCC_W49 As String
        Dim mDate As Date
        Dim mDate_Str As String
        Dim mYear As Integer
        Dim mReport_RR As Integer
        Dim mU_Rev As Single, mNew_U_Rev As Single
        Dim mTotal_Rev As Single, mNew_Total_Rev As Single
        Dim mORR_Rev As Single, mNew_ORR_Rev As Single
        Dim mJRR1_Rev As Single, mNew_JRR1_Rev As Single
        Dim mJRR2_Rev As Single, mNew_JRR2_Rev As Single
        Dim mJRR3_Rev As Single, mNew_JRR3_Rev As Single
        Dim mJRR4_Rev As Single, mNew_JRR4_Rev As Single
        Dim mJRR5_Rev As Single, mNew_JRR5_Rev As Single
        Dim mJRR6_Rev As Single, mNew_JRR6_Rev As Single
        Dim mTRR_Rev As Single, mNew_TRR_Rev As Single
        Dim mJF As Integer, mExp_Factor_Th As Integer
        Dim mWaybill As Class_Waybill
        Dim rand As New Random()

        mWaybill = New Class_Waybill

        'Get the values needed to do the masking routine

        ' STCC_w49
        workstr = Trim(Mid(strInline, 58, 7))
        Do While Len(workstr) < 7
            workstr = "0" & workstr
        Loop
        mSTCC_W49 = workstr

        ' Date
        workstr = Mid(strInline, 13, 2) & "/" & Mid(strInline, 15, 2) & "/" &
                Mid(strInline, 17, 4)
        workstr = Replace(workstr, " ", "0")
        mDate_Str = workstr
        mDate = CDate(workstr)

        ' Year
        mYear = Val(mYearStr)

        ' JF
        mJF = CInt(Mid(strInline, 350, 1))

        ' Rate_Flg
        mRate_Flg = CInt(Mid(strInline, 124, 1))

        ' Exp_Factor_Th
        mExp_Factor_Th = CInt(Mid(strInline, 351, 3))

        ' Report_rr
        mReport_RR = Val(Mid(strInline, 150, 3))

        'Load original data for revenue values to memvars
        ' u_rev
        mU_Rev = ReturnValidNumber(Mid(strInline, 83, 9))
        ' total_rev
        mTotal_Rev = ReturnValidNumber(Mid(strInline, 405, 10))
        ' orr_rev
        mORR_Rev = ReturnValidNumber(Mid(strInline, 415, 10))
        ' jrr1_rev
        mJRR1_Rev = ReturnValidNumber(Mid(strInline, 425, 10))
        ' jrr2_rev
        mJRR2_Rev = ReturnValidNumber(Mid(strInline, 435, 10))
        ' jrr3_rev
        mJRR3_Rev = ReturnValidNumber(Mid(strInline, 445, 10))
        ' jrr4_rev
        mJRR4_Rev = ReturnValidNumber(Mid(strInline, 455, 10))
        ' jrr5_rev
        mJRR5_Rev = ReturnValidNumber(Mid(strInline, 465, 10))
        ' jrr6_rev
        mJRR6_Rev = ReturnValidNumber(Mid(strInline, 475, 10))
        ' trr_rev
        mTRR_Rev = ReturnValidNumber(Mid(strInline, 485, 10))

        'now, load the original values to the new memvars
        mNew_U_Rev = mU_Rev
        mNew_Total_Rev = mTotal_Rev
        mNew_ORR_Rev = mORR_Rev
        mNew_JRR1_Rev = mJRR1_Rev
        mNew_JRR2_Rev = mJRR2_Rev
        mNew_JRR3_Rev = mJRR3_Rev
        mNew_JRR4_Rev = mJRR4_Rev
        mNew_JRR5_Rev = mJRR5_Rev
        mNew_JRR6_Rev = mJRR6_Rev
        mNew_TRR_Rev = mTRR_Rev

        'Check for need to mask. - Do not mask class 1 railroads
        Select Case mReport_RR
            Case 190, 131, 712, 555, 105, 482, 777, 802
                ' Do nothing for records by:
                ' 190 - Unknown
                ' 131 - Unknown
                ' 712 - CSXT
                ' 555 - NS
                ' 105 - CP
                ' 482 - Unknown
                ' 777 - BNSF
                ' 802 - UP
            Case Else
                If (mRate_Flg = 1) And (mU_Rev > 0) Then
                    'mask the u_rev field
                    mNew_U_Rev = CStr(MaskGenericValue(mSTCC_W49, mDate, mYear, Val(mU_Rev)))
                    'calculate the new total rev
                    mNew_Total_Rev = mNew_U_Rev * mExp_Factor_Th
                    'calculate and round the new orr_rev
                    mNew_ORR_Rev = Math.Round(mNew_Total_Rev * (mORR_Rev / mTotal_Rev))
                    'calculate and round the new jrr1_rev
                    mNew_JRR1_Rev = Math.Round(mNew_Total_Rev * (mJRR1_Rev / mTotal_Rev))
                    'calculate and round the new jrr2_rev
                    mNew_JRR2_Rev = Math.Round(mNew_Total_Rev * (mJRR2_Rev / mTotal_Rev))
                    'calculate and round the new jrr3_rev
                    mNew_JRR3_Rev = Math.Round(mNew_Total_Rev * (mJRR3_Rev / mTotal_Rev))
                    'calculate and math.round the new jrr4_rev
                    mNew_JRR4_Rev = Math.Round(mNew_Total_Rev * (mJRR4_Rev / mTotal_Rev))
                    'calculate and math.round the new jrr5_rev
                    mNew_JRR5_Rev = Math.Round(mNew_Total_Rev * (mJRR5_Rev / mTotal_Rev))
                    'calculate and math.round the new jrr6_rev
                    mNew_JRR6_Rev = Math.Round(mNew_Total_Rev * (mJRR6_Rev / mTotal_Rev))
                    'calculate and math.round the new trr_rev
                    mNew_TRR_Rev = Math.Round(mNew_Total_Rev * (mTRR_Rev / mTotal_Rev))
                    If mJF > 0 Then
                        mNew_TRR_Rev = mNew_TRR_Rev + (mNew_Total_Rev -
                            (mNew_ORR_Rev + mNew_JRR1_Rev + mNew_JRR2_Rev +
                            mNew_JRR3_Rev + mNew_JRR4_Rev + mNew_JRR5_Rev +
                            mNew_JRR6_Rev + mNew_TRR_Rev))
                    End If
                End If
        End Select

        mSQLStr = New StringBuilder

        ' If this table does not contain a Tracking_No field, add it to the record
        If Column_Exist(mWBTableName, "Tracking_No") = False Then
            Column_Add(mWBTableName, "Tracking_No", "BigInt")
        End If

        mSQLStr.Append("INSERT INTO [dbo].[" & mWBTableName & "] (")
        mSQLStr.Append("[serial_no]" &
            ", [wb_num]" &
            ", [wb_date]" &
            ", [acct_period]" &
            ", [u_cars]" &
            ", [u_car_init]" &
            ", [u_car_num]" &
            ", [tofc_serv_code]" &
            ", [u_tc_units]" &
            ", [u_tc_init]" &
            ", [u_tc_num]" &
            ", [stcc_w49]" &
            ", [bill_wght]" &
            ", [act_wght]" &
            ", [u_rev]" &
            ", [tran_chrg]")
        mSQLStr.Append(", [misc_chrg]" &
            ", [intra_state_code]" &
            ", [transit_code]" &
            ", [all_rail_code]" &
            ", [type_move]" &
            ", [move_via_water]" &
            ", [truck_for_rail]" &
            ", [Shortline_Miles]" &
            ", [rebill]" &
            ", [stratum]" &
            ", [subsample]" &
            ", [int_eq_flg]")
        mSQLStr.Append(", [rate_flg]" &
            ", [wb_id]" &
            ", [report_rr]" &
            ", [o_fsac]" &
            ", [orr]" &
            ", [jct1]" &
            ", [jrr1]" &
            ", [jct2]" &
            ", [jrr2]" &
            ", [jct3]" &
            ", [jrr3]" &
            ", [jct4]" &
            ", [jrr4]" &
            ", [jct5]" &
            ", [jrr5]" &
            ", [jct6]" &
            ", [jrr6]" &
            ", [jct7]" &
            ", [trr]" &
            ", [t_fsac]")
        mSQLStr.Append(", [pop_cnt]" &
            ", [stratum_cnt]" &
            ", [report_period]" &
            ", [car_own_mark]" &
            ", [car_lessee_mark]")
        mSQLStr.Append(", [car_cap]" &
            ", [nom_car_cap]" &
            ", [tare]" &
            ", [outside_l]" &
            ", [outside_w]" &
            ", [outside_h]" &
            ", [ex_outside_h]" &
            ", [type_wheel]" &
            ", [no_axles]" &
            ", [draft_gear]" &
            ", [art_units]" &
            ", [pool_code]" &
            ", [car_typ]" &
            ", [mech]")
        mSQLStr.Append(", [lic_st]" &
            ", [mx_wght_rail]" &
            ", [o_splc]" &
            ", [t_splc]")
        mSQLStr.Append(", [stcc]" &
            ", [orr_alpha]" &
            ", [jrr1_alpha]" &
            ", [jrr2_alpha]" &
            ", [jrr3_alpha]" &
            ", [jrr4_alpha]" &
            ", [jrr5_alpha]" &
            ", [jrr6_alpha]" &
            ", [trr_alpha]" &
            ", [jf]" &
            ", [exp_factor_th]" &
            ", [error_flg]" &
            ", [stb_car_typ]")
        mSQLStr.Append(", [err_code1]" &
            ", [err_code2]" &
            ", [err_code3]" &
            ", [car_own]" &
            ", [tofc_unit_type]")
        'check to see if we need to insert the deregulation date
        If Val(Mid(strInline, 372, 2)) <> 0 Then
            mSQLStr.Append(", [dereg_date]")
        End If
        mSQLStr.Append(", [dereg_flg]" &
            ", [service_type]" &
            ", [cars]" &
            ", [bill_wght_tons]" &
            ", [tons]" &
            ", [tc_units]" &
            ", [total_rev]" &
            ", [orr_rev]" &
            ", [jrr1_rev]" &
            ", [jrr2_rev]" &
            ", [jrr3_rev]" &
            ", [jrr4_rev]" &
            ", [jrr5_rev]" &
            ", [jrr6_rev]" &
            ", [trr_rev]")
        mSQLStr.Append(", [orr_dist]" &
            ", [jrr1_dist]" &
            ", [jrr2_dist]" &
            ", [jrr3_dist]" &
            ", [jrr4_dist]" &
            ", [jrr5_dist]" &
            ", [jrr6_dist]" &
            ", [trr_dist]" &
            ", [total_dist]" &
            ", [o_st]" &
            ", [jct1_st]" &
            ", [jct2_st]" &
            ", [jct3_st]" &
            ", [jct4_st]" &
            ", [jct5_st]" &
            ", [jct6_st]" &
            ", [jct7_st]")
        mSQLStr.Append(", [t_st]" &
            ", [o_bea]" &
            ", [t_bea]" &
            ", [o_fips]" &
            ", [t_fips]" &
            ", [o_fa]" &
            ", [t_fa]" &
            ", [o_ft]" &
            ", [t_ft]" &
            ", [o_smsa]" &
            ", [t_smsa]" &
            ", [onet]" &
            ", [net1]" &
            ", [net2]" &
            ", [net3]" &
            ", [net4]" &
            ", [net5]" &
            ", [net6]" &
            ", [net7]" &
            ", [tnet]")
        mSQLStr.Append(", [al_flg]" &
            ", [az_flg]" &
            ", [ar_flg]" &
            ", [ca_flg]" &
            ", [co_flg]" &
            ", [ct_flg]" &
            ", [de_flg]" &
            ", [dc_flg]" &
            ", [fl_flg]" &
            ", [ga_flg]" &
            ", [id_flg]" &
            ", [il_flg]" &
            ", [in_flg]" &
            ", [ia_flg]" &
            ", [ks_flg]" &
            ", [ky_flg]" &
            ", [la_flg]" &
            ", [me_flg]" &
            ", [md_flg]" &
            ", [ma_flg]" &
            ", [mi_flg]" &
            ", [mn_flg]" &
            ", [ms_flg]")
        mSQLStr.Append(", [mo_flg]" &
            ", [mt_flg]" &
            ", [ne_flg]" &
            ", [nv_flg]" &
            ", [nh_flg]" &
            ", [nj_flg]" &
            ", [nm_flg]" &
            ", [ny_flg]" &
            ", [nc_flg]" &
            ", [nd_flg]" &
            ", [oh_flg]" &
            ", [ok_flg]")
        mSQLStr.Append(", [or_flg]" &
            ", [pa_flg]" &
            ", [ri_flg]" &
            ", [sc_flg]" &
            ", [sd_flg]" &
            ", [tn_flg]" &
            ", [tx_flg]" &
            ", [ut_flg]" &
            ", [vt_flg]" &
            ", [va_flg]" &
            ", [wa_flg]" &
            ", [wv_flg]" &
            ", [wi_flg]" &
            ", [wy_flg]" &
            ", [cd_flg]" &
            ", [mx_flg]" &
            ", [othr_st_flg]")
        mSQLStr.Append(", [int_harm_code]" &
            ", [indus_class]" &
            ", [inter_sic]" &
            ", [dom_canada]" &
            ", [cs_54]" &
            ", [o_fs_type]" &
            ", [t_fs_type]" &
            ", [o_fs_ratezip]" &
            ", [t_fs_ratezip]" &
            ", [o_rate_splc]" &
            ", [t_rate_splc]" &
            ", [o_swlimit_splc]" &
            ", [t_swlimit_splc]")
        mSQLStr.Append(", [o_customs_flg]" &
            ", [t_customs_flg]" &
            ", [o_grain_flg]" &
            ", [t_grain_flg]" &
            ", [o_ramp_code]" &
            ", [t_ramp_code]" &
            ", [o_im_flg]" &
            ", [t_im_flg]" &
            ", [transborder_flg]")
        mSQLStr.Append(", [orr_cntry]" &
            ", [jrr1_cntry]" &
            ", [jrr2_cntry]" &
            ", [jrr3_cntry]" &
            ", [jrr4_cntry]" &
            ", [jrr5_cntry]" &
            ", [jrr6_cntry]" &
            ", [trr_cntry]" &
            ", [u_fuel_surchrg]")
        mSQLStr.Append(", [o_census_reg]" &
            ", [t_census_reg]" &
            ", [exp_factor]" &
            ", [total_vc]" &
            ", [rr1_vc]" &
            ", [rr2_vc]" &
            ", [rr3_vc]" &
            ", [rr4_vc]" &
            ", [rr5_vc]" &
            ", [rr6_vc]" &
            ", [rr7_vc]" &
            ", [rr8_vc]" &
            ", [Tracking_No]")


        mSQLStr.Append(") VALUES (")            'do not delete

        ' serial_no
        mSQLStr.Append("'" & Mid(strInline, 1, 6) & "',")
        ' wb_no
        mSQLStr.Append(Mid(strInline, 7, 6) & ",")
        ' wb_date
        mSQLStr.Append("'" & mDate_Str & "', ")
        ' acct_period
        ' This gets the month and year value
        workstr = Trim(Mid(strInline, 21, 6))
        workstr = Replace(workstr, " ", "0")
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' u_cars
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 27, 4)) & ",")
        ' u_car_init
        mSQLStr.Append("'" & Trim(Mid(strInline, 31, 4)) & "', ")
        ' u_car_num
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 35, 6)) & ",")
        ' tofc_serv_code
        mSQLStr.Append("'" & Trim(Mid(strInline, 41, 1)) & "', ")
        ' u_tc_units
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 44, 4)) & ",")
        ' u_tc_init
        mSQLStr.Append("'" & Trim(Mid(strInline, 48, 4)) & "', ")
        ' u_tc_num
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 52, 6)) & ",")
        ' stcc_w49
        mSQLStr.Append("'" & mSTCC_W49 & "', ")
        ' bill_wght
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 65, 9)) & ",")
        ' act_wght
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 74, 9)) & ",")
        ' u_rev
        mSQLStr.Append(CStr(mNew_U_Rev) & ",")
        ' tran_chrg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 92, 9)) & ",")
        ' misc_chrg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 101, 9)) & ",")
        ' intra_state_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 110, 1)) & ",")
        ' transit_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 111, 1)) & ",")
        ' all_rail_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 112, 1)) & ",")
        ' type_move
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 113, 1)) & ",")
        ' move_via_water
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 114, 1)) & ",")
        ' truck_for_rail
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 115, 1)) & ",")
        ' Shortline_Miles
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 116, 4)) & ",")
        ' rebill
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 120, 1)) & ",")
        ' stratum
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 121, 1)) & ",")
        ' subsample
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 122, 1)) & ",")
        ' int_eq_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 123, 1)) & ",")
        ' rate_flg
        mSQLStr.Append(CStr(mRate_Flg) & ",")
        ' wb_id
        mSQLStr.Append("'" & Trim(Mid(strInline, 125, 25)) & "', ")
        ' report_rr
        mSQLStr.Append(CStr(mReport_RR) & ",")
        ' o_fsac
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 153, 5)) & ",")
        ' orr
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 158, 3)) & ",")
        ' jct1
        mSQLStr.Append("'" & Trim(Mid(strInline, 161, 5)) & "', ")
        ' jrr1
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 166, 3)) & ",")
        ' jct2
        mSQLStr.Append("'" & Trim(Mid(strInline, 169, 5)) & "', ")
        ' jrr2
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 174, 3)) & ",")
        ' jct3
        mSQLStr.Append("'" & Trim(Mid(strInline, 177, 5)) & "', ")
        ' jrr3
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 182, 3)) & ",")
        ' jct4
        mSQLStr.Append("'" & Trim(Mid(strInline, 185, 5)) & "', ")
        ' jrr4
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 190, 3)) & ",")
        ' jct5
        mSQLStr.Append("'" & Trim(Mid(strInline, 193, 5)) & "', ")
        ' jrr5
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 198, 3)) & ",")
        ' jct6
        mSQLStr.Append("'" & Trim(Mid(strInline, 201, 5)) & "', ")
        ' jrr6
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 206, 3)) & ",")
        ' jct7
        mSQLStr.Append("'" & Trim(Mid(strInline, 209, 5)) & "', ")
        ' trr
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 214, 3)) & ",")
        ' t_fsac
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 217, 5)) & ",")
        ' pop_cnt
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 222, 8)) & ",")
        ' stratum_cnt
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 230, 6)) & ",")
        ' report_period
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 236, 1)) & ",")
        ' car_own_mark
        mSQLStr.Append("'" & Trim(Mid(strInline, 237, 4)) & "', ")
        ' car_lessee_mark
        mSQLStr.Append("'" & Trim(Mid(strInline, 241, 4)) & "', ")
        ' car_cap
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 245, 5)) & ",")
        ' nom_car_cap
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 250, 3)) & ",")
        ' tare
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 253, 4)) & ",")
        ' outside_l
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 257, 5)) & ",")
        ' outside_w
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 262, 4)) & ",")
        ' outside_h
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 266, 4)) & ",")
        ' ex_outside_h
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 270, 4)) & ",")
        ' type_wheel
        mSQLStr.Append("'" & Trim(Mid(strInline, 274, 1)) & "', ")
        ' no_axles
        mSQLStr.Append("'" & Trim(Mid(strInline, 275, 1)) & "', ")
        ' draft_gear
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 276, 2)) & ",")
        ' art_units
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 278, 1)) & ",")
        ' pool_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 279, 7)) & ",")
        ' car_typ
        mSQLStr.Append("'" & Trim(Mid(strInline, 286, 4)) & "', ")
        ' mech
        mSQLStr.Append("'" & Trim(Mid(strInline, 290, 4)) & "', ")
        ' lic_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 294, 2)) & "', ")
        ' mx_wght_rail
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 296, 3)) & ",")
        ' o_splc
        workstr = Trim(Mid(strInline, 299, 6))
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' t_splc
        workstr = Trim(Mid(strInline, 305, 6))
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")

        ' stcc
        workstr = Trim(Mid(strInline, 311, 7))
        Do While Len(workstr) < 7
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' orr_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 318, 4)) & "', ")
        ' jrr1_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 322, 4)) & "', ")
        ' jrr2_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 326, 4)) & "', ")
        ' jrr3_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 330, 4)) & "', ")
        ' jrr4_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 334, 4)) & "', ")
        ' jrr5_slpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 338, 4)) & "', ")
        ' jrr6_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 342, 4)) & "', ")
        ' trr_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 346, 4)) & "', ")
        ' jf
        mSQLStr.Append(CStr(mJF) & ",")
        ' exp_factor_th
        mSQLStr.Append(CStr(mExp_Factor_Th) & ",")
        ' error_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 354, 1)) & "', ")
        ' stb_car_typ
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 355, 2)) & ",")
        ' err_code1
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 357, 2)) & ",")
        ' err_code2
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 359, 2)) & ",")
        ' err_code3
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 361, 2)) & ",")
        'car_own
        mSQLStr.Append("'" & Trim(Mid(strInline, 363, 1)) & "', ")
        'tofc_unit_type
        mSQLStr.Append("'" & Mid(strInline, 364, 4) & "', ")
        'build the date string for the deregulation date
        If Val(Mid(strInline, 372, 2)) <> 0 Then
            workstr = Mid(strInline, 372, 2) & "/" & Mid(strInline, 374, 2) & "/" & Mid(strInline, 368, 4)
            mSQLStr.Append("'" & workstr & "', ")
        End If

        'back to the grind
        ' dereg_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 376, 1)) & ",")
        ' service_type
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 377, 1)) & ",")
        ' cars
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 378, 6)) & ",")
        ' bill_wght_tons
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 384, 7)) & ",")
        'tons
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 391, 8)) & ",")
        ' tc_units
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 399, 6)) & ",")
        ' total_rev
        mSQLStr.Append(CStr(mNew_Total_Rev) & ",")
        ' orr_rev
        mSQLStr.Append(CStr(mNew_ORR_Rev) & ",")
        ' jrr1_rev
        mSQLStr.Append(CStr(mNew_JRR1_Rev) & ",")
        ' jrr2_rev
        mSQLStr.Append(CStr(mNew_JRR2_Rev) & ",")
        ' jrr3_rev
        mSQLStr.Append(CStr(mNew_JRR3_Rev) & ",")
        ' jrr4_rev
        mSQLStr.Append(CStr(mNew_JRR4_Rev) & ",")
        ' jrr5_rev
        mSQLStr.Append(CStr(mNew_JRR5_Rev) & ",")
        ' jrr6_rev
        mSQLStr.Append(CStr(mNew_JRR6_Rev) & ",")
        ' trr_rev
        mSQLStr.Append(CStr(mNew_TRR_Rev) & ",")
        ' orr_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 495, 5)) & ",")
        ' jrr1_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 500, 5)) & ",")
        ' jrr2_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 505, 5)) & ",")
        ' jrr3_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 510, 5)) & ",")
        ' jrr4_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 515, 5)) & ",")
        ' jrr5_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 520, 5)) & ",")
        ' jrr6_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 525, 5)) & ",")
        ' trr_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 530, 5)) & ",")
        ' total_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 535, 5)) & ",")
        ' o_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 540, 2)) & "', ")
        ' jct1_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 542, 2)) & "', ")
        ' jct2_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 544, 2)) & "', ")
        ' jct3_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 546, 2)) & "', ")
        ' jct4_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 548, 2)) & "', ")
        ' jct5_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 550, 2)) & "', ")
        ' jct6_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 552, 2)) & "', ")
        ' jct7_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 554, 2)) & "', ")
        ' t_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 556, 2)) & "', ")
        ' o_bea
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 558, 3)) & ",")
        ' t_bea
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 561, 3)) & ",")
        ' o_fips
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 564, 5)) & ",")
        ' t_fips
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 569, 5)) & ",")
        ' o_fa
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 574, 2)) & ",")
        ' t_fa
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 576, 2)) & ",")
        ' o_ft
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 578, 1)) & ",")
        ' t_ft
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 579, 1)) & ",")
        ' o_smsa
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 580, 4)) & ",")
        ' t_smsa
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 584, 4)) & ",")
        ' onet
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 588, 5)) & ",")
        ' net1
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 593, 5)) & ",")
        ' net2
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 598, 5)) & ",")
        ' net3
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 603, 5)) & ",")
        ' net4
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 608, 5)) & ",")
        ' net5
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 613, 5)) & ",")
        ' net6
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 618, 5)) & ",")
        ' net7
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 623, 5)) & ",")
        ' tnet
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 628, 5)) & ",")
        ' al_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 633, 1)) & ",")
        ' az_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 634, 1)) & ",")
        ' ar_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 635, 1)) & ",")
        ' ca_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 636, 1)) & ",")
        ' co_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 637, 1)) & ",")
        ' ct_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 638, 1)) & ",")
        ' de_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 639, 1)) & ",")
        ' dc_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 640, 1)) & ",")
        ' fl_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 641, 1)) & ",")
        ' ga_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 642, 1)) & ",")
        ' id_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 643, 1)) & ",")
        ' il_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 644, 1)) & ",")
        ' in_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 645, 1)) & ",")
        ' ia_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 646, 1)) & ",")
        ' ks_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 647, 1)) & ",")
        ' ky_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 648, 1)) & ",")
        ' la_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 649, 1)) & ",")
        ' me_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 650, 1)) & ",")
        ' md_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 651, 1)) & ",")
        ' ma_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 652, 1)) & ",")
        ' mi_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 653, 1)) & ",")
        ' mn_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 654, 1)) & ",")
        ' ms_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 655, 1)) & ",")
        ' mo_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 656, 1)) & ",")
        ' mt_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 657, 1)) & ",")
        ' ne_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 658, 1)) & ",")
        ' nv_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 659, 1)) & ",")
        ' nh_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 660, 1)) & ",")
        ' nj_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 661, 1)) & ",")
        ' nm_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 662, 1)) & ",")
        ' ny_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 663, 1)) & ",")
        ' nc_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 664, 1)) & ",")
        ' nd_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 665, 1)) & ",")
        ' oh_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 666, 1)) & ",")
        ' ok_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 667, 1)) & ",")
        ' or_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 668, 1)) & ",")
        ' pa_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 669, 1)) & ",")
        ' ri_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 670, 1)) & ",")
        ' sc_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 671, 1)) & ",")
        ' sd_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 672, 1)) & ",")
        ' tn_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 673, 1)) & ",")
        ' tx_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 674, 1)) & ",")
        'ut_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 675, 1)) & ",")
        ' vt_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 676, 1)) & ",")
        ' va_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 677, 1)) & ",")
        ' wa_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 678, 1)) & ",")
        ' wv_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 679, 1)) & ",")
        ' wi_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 680, 1)) & ",")
        ' wy_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 681, 1)) & ",")
        ' cd_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 682, 1)) & ",")
        ' mx_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 683, 1)) & ",")
        ' othr_st_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 684, 1)) & ",")
        ' int_harm_code
        mSQLStr.Append("'" & Trim(Mid(strInline, 685, 12)) & "', ")
        ' indus_class
        mSQLStr.Append("'" & Trim(Mid(strInline, 697, 4)) & "', ")
        ' inter_sic
        mSQLStr.Append("'" & Trim(Mid(strInline, 701, 4)) & "', ")
        ' dom_canada
        mSQLStr.Append("'" & Trim(Mid(strInline, 705, 3)) & "', ")
        ' cs_54
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 708, 2)) & ",")
        ' o_fs_type
        mSQLStr.Append("'" & Trim(Mid(strInline, 710, 4)) & "', ")
        ' t_fs_type
        mSQLStr.Append("'" & Trim(Mid(strInline, 714, 4)) & "', ")
        ' o_fs_ratezip
        mSQLStr.Append("'" & Trim(Mid(strInline, 718, 9)) & "', ")
        ' t_fs_ratezip
        mSQLStr.Append("'" & Trim(Mid(strInline, 727, 9)) & "', ")
        ' o_rate_splc
        mSQLStr.Append("'" & Trim(Mid(strInline, 736, 9)) & "', ")
        ' t_rate_splc
        mSQLStr.Append("'" & Trim(Mid(strInline, 745, 9)) & "', ")
        ' o_swlimit_splc
        mSQLStr.Append("'" & Trim(Mid(strInline, 754, 9)) & "', ")
        ' t_swlimit_splc
        mSQLStr.Append("'" & Trim(Mid(strInline, 763, 9)) & "', ")
        ' o_customs_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 772, 1)) & "', ")
        ' t_customs_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 773, 1)) & "', ")
        ' o_grain_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 774, 1)) & "', ")
        ' t_grain_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 775, 1)) & "', ")
        ' o_ramp_code
        mSQLStr.Append("'" & Trim(Mid(strInline, 776, 1)) & "', ")
        ' t_ramp_code
        mSQLStr.Append("'" & Trim(Mid(strInline, 777, 1)) & "', ")
        ' o_im_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 778, 1)) & "', ")
        ' t_im_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 779, 1)) & "', ")
        ' transborder_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 780, 1)) & "', ")
        ' orr_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 781, 2)) & "', ")
        ' jrr1_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 783, 2)) & "', ")
        ' jrr2_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 785, 2)) & "', ")
        ' jrr3_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 787, 2)) & "', ")
        ' jrr4_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 789, 2)) & "', ")
        ' jrr5_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 791, 2)) & "', ")
        ' jrr6_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 793, 2)) & "', ")
        ' trr_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 795, 2)) & "', ")
        ' u_fuel_surchrg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 797, 9)) & ",")
        ' there are 13 blanks columns at this point
        ' o_census_reg
        mSQLStr.Append("'" & Trim(Mid(strInline, 819, 4)) & "', ")
        ' t_census_reg
        mSQLStr.Append("'" & Trim(Mid(strInline, 823, 4)) & "', ")
        ' exp_factor
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 827, 7)) & ",")
        ' total_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 834, 8)) & ",")
        ' rr1_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 842, 8)) & ",")
        ' rr2_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 850, 8)) & ",")
        ' rr3_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 858, 8)) & ",")
        ' rr4_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 866, 7)) & ",")
        ' rr5_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 873, 7)) & ",")
        ' rr6_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 880, 7)) & ",")
        ' rr7_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 887, 7)) & ",")
        ' rr8_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 894, 7)) & ",")
        'Tracking_No
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 901, 13)))

        mSQLStr.Append(")")
        Build913SQL = mSQLStr.ToString

    End Function

    Function Build_Segment_SQL(
    ByVal mTableName As String,
    ByVal mSerial_No As String,
    ByVal mSeg_no As Byte,
    ByVal mTotal_Segs As Byte,
    ByVal mRR_Num As Integer,
    ByVal mRR_Alpha As String,
    ByVal mRR_Dist As Object,
    ByVal mRR_Cntry As String,
    ByVal mRR_Rev As Object,
    ByVal mRR_VC As Object,
    ByVal mShip_Type As String,
    ByVal mFrom_Node As Object,
    ByVal mTo_Node As Object,
    ByVal mFrom_Loc As Object,
    ByVal mFrom_St As String,
    ByVal mTo_Loc As Object,
    ByVal mTo_St As String) As StringBuilder

        Build_Segment_SQL = New StringBuilder
        Build_Segment_SQL.Append("INSERT INTO " & mTableName & "(" &
            "Serial_No, " &
            "Seg_No, " &
            "Total_Segs, " &
            "RR_Num, " &
            "RR_Alpha, " &
            "RR_Dist, " &
            "RR_Cntry, " &
            "RR_Rev, " &
            "RR_VC, " &
            "Seg_Type, " &
            "From_Node, " &
            "To_Node, " &
            "From_Loc, " &
            "From_St, " &
            "To_Loc, " &
            "To_St) VALUES ('")

        Build_Segment_SQL.Append(mSerial_No & "', ")
        Build_Segment_SQL.Append(mSeg_no & ", ")
        Build_Segment_SQL.Append(mTotal_Segs & ", ")
        Build_Segment_SQL.Append(mRR_Num & ", ")
        Build_Segment_SQL.Append("'" & mRR_Alpha & "', ")
        Build_Segment_SQL.Append(mRR_Dist & ", ")
        Build_Segment_SQL.Append("'" & mRR_Cntry & "', ")
        Build_Segment_SQL.Append(mRR_Rev & ", ")
        Build_Segment_SQL.Append(mRR_VC & ", ")
        Build_Segment_SQL.Append("'" & mShip_Type & "', ")
        Build_Segment_SQL.Append(mFrom_Node & ", ")
        Build_Segment_SQL.Append(mTo_Node & ", ")
        Build_Segment_SQL.Append("'" & mFrom_Loc & "', ")
        Build_Segment_SQL.Append("'" & mFrom_St & "', ")
        Build_Segment_SQL.Append("'" & mTo_Loc & "', ")
        Build_Segment_SQL.Append("'" & mTo_St & "')")

    End Function

    Function Build_Unmasked_Segment_SQL(
    ByVal mTableName As String,
    ByVal mSerial_No As String,
    ByVal mSeg_no As Byte,
    ByVal mRR_Unmasked_Rev As Single) As String

        Build_Unmasked_Segment_SQL = "INSERT INTO " & mTableName & "(" &
            "Serial_No, " &
            "Seg_No, " &
            "RR_Unmasked_Rev) VALUES ("

        Build_Unmasked_Segment_SQL = Build_Unmasked_Segment_SQL & "'" & mSerial_No & "', "
        Build_Unmasked_Segment_SQL = Build_Unmasked_Segment_SQL & mSeg_no & ", "
        Build_Unmasked_Segment_SQL = Build_Unmasked_Segment_SQL & mRR_Unmasked_Rev & ")"

    End Function

    Function Get_URCS_Code(ByRef mURCS_Schedule As String) As String
        Dim mSQLstr As String
        Dim mDataTable As New DataTable

        Gbl_URCS_Codes_TableName = Get_Table_Name_From_SQL("1", "URCS_CODES")

        ' open the connection to the Controls database
        OpenSQLConnection(Gbl_Controls_Database_Name)

        mSQLstr = "Select SCH from " & Gbl_URCS_Codes_TableName & " WHERE Schedule LIKE '" & mURCS_Schedule & "'"

        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        Get_URCS_Code = mDataTable.Rows(0)("SCH")

        mDataTable = Nothing

    End Function

    Function Get_URCS_Schedule(ByRef mURCS_Code As String) As String
        Dim mSQLstr As String
        Dim mDataTable As New DataTable

        Gbl_URCS_Schedules_TableName = Get_Table_Name_From_SQL("1", "URCS_Schedules")

        ' open the connection to the Controls database
        OpenSQLConnection(Gbl_Controls_Database_Name)

        mSQLstr = "Select Schedule from " & Gbl_URCS_Schedules_TableName & " WHERE Sch =" & mURCS_Code & ""

        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        Get_URCS_Schedule = mDataTable.Rows(0)("Schedule")

        mDataTable = Nothing

    End Function

    Function Get_URCS_Columns(ByRef mURCS_Code As String) As Integer
        Dim mSQLstr As String
        Dim mDataTable As New DataTable

        Gbl_URCS_Schedules_TableName = Get_Table_Name_From_SQL("1", "URCS_Schedules")

        ' open the connection to the Controls database
        OpenSQLConnection(Gbl_Controls_Database_Name)

        mSQLstr = "Select Num_Cols from " & Gbl_URCS_Schedules_TableName & " WHERE Sch =" & mURCS_Code & ""

        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        Get_URCS_Columns = mDataTable.Rows(0)("Num_Cols")

        mDataTable = Nothing

    End Function

    Function Get_Trans_Values_Sum(
        ByVal mYear As String,
        ByVal mRRICC As String,
        ByVal mSch As String,
        ByVal mColumn As String,
        ByVal mStartLine As String,
        Optional ByVal mEndLine As String = "") As Long

        Dim mDataTable As New DataTable
        Dim mSQLStr As String

        Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "TRANS")
        Gbl_Trans_DatabaseName = Get_Database_Name_From_SQL("1", "TRANS")

        'Open the connection to the database
        OpenSQLConnection(Gbl_Controls_Database_Name)

        mColumn = ConvertColumn(mColumn)

        ' Get the sum of the values
        If mEndLine = "" Then
            mSQLStr = "Select Sum(" & mColumn & ") as Total from " & Global_Variables.Gbl_Trans_TableName &
                 " WHERE Year = " & mYear & " AND " &
                 "RRICC = " & mRRICC & " AND " &
                 "SCH = " & mSch & " AND " &
                 "Line = " & mStartLine
        Else
            mSQLStr = "Select Sum(" & mColumn & ") as Total from " & Global_Variables.Gbl_Trans_TableName &
                 " WHERE Year = " & mYear & " AND " &
                 "RRICC = " & mRRICC & " AND " &
                 "SCH = " & mSch & " AND " &
                 "Line >= " & mStartLine & " AND " &
                 "Line <= " & mEndLine
        End If

        Using daAdapter As New SqlDataAdapter(mSQLStr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        Get_Trans_Values_Sum = mDataTable.Rows(0)("Total")

        mDataTable = Nothing

    End Function

    Function Get_Trans_Value(
        ByVal mYear As String,
        ByVal mRRICC As String,
        ByVal mSch As String,
        ByVal Mline As String,
        ByVal mColumn As String) As Long

        Dim mDataTable As New DataTable
        Dim mSQLStr As String

        Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "TRANS")
        Gbl_Trans_DatabaseName = Get_Database_Name_From_SQL("1", "TRANS")

        'Open the connection to the database
        OpenSQLConnection(Gbl_Trans_DatabaseName)

        mSQLStr = "Select " & mColumn & " from " & Gbl_Trans_TableName &
             " WHERE Year = " & mYear & " AND " &
             "RRICC = " & mRRICC & " AND " &
             "SCH = " & mSch & " AND " &
             "Line = " & Mline

        Using daAdapter As New SqlDataAdapter(mSQLStr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        If mDataTable.Rows.Count = 0 Then
            Get_Trans_Value = 0
        Else
            Get_Trans_Value = mDataTable.Rows(0)(mColumn)
        End If

        mDataTable = Nothing

    End Function

    Function Get_RRICC_By_Short_Name(ByVal mShort_Name As String) As String

        Dim mDataTable As New DataTable
        Dim mStrSQL As String

        Get_RRICC_By_Short_Name = ""

        If mShort_Name = "CSX" Then
            mShort_Name = "CSXT"
        End If

        Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "CLASS1RAILLIST")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        mStrSQL = "SELECT rricc FROM " & Gbl_Class1RailList_TableName & " WHERE short_name = '" & mShort_Name & "'"

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        Get_RRICC_By_Short_Name = mDataTable.Rows(0)("rricc")

        mDataTable = Nothing

    End Function

    Function Get_Short_Name_By_RRICC(ByVal mRRICC As String) As String

        Dim mDataTable As New DataTable
        Dim mStrSQL As String

        Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "Class1RailList")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        mStrSQL = "SELECT short_name FROM " & Gbl_Class1RailList_TableName & " WHERE rricc = '" & mRRICC & "'"

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        Get_Short_Name_By_RRICC = mDataTable.Rows(0)("short_name")

        mDataTable = Nothing

    End Function

    Function Get_Short_Name_By_RRID(ByVal mRRID As String) As String

        Dim mDataTable As New DataTable
        Dim mStrSQL As String

        Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "Class1RailList")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        mStrSQL = "SELECT short_name FROM " & Gbl_Class1RailList_TableName & " WHERE rr_id = '" & mRRID & "'"

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        Get_Short_Name_By_RRID = mDataTable.Rows(0)("short_name")

        mDataTable = Nothing

    End Function

    Function Get_RR_Name_By_RRID(ByVal mRRID As String) As String

        Dim mDataTable As New DataTable
        Dim mStrSQL As String

        Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "Class1RailList")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        mStrSQL = "SELECT name FROM " & Gbl_Class1RailList_TableName & " WHERE rr_id = '" & mRRID & "'"

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        Get_RR_Name_By_RRID = mDataTable.Rows(0)("name")

        mDataTable = Nothing

    End Function

    Function Get_ECode_Id_By_ECode(ByVal m_eCode As String) As String

        Dim mDataTable As New DataTable
        Dim mStrSQL As String

        Get_ECode_Id_By_ECode = ""

        Gbl_ECode_TableName = Get_Table_Name_From_SQL("1", "ECODES")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        mStrSQL = "SELECT ecode_id FROM " & Gbl_ECode_TableName & " WHERE ecode = '" & m_eCode & "'"

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        Get_ECode_Id_By_ECode = mDataTable.Rows(0)("ecode_id")

        mDataTable = Nothing

    End Function

    Function Get_STCC_W49_Translation(ByVal mSTCC_W49 As String) As String

        Dim mDataTable As New DataTable
        Dim mStrSQL As String

        Get_STCC_W49_Translation = ""

        Gbl_STCC_W49_Translation_TableName = Get_Table_Name_From_SQL("1", "STCC_49_Translation")

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        mStrSQL = "SELECT STCC7 FROM " & Gbl_STCC_W49_Translation_TableName & " WHERE " & Gbl_STCC_W49_Translation_TableName & ".STCC49 = '" & mSTCC_W49 & "'"

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        If mDataTable.Rows.Count < 1 Then
            ' Value not found, return the passed value 
            Get_STCC_W49_Translation = mSTCC_W49
        Else
            Get_STCC_W49_Translation = mDataTable.Rows(0)("STCC7")
        End If

        mDataTable = Nothing

    End Function

    Public Sub Write_XML_Data_SQL(
        ByVal mYear As String,
        ByVal mRRICC As String,
        ByVal mXML_Name As String,
        ByVal mXML_Title As String,
        ByVal meCode_ID As String,
        ByVal meCode As String,
        ByVal mValue As String)

        Dim mDatatable As DataTable
        Dim mCommand As New SqlCommand

        Dim mTable_Name As String
        Dim mDatabase_Name As String
        Dim mSQLStr As String
        Dim mComma As String
        Dim mDValue As Double

        mComma = ", "
        mDValue = Convert.ToDouble(mValue)

        'Get the database_name and table_name for the CLASS1RAILLIST value from the database
        gbl_Table_Name = Get_Table_Name_From_SQL(1, "XML_DATA")
        mDatabase_Name = "URCS_CONTROLS"
        mTable_Name = gbl_Table_Name

        ' Open/Check the SQL connection
        OpenSQLConnection(mDatabase_Name)

        mDatatable = New DataTable

        ' Determine if the record already exists
        mSQLStr = "SELECT Count(*) As MyCount FROM " & mTable_Name & " " &
            "WHERE Year = " & mYear & " AND " &
            "RRICC = " & mRRICC & " AND " &
            "ECODE = '" & meCode & "'"

        Using daAdapter As New SqlDataAdapter(mSQLStr, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        If mDatatable.Rows(0)("MyCount") = 0 Then
            'We have to insert the record
            mSQLStr = "INSERT INTO " & mTable_Name & " (" &
            "Year, rricc, xml_name, xml_title, eCode_id, ecode, value " &
            ") VALUES ("
            mSQLStr = mSQLStr & mYear & mComma
            mSQLStr = mSQLStr & mRRICC & mComma
            mSQLStr = mSQLStr & "'" & mXML_Name & "'" & mComma
            mSQLStr = mSQLStr & "'" & mXML_Title & "'" & mComma
            mSQLStr = mSQLStr & meCode_ID & mComma
            mSQLStr = mSQLStr & "'" & meCode & "'" & mComma
            mSQLStr = mSQLStr & Format(mDValue, "F20") & ")"
        Else
            'We have to update it
            mSQLStr = "UPDATE " & mTable_Name & " " &
                "SET XML_Name = '" & mXML_Name & "'" & mComma &
                "XML_Title = '" & mXML_Title & "'" & mComma &
                "eCode_ID = " & meCode_ID & mComma &
                "value = " & Format(mDValue, "F20") & " " &
                "WHERE year = " & mYear & " AND " &
                "rricc = " & mRRICC & " AND " &
                "eCode = '" & meCode & "'"
        End If

        mCommand.Connection = gbl_SQLConnection
        mCommand.CommandType = CommandType.Text
        mCommand.CommandText = mSQLStr

        mCommand.ExecuteNonQuery()

    End Sub

    Function VerifyTableExist(ByRef mDatabaseName As String, ByVal mWBTableName As String) As Boolean

        Dim mDataTable As DataTable
        Dim strSQL As String

        ' Open/Check the SQL connection
        OpenSQLConnection(mDatabaseName)

        mDataTable = New DataTable

        strSQL = "SELECT * FROM dbo.sysobjects WHERE name = '" & mWBTableName & "' and xType = 'U'"

        Using daAdapter As New SqlDataAdapter(strSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        If mDataTable.Rows.Count = 0 Then
            VerifyTableExist = False
        Else
            VerifyTableExist = True
        End If

        mDataTable = Nothing

    End Function

    Function Count_Any_Table(ByRef mDatabaseName As String, ByRef mTableName As String) As String

        Dim mDataTable As DataTable
        Dim strSQL As String

        ' Open/Check the SQL connection
        OpenSQLConnection(mDatabaseName)
        OpenADOConnection(mDatabaseName)

        ' Determine if the table exists
        If VerifyTableExist(mDatabaseName, mTableName) = True Then

            mDataTable = New DataTable

            strSQL = "SELECT COUNT(*) AS MyCount from " & mTableName

            Using daAdapter As New SqlDataAdapter(strSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            Count_Any_Table = mDataTable.Rows(0)("MyCount").ToString
        Else
            Count_Any_Table = "0"
        End If

        mDataTable = Nothing

    End Function

    Function Is_Record_In_Table(ByRef mDatabaseName As String,
                                        ByRef mTableName As String,
                                        ByRef mKeyField As String,
                                        ByRef mValue As String) As Boolean

        Dim mDataTable As DataTable
        Dim strSQL As String

        Is_Record_In_Table = False

        ' Open/Check the SQL connection
        OpenSQLConnection(mDatabaseName)

        ' Determine if the table exists
        If VerifyTableExist(mDatabaseName, mTableName) = True Then

            mDataTable = New DataTable

            strSQL = "SELECT * from " & mTableName & " WHERE " & mKeyField & " = '" & mValue & "'"

            Using daAdapter As New SqlDataAdapter(strSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count = 0 Then
                Is_Record_In_Table = False
            Else
                Is_Record_In_Table = True
            End If
        End If

        mDataTable = Nothing

    End Function

    Function Count_Ordnance_Records_By_STCC(ByRef mDatabaseName As String, ByRef mTableName As String, ByRef mSTCC As String) As Integer

        Dim mDataTable As DataTable
        Dim mSQLstr As String

        mDataTable = New DataTable

        ' Get the table and database name from the Table Locator table
        mTableName = Get_Table_Name_From_SQL("1", "ORDNANCE_STCCS")
        mDatabaseName = Get_Database_Name_From_SQL("1", "ORDNANCE_STCCS")

        OpenSQLConnection(mDatabaseName)

        ' Build the SQL statement
        mSQLstr = "select COUNT(*) as MyCount from " & mTableName & " WHERE STCC = " & mSTCC

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        Count_Ordnance_Records_By_STCC = mDataTable.Rows(0)("MyCount")

    End Function

    Function Count_SQL_Records_By_FieldName(ByRef mYear As String, ByRef mDatabaseName As String, ByRef mTableName As String, ByRef mFieldName As String, ByRef mValue As String) As Integer

        Dim mDataTable As DataTable
        Dim mSQLstr As String

        ' Open/Check the SQL connection
        OpenSQLConnection(mDatabaseName)

        mDataTable = New DataTable

        ' Get the table and database name from the Table Locator table
        mTableName = Get_Table_Name_From_SQL(mYear, mFieldName)
        mDatabaseName = Get_Database_Name_From_SQL(mYear, mFieldName)

        OpenSQLConnection(mDatabaseName)

        ' Build the SQL statement
        mSQLstr = "select COUNT(*) as MyCount from " & mTableName & " WHERE " & mFieldName & " = " & mValue

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        Count_SQL_Records_By_FieldName = mDataTable.Rows(0)("MyCount")

    End Function

    Function Count_MakeWhole_Factors_Table(ByRef Year As String, ByRef RR_Id As String, ByRef eCode_Id As String) As Integer

        Dim mDataTable As DataTable
        Dim strSQL As String

        Gbl_Makewhole_Factors_Tablename = Get_Table_Name_From_SQL(My.Settings.CurrentYear.ToString, "MAKEWHOLE_FACTORS")
        gbl_Database_Name = Get_Database_Name_From_SQL(My.Settings.CurrentYear.ToString, "MAKEWHOLE_FACTORS")

        ' Open/Check the SQL connection
        OpenSQLConnection(gbl_Database_Name)

        If VerifyTableExist(gbl_Database_Name, Gbl_Makewhole_Factors_Tablename) Then

            mDataTable = New DataTable

            strSQL = "SELECT COUNT(*) AS MyCount from " & Global_Variables.Gbl_Makewhole_Factors_Tablename &
                " WHERE Year = " & Year.ToString & " AND RR_ID = " & RR_Id.ToString & " AND eCODE_Id = " & eCode_Id.ToString

            Using daAdapter As New SqlDataAdapter(strSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            Count_MakeWhole_Factors_Table = mDataTable.Rows(0)("Mycount")
        Else
            Count_MakeWhole_Factors_Table = 0
        End If

        mDataTable = Nothing

    End Function

    Function Build_Insert_HazMat_STCC_Rec(ByVal mInString As String, ByVal mTableName As String) As StringBuilder
        Dim mStrSQL As StringBuilder
        Dim mThisStr As String

        mStrSQL = New StringBuilder
        mStrSQL.Append("INSERT INTO " & mTableName & " (")
        mStrSQL.Append("STCC,")
        mStrSQL.Append("STCC_Code2,")
        mStrSQL.Append("Transaction_Date,")
        mStrSQL.Append("Transaction_Time,")
        mStrSQL.Append("Status_Code,")
        mStrSQL.Append("Effective_Date,")
        mStrSQL.Append("Header1,")
        mStrSQL.Append("Header2,")
        mStrSQL.Append("Header3,")
        mStrSQL.Append("Header4,")
        mStrSQL.Append("STCC_Desc_15,")
        mStrSQL.Append("Alternate_Num,")
        mStrSQL.Append("Expiration_Date,")
        mStrSQL.Append("Description_US,")
        mStrSQL.Append("Primary_Hazard_Class_Intl,")
        mStrSQL.Append("NOS_Flg_Intl,")
        mStrSQL.Append("Technical_Name_Intl,")
        mStrSQL.Append("UN_Number_Intl,")
        mStrSQL.Append("Packing_Grp_Intl,")
        mStrSQL.Append("Poison_Flg_intl,")
        mStrSQL.Append("Primary_Placard_Intl,")
        mStrSQL.Append("Shipping_Name_Intl,")
        mStrSQL.Append("Primary_Class_CN,")
        mStrSQL.Append("Sub_Class_CN,")
        mStrSQL.Append("CN_Orig_US_Dest_Flg,")
        mStrSQL.Append("Emer_Resp_Asst_Plan_CN_Flg,")
        mStrSQL.Append("Primary_Placard_CN,")
        mStrSQL.Append("Special_Commodity_CN_Flg,")
        mStrSQL.Append("NOS_Flag_CN,")
        mStrSQL.Append("Sub_Placard_CN,")
        mStrSQL.Append("Technical_Name_CN,")
        mStrSQL.Append("UN_Number_CN,")
        mStrSQL.Append("Packing_Grp_CN,")
        mStrSQL.Append("Poison_Flg_CN,")
        mStrSQL.Append("Shipping_Name_CN,")
        mStrSQL.Append("EPA_Waste_Stream_Num_US,")
        mStrSQL.Append("Haz_Placard_US,")
        mStrSQL.Append("Primary_Hazard_Class_US,")
        mStrSQL.Append("Sub_Hazard_US,")
        mStrSQL.Append("Hazard_Zone_US,")
        mStrSQL.Append("NOS_Flag_US,")
        mStrSQL.Append("Sub_Placard_US,")
        mStrSQL.Append("Technical_Name_US,")
        mStrSQL.Append("UN_NA_ID_Num_US,")
        mStrSQL.Append("US_Orig_CN_Dest_Flg,")
        mStrSQL.Append("Packing_Grp_US,")
        mStrSQL.Append("Poison_Flg_US,")
        mStrSQL.Append("Primary_Placard_US,")
        mStrSQL.Append("Shipping_Name_US,")
        mStrSQL.Append("Alpha_Desc,")
        mStrSQL.Append("OT55_Flg,")
        mStrSQL.Append("Reportable_Quantity_Flg_US,")
        mStrSQL.Append("Marine_Pollutant_Flg_US,")
        mStrSQL.Append("HazMat_Name_US,")
        mStrSQL.Append("Marine_Pollutant_Name_US,")
        mStrSQL.Append("Special_Shipping_Name_Flg_CN,")
        mStrSQL.Append("Spcl_Proper_Ship_Name_Flg_Intl,")
        mStrSQL.Append("Spcl_Proper_Ship_Name_Flg_US,")
        mStrSQL.Append("Intermodal_Flg_CN,")
        mStrSQL.Append("Internodal_Flg_Intl,")
        mStrSQL.Append("Intermodal_Flg_US,")
        mStrSQL.Append("RSSM_Flg_US,")
        mStrSQL.Append("STCC_Desc,")
        mStrSQL.Append("Sub_Placard_Intl,")
        mStrSQL.Append("HazMat_Class_Intl,")
        mStrSQL.Append("Deletion_Date,")
        mStrSQL.Append("Alt_Shipping_Name_CN,")
        mStrSQL.Append("Alt_Proper_Ship_Name_Intl,")
        mStrSQL.Append("Alt_Proper_Ship_Name_US")
        mStrSQL.Append(") VALUES (")
        mStrSQL.Append("'" & Trim(mInString.Substring(0, 7)) & "',") 'Hazmat Response Code/STCC Code
        mStrSQL.Append("'" & Trim(mInString.Substring(7, 7)) & "',") 'STCC Code
        mStrSQL.Append("'" & mInString.Substring(18, 2) & "/" &
                                  mInString.Substring(20, 2) & "/" &
                                  mInString.Substring(14, 4) & "',") 'Transaction Date
        mStrSQL.Append("'" & mInString.Substring(22, 2) & ":" &
            mInString.Substring(24, 2) & ":" &
            mInString.Substring(26, 2) & "',") 'Transaction Time
        mStrSQL.Append("'" & Trim(mInString.Substring(28, 1)) & "',") 'Status Code
        mStrSQL.Append("'" & mInString.Substring(33, 2) & "/" &
                                  mInString.Substring(35, 2) & "/" &
                                  mInString.Substring(29, 4) & "',") 'Effective Date
        mStrSQL.Append("'" & Trim(mInString.Substring(37, 2)) & "',") 'Header1
        mStrSQL.Append("'" & Trim(mInString.Substring(39, 3)) & "',") 'Header2
        mStrSQL.Append("'" & Trim(mInString.Substring(42, 4)) & "',") 'Header3
        mStrSQL.Append("'" & Trim(mInString.Substring(46, 5)) & "',") 'Header4
        mStrSQL.Append("'" & Trim(mInString.Substring(51, 15)) & "',") 'STCC Description 15
        mStrSQL.Append("'" & Trim(mInString.Substring(66, 2)) & "',") 'Alt Number
        mStrSQL.Append("'" & mInString.Substring(72, 2) & "/" &
                                  mInString.Substring(74, 2) & "/" &
                                  mInString.Substring(68, 4) & "',") 'Expiration Date
        mThisStr = Trim(RemoveSpaces(mInString.Substring(76, 250)))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & mThisStr & "',") 'Addl Desc US
        mStrSQL.Append("'" & Trim(mInString.Substring(326, 4)) & "',") 'Primary Hazard Class Intl
        mStrSQL.Append("'" & Trim(mInString.Substring(330, 1)) & "',") 'NOS Flg Intl
        mThisStr = Trim(RemoveSpaces(mInString.Substring(331, 125)))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & mThisStr & "',") 'Technical Name Intl
        mStrSQL.Append("'" & Trim(mInString.Substring(456, 6)) & "',") 'UN Number Intl
        mStrSQL.Append("'" & Trim(mInString.Substring(462, 1)) & "',") 'Packing Grp Intl
        mStrSQL.Append("'" & Trim(mInString.Substring(463, 1)) & "',") 'Poison Flg Intl
        mStrSQL.Append("'" & Trim(mInString.Substring(464, 2)) & "',") 'Primary Placard Intl
        mThisStr = Trim(RemoveSpaces(mInString.Substring(466, 125)))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & mThisStr & "',") 'Proper Shipping Name Intl
        mStrSQL.Append("'" & Trim(mInString.Substring(591, 4)) & "',") 'Primary Class CN
        mStrSQL.Append("'" & Trim(mInString.Substring(595, 9)) & "',") 'Sub Class CN
        mStrSQL.Append("'" & Trim(mInString.Substring(604, 1)) & "',") 'CN Orig US Dest Flg
        mStrSQL.Append("'" & Trim(mInString.Substring(605, 4)) & "',") 'Emergency Response Assistance Plan Flg CN
        mStrSQL.Append("'" & Trim(mInString.Substring(609, 2)) & "',") 'Primary Placard CN
        mStrSQL.Append("'" & Trim(mInString.Substring(611, 1)) & "',") 'Specialized Commodity Flg CN
        ''612 is reserved/blank
        mStrSQL.Append("'" & Trim(mInString.Substring(613, 1)) & "',") 'NOS Flg CN
        mStrSQL.Append("'" & Trim(mInString.Substring(614, 2)) & "',") 'Sub Plcard CN
        mThisStr = Trim(RemoveSpaces(mInString.Substring(616, 125)))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & mThisStr & "',") 'Technical Name CN
        mStrSQL.Append("'" & Trim(mInString.Substring(741, 6)) & "',") 'UN Number CN
        mStrSQL.Append("'" & Trim(mInString.Substring(747, 1)) & "',") 'Packing Grp CN
        mStrSQL.Append("'" & Trim(mInString.Substring(748, 1)) & "',") 'Poison Flg CN
        mThisStr = Trim(RemoveSpaces(mInString.Substring(749, 125)))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & mThisStr & "',") 'Shipping Name CN
        ''874-876 are reserved/blank
        mStrSQL.Append("'" & Trim(mInString.Substring(877, 18)) & "',") 'EPA Waste Stream Numbers US
        mStrSQL.Append("'" & Trim(mInString.Substring(895, 2)) & "',") 'Hazardous Placard US
        mStrSQL.Append("'" & Trim(mInString.Substring(897, 4)) & "',") 'Primary Hazard Class US
        mStrSQL.Append("'" & Trim(mInString.Substring(901, 6)) & "',") 'Sub Hazard US
        mStrSQL.Append("'" & Trim(mInString.Substring(907, 1)) & "',") 'Hazard Zone US
        mStrSQL.Append("'" & Trim(mInString.Substring(908, 1)) & "',") 'NOS Flg US
        mStrSQL.Append("'" & Trim(mInString.Substring(909, 2)) & "',") 'Sub Placard US
        mThisStr = Trim(RemoveSpaces(mInString.Substring(911, 125)))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & mThisStr & "',") 'Technical Name US
        mStrSQL.Append("'" & Trim(mInString.Substring(1036, 6)) & "',") 'UN/NA ID Number US
        mStrSQL.Append("'" & Trim(mInString.Substring(1042, 1)) & "',") ' US Orig CN Dest Flg
        mStrSQL.Append("'" & Trim(mInString.Substring(1043, 1)) & "',") 'Packing Grp US
        mStrSQL.Append("'" & Trim(mInString.Substring(1044, 1)) & "',") 'Poison Flg US
        mStrSQL.Append("'" & Trim(mInString.Substring(1045, 2)) & "',") 'Primary Placard US
        mThisStr = Trim(RemoveSpaces(mInString.Substring(1047, 125)))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & mThisStr & "',") 'Shipping Name US
        mStrSQL.Append("'" & mThisStr & "',") 'Alpha Desc
        mStrSQL.Append("'" & Trim(mInString.Substring(1172, 1)) & "',") 'OT55 Flg
        ''1173 is reserved/blank
        mStrSQL.Append("'" & Trim(mInString.Substring(1174, 1)) & "',") 'Reportable Quantity Flg US
        mStrSQL.Append("'" & Trim(mInString.Substring(1175, 1)) & "',") 'Marine Pollutant Flg US
        mThisStr = Trim(RemoveSpaces(mInString.Substring(1176, 125)))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & mThisStr & "',") 'Hazardous Substance Name US
        mThisStr = Trim(RemoveSpaces(mInString.Substring(1301, 125)))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & mThisStr & "',") 'Marine Pollutant Name US
        ''1426-1429 are reseved/blank
        mStrSQL.Append("'" & Trim(mInString.Substring(1430, 1)) & "',") 'Special Shipping Name Flg CN
        mStrSQL.Append("'" & Trim(mInString.Substring(1431, 1)) & "',") 'Special Proper Shipping Name Flg Intl
        mStrSQL.Append("'" & Trim(mInString.Substring(1433, 1)) & "',") 'Special Proper Shipping Name Flg US
        mStrSQL.Append("'" & Trim(mInString.Substring(1433, 1)) & "',") 'Intermodal Flg CN
        mStrSQL.Append("'" & Trim(mInString.Substring(1434, 1)) & "',") 'Intermodal Flg Intl
        mStrSQL.Append("'" & Trim(mInString.Substring(1435, 1)) & "',") 'Intermodal Flg US
        '1436-1437 are reserved/blank
        mStrSQL.Append("'" & Trim(mInString.Substring(1438, 2)) & "',") 'RSSM Flg US
        mThisStr = Trim(RemoveSpaces(mInString.Substring(1690, 250)))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & mThisStr & "',") 'STCC Description
        mStrSQL.Append("'" & Trim(mInString.Substring(1940, 2)) & "',") 'Sub Placard Intl
        mStrSQL.Append("'" & Trim(mInString.Substring(1942, 9)) & "',") 'Sub Hazard Class Intl
        mStrSQL.Append("'" & mInString.Substring(1955, 2) & "/" &
                                  mInString.Substring(1957, 2) & "/" &
                                  mInString.Substring(1951, 4) & "',") 'Deletion Date
        mStrSQL.Append("'" & RemoveSpaces(mInString.Substring(1959, 625)) & "',") 'Alt Shipping Name CN
        mStrSQL.Append("'" & RemoveSpaces(mInString.Substring(2584, 625)) & "',") 'Alt Proper Shipping Name Intl
        mStrSQL.Append("'" & RemoveSpaces(mInString.Substring(3209, 625)) & "'") 'Alt Proper Shipping Name US
        '3834 to End of record are reserved/blank
        mStrSQL.Append(")")

        Build_Insert_HazMat_STCC_Rec = mStrSQL

    End Function

    Function Build_Insert_Non_HazMat_STCC_Rec(ByVal mInString As String, ByVal mTableName As String) As StringBuilder
        Dim mStrSQL As StringBuilder
        Dim mThisStr As String

        mStrSQL = New StringBuilder
        mStrSQL.Append("INSERT INTO " & mTableName & " (")
        mStrSQL.Append("STCC,")
        mStrSQL.Append("STCC_Code2,")
        mStrSQL.Append("Transaction_Date,")
        mStrSQL.Append("Transaction_Time,")
        mStrSQL.Append("Status_Code,")
        mStrSQL.Append("Int_Harm_Code,")
        mStrSQL.Append("Indus_Class,")
        mStrSQL.Append("Inter_SIC,")
        mStrSQL.Append("Dom_Canada,")
        mStrSQL.Append("CS_54_Code,")
        mStrSQL.Append("CS_54_Name,")
        mStrSQL.Append("Dereg_Flg,")
        mStrSQL.Append("Dereg_Date,")
        mStrSQL.Append("Car_Grade,")
        mStrSQL.Append("Effective_Date,")
        mStrSQL.Append("Header1,")
        mStrSQL.Append("Header2,")
        mStrSQL.Append("Header3,")
        mStrSQL.Append("Header4,")
        mStrSQL.Append("Alpha_Desc,")
        mStrSQL.Append("STCC_Desc_15,")
        mStrSQL.Append("Alternate_Num,")
        mStrSQL.Append("STCC_Repl_Code,")
        mStrSQL.Append("Expiration_Date,")
        mStrSQL.Append("Deletion_Date")
        mStrSQL.Append(") VALUES (")
        mStrSQL.Append("'" & Trim(mInString.Substring(0, 7)) & "',") 'STCC Code
        mStrSQL.Append("'" & Trim(mInString.Substring(7, 7)) & "',") 'STCC Code2
        If IsDate(mInString.Substring(18, 2) & "/" &
            mInString.Substring(20, 2) & "/" &
                mInString.Substring(14, 4)) Then
            mStrSQL.Append("'" & mInString.Substring(18, 2) & "/" &
            mInString.Substring(20, 2) & "/" &
            mInString.Substring(14, 4) & "',") 'Transaction Date
        Else
            mStrSQL.Append("'',")
        End If
        If IsDate(mInString.Substring(22, 2) & ":" & mInString.Substring(24, 2) & ":" & mInString.Substring(26, 2)) Then
            mStrSQL.Append("'" & mInString.Substring(22, 2) & ":" &
            mInString.Substring(24, 2) & ":" &
            mInString.Substring(26, 2) & "',") 'Transaction Time
        Else
            mStrSQL.Append("'',")
        End If
        mStrSQL.Append("'" & mInString.Substring(28, 1) & "',") 'Status_Code
        mStrSQL.Append("'" & Trim(mInString.Substring(29, 12)) & "',") 'Int_Harm_Code
        mStrSQL.Append("'" & Trim(mInString.Substring(149, 4)) & "',") 'Indus_Class
        mStrSQL.Append("'" & Trim(mInString.Substring(181, 4)) & "',") 'Inter_SIC
        mStrSQL.Append("'" & Trim(mInString.Substring(213, 5)) & "',") 'Dom_Canada
        mStrSQL.Append("'" & Trim(mInString.Substring(218, 2)) & "',") 'CS_54_Code
        mStrSQL.Append("'" & Trim(mInString.Substring(220, 30)) & "',") 'CS_54_Name
        mStrSQL.Append("'" & Trim(mInString.Substring(250, 1)) & "',") 'Dereg_Flg
        If IsDate(mInString.Substring(255, 2) & "/" &
            mInString.Substring(257, 2) & "/" &
                mInString.Substring(251, 4)) Then
            mStrSQL.Append("'" & mInString.Substring(255, 2) & "/" &
            mInString.Substring(257, 2) & "/" &
            mInString.Substring(251, 4) & "',") 'Dereg_Date
        Else
            mStrSQL.Append("'',")
        End If
        mStrSQL.Append("'" & Trim(mInString.Substring(259, 1)) & "',") 'Car_Grade
        If IsDate(mInString.Substring(260, 2) & "/" &
            mInString.Substring(262, 2) & "/" &
                mInString.Substring(264, 4)) Then
            mStrSQL.Append("'" & mInString.Substring(260, 2) & "/" &
                mInString.Substring(262, 2) & "/" &
                mInString.Substring(264, 4) & "',") 'Effective Date
        Else
            mStrSQL.Append("'',")
        End If
        mStrSQL.Append("'" & Trim(mInString.Substring(268, 2)) & "',") 'Header1
        mStrSQL.Append("'" & Trim(mInString.Substring(270, 3)) & "',") 'Header2
        mStrSQL.Append("'" & Trim(mInString.Substring(273, 4)) & "',") 'Header3
        mStrSQL.Append("'" & Trim(mInString.Substring(277, 5)) & "',") 'Header4
        mThisStr = Trim(RemoveSpaces(mInString.Substring(282, 2500)))
        mThisStr = Replace(mThisStr, "- ", "")
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & Left(mThisStr, 255) & "',") 'Alpha Description - first 255 chars
        mThisStr = Trim(mInString.Substring(2782, 15))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("'" & mThisStr & "',") 'STCC Description 15
        mStrSQL.Append("'" & Trim(mInString.Substring(2797, 2)) & "',") 'Alt Number
        mStrSQL.Append("'" & Trim(mInString.Substring(2799, 7)) & "',") 'STCC_Repl_Code
        If IsDate(mInString.Substring(2810, 2) & "/" &
            mInString.Substring(2812, 2) & "/" &
            mInString.Substring(2806, 4)) Then
            mStrSQL.Append("'" & mInString.Substring(2810, 2) & "/" &
                mInString.Substring(2812, 2) & "/" &
                mInString.Substring(2806, 4) & "',") 'Expiration Date
        Else
            mStrSQL.Append("'',")
        End If
        If IsDate(mInString.Substring(2818, 2) & "/" &
            mInString.Substring(2820, 2) & "/" &
            mInString.Substring(2814, 4)) Then
            mStrSQL.Append("'" & mInString.Substring(2818, 2) & "/" &
            mInString.Substring(2820, 2) & "/" &
            mInString.Substring(2814, 4) & "'") 'Deletion Date
        Else
            mStrSQL.Append("''")
        End If

        mStrSQL.Append(")")

        Build_Insert_Non_HazMat_STCC_Rec = mStrSQL

    End Function

    Function Build_Update_HazMat_STCC_Rec(ByVal mInString As String, ByVal mTableName As String) As StringBuilder
        Dim mStrSQL As StringBuilder
        Dim mThisStr As String

        mStrSQL = New StringBuilder
        mStrSQL.Append("UPDATE " & mTableName & " SET ")
        mStrSQL.Append("STCC_Code2 = '" & Trim(mInString.Substring(7, 7)) & "',")
        mStrSQL.Append("Transaction_Date = '" & mInString.Substring(18, 2) & "/" &
            mInString.Substring(20, 2) & "/" &
            mInString.Substring(14, 4) & "',")
        mStrSQL.Append("Transaction_Time = '" & mInString.Substring(22, 2) & ":" &
            mInString.Substring(24, 2) & ":" &
            mInString.Substring(26, 2) & "',")
        mStrSQL.Append("Status_Code = '" & Trim(mInString.Substring(28, 1)) & "',")
        mStrSQL.Append("Effective_Date = '" & mInString.Substring(33, 2) & "/" &
                                  mInString.Substring(35, 2) & "/" &
                                  mInString.Substring(29, 4) & "',")
        mStrSQL.Append("Header1 = '" & Trim(mInString.Substring(37, 2)) & "',")
        mStrSQL.Append("Header2 = '" & Trim(mInString.Substring(39, 3)) & "',")
        mStrSQL.Append("Header3 = '" & Trim(mInString.Substring(42, 4)) & "',")
        mStrSQL.Append("Header4 = '" & Trim(mInString.Substring(46, 5)) & "',")
        mStrSQL.Append("STCC_Desc_15 = '" & Trim(mInString.Substring(51, 15)) & "',")
        mStrSQL.Append("Alternate_Num = '" & Trim(mInString.Substring(66, 2)) & "',")
        mStrSQL.Append("Expiration_Date = '" & mInString.Substring(72, 2) & "/" &
                                  mInString.Substring(74, 2) & "/" &
                                  mInString.Substring(68, 4) & "',")
        mStrSQL.Append("Description_US = '" & Trim(mInString.Substring(76, 250)) & "',")
        mStrSQL.Append("Primary_Hazard_Class_Intl = '" & Trim(mInString.Substring(326, 4)) & "',")
        mStrSQL.Append("NOS_Flg_Intl = '" & Trim(mInString.Substring(330, 1)) & "',")
        mStrSQL.Append("Technical_Name_Intl = '" & Trim(mInString.Substring(331, 125)) & "',")
        mStrSQL.Append("UN_Number_Intl = '" & Trim(mInString.Substring(456, 6)) & "',")
        mStrSQL.Append("Packing_Grp_Intl = '" & Trim(mInString.Substring(462, 1)) & "',")
        mStrSQL.Append("Poison_Flg_intl = '" & Trim(mInString.Substring(463, 1)) & "',")
        mStrSQL.Append("Primary_Placard_Intl = '" & Trim(mInString.Substring(464, 2)) & "',")
        mStrSQL.Append("Shipping_Name_Intl = '" & Trim(mInString.Substring(466, 125)) & "',")
        mStrSQL.Append("Primary_Class_CN = '" & Trim(mInString.Substring(591, 4)) & "',")
        mStrSQL.Append("Sub_Class_CN = '" & Trim(mInString.Substring(595, 9)) & "',")
        mStrSQL.Append("CN_Orig_US_Dest_Flg = '" & Trim(mInString.Substring(604, 1)) & "',")
        mStrSQL.Append("Emer_Resp_Asst_Plan_CN_Flg = '" & Trim(mInString.Substring(605, 4)) & "',")
        mStrSQL.Append("Primary_Placard_CN = '" & Trim(mInString.Substring(609, 2)) & "',")
        mStrSQL.Append("Special_Commodity_CN_Flg = '" & Trim(mInString.Substring(611, 1)) & "',")
        mStrSQL.Append("NOS_Flag_CN = '" & Trim(mInString.Substring(613, 1)) & "',")
        mStrSQL.Append("Sub_Placard_CN = '" & Trim(mInString.Substring(614, 2)) & "',")
        mStrSQL.Append("Technical_Name_CN = '" & Trim(mInString.Substring(616, 125)) & "',")
        mStrSQL.Append("UN_Number_CN = '" & Trim(mInString.Substring(741, 6)) & "',")
        mStrSQL.Append("Packing_Grp_CN = '" & Trim(mInString.Substring(747, 1)) & "',")
        mStrSQL.Append("Poison_Flg_CN = '" & Trim(mInString.Substring(748, 1)) & "',")
        mStrSQL.Append("Shipping_Name_CN = '" & Trim(mInString.Substring(749, 125)) & "',")
        mStrSQL.Append("EPA_Waste_Stream_Num_US = '" & Trim(mInString.Substring(877, 18)) & "',")
        mStrSQL.Append("Haz_Placard_US = '" & Trim(mInString.Substring(895, 2)) & "',")
        mStrSQL.Append("Primary_Hazard_Class_US = '" & Trim(mInString.Substring(897, 4)) & "',")
        mStrSQL.Append("Sub_Hazard_US = '" & Trim(mInString.Substring(901, 6)) & "',")
        mStrSQL.Append("Hazard_Zone_US = '" & Trim(mInString.Substring(907, 1)) & "',")
        mStrSQL.Append("NOS_Flag_US = '" & Trim(mInString.Substring(908, 1)) & "',")
        mStrSQL.Append("Sub_Placard_US = '" & Trim(mInString.Substring(909, 2)) & "',")
        mStrSQL.Append("Technical_Name_US = '" & Trim(mInString.Substring(911, 125)) & "',")
        mStrSQL.Append("UN_NA_ID_Num_US = '" & Trim(mInString.Substring(1036, 6)) & "',")
        mStrSQL.Append("US_Orig_CN_Dest_Flg = '" & Trim(mInString.Substring(1042, 1)) & "',")
        mStrSQL.Append("Packing_Grp_US = '" & Trim(mInString.Substring(1043, 1)) & "',")
        mStrSQL.Append("Poison_Flg_US = '" & Trim(mInString.Substring(1044, 1)) & "',")
        mStrSQL.Append("Primary_Placard_US = '" & Trim(mInString.Substring(1045, 2)) & "',")
        mThisStr = RemoveSpaces(mInString.Substring(1047, 125))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("Shipping_Name_US = '" & Left(mThisStr, 255) & "',")
        mStrSQL.Append("OT55_Flg = '" & Trim(mInString.Substring(1172, 1)) & "',")
        mStrSQL.Append("Reportable_Quantity_Flg_US = '" & Trim(mInString.Substring(1174, 1)) & "',")
        mStrSQL.Append("Marine_Pollutant_Flg_US = '" & Trim(mInString.Substring(1175, 1)) & "',")
        mStrSQL.Append("HazMat_Name_US = '" & RemoveSpaces(mInString.Substring(1176, 125)) & "',")
        mStrSQL.Append("Marine_Pollutant_Name_US = '" & RemoveSpaces(mInString.Substring(1301, 125)) & "',")
        mStrSQL.Append("Special_Shipping_Name_Flg_CN = '" & Trim(mInString.Substring(1430, 1)) & "',")
        mStrSQL.Append("Spcl_Proper_Ship_Name_Flg_Intl = '" & Trim(mInString.Substring(1431, 1)) & "',")
        mStrSQL.Append("Spcl_Proper_Ship_Name_Flg_US = '" & Trim(mInString.Substring(1433, 1)) & "',")
        mStrSQL.Append("Intermodal_Flg_CN = '" & Trim(mInString.Substring(1433, 1)) & "',")
        mStrSQL.Append("Internodal_Flg_Intl = '" & Trim(mInString.Substring(1434, 1)) & "',")
        mStrSQL.Append("Intermodal_Flg_US = '" & Trim(mInString.Substring(1435, 1)) & "',")
        mStrSQL.Append("RSSM_Flg_US = '" & Trim(mInString.Substring(1438, 2)) & "',")
        mStrSQL.Append("Alpha_Desc = '" & RemoveSpaces(mInString.Substring(1440, 250)) & "',")
        mStrSQL.Append("STCC_Desc = '" & RemoveSpaces(mInString.Substring(1690, 250)) & "',")
        mStrSQL.Append("Sub_Placard_Intl = '" & Trim(mInString.Substring(1940, 2)) & "',")
        mStrSQL.Append("HazMat_Class_Intl = '" & Trim(mInString.Substring(1942, 9)) & "',")
        mStrSQL.Append("Deletion_Date = '" & mInString.Substring(1955, 2) & "/" &
            mInString.Substring(1957, 2) & "/" &
            mInString.Substring(1951, 4) & "',")
        mStrSQL.Append("Alt_Shipping_Name_CN = '" & RemoveSpaces(mInString.Substring(1959, 625)) & "',")
        mStrSQL.Append("Alt_Proper_Ship_Name_Intl = '" & RemoveSpaces(mInString.Substring(2584, 625)) & "',")
        mStrSQL.Append("Alt_Proper_Ship_Name_US = '" & RemoveSpaces(mInString.Substring(3209, 625)) & "'")
        mStrSQL.Append(" WHERE STCC = '" & Trim(mInString.Substring(0, 7)) & "'")

        mStrSQL.Append("'" & RemoveSpaces(mInString.Substring(3209, 625)) & "'") 'Alt Proper Shipping Name US

        Build_Update_HazMat_STCC_Rec = mStrSQL

    End Function

    Function Build_Update_Non_HazMat_STCC_Rec(ByVal mInString As String, ByVal mTableName As String) As StringBuilder
        Dim mStrSQL As StringBuilder
        Dim mThisStr As String

        mStrSQL = New StringBuilder
        mStrSQL.Append("UPDATE " & mTableName & " SET ")
        mStrSQL.Append("STCC_Code2 = '" & Trim(mInString.Substring(7, 7)) & "',")
        If IsDate(mInString.Substring(18, 2) & "/" &
            mInString.Substring(20, 2) & "/" &
                mInString.Substring(14, 4)) Then
            mStrSQL.Append("Transaction_Date = '" & mInString.Substring(18, 2) & "/" &
            mInString.Substring(20, 2) & "/" &
            mInString.Substring(14, 4) & "',")
        Else
            mStrSQL.Append("Transaction_Date = '',")
        End If
        mStrSQL.Append("Transaction_Time = '" & mInString.Substring(22, 2) & ":" &
            mInString.Substring(24, 2) & ":" &
            mInString.Substring(26, 2) & "',") '
        mStrSQL.Append("Status_Code = '" & mInString.Substring(28, 1) & "',")
        mStrSQL.Append("Int_Harm_Code = '" & Trim(mInString.Substring(29, 12)) & "',")
        mStrSQL.Append("Indus_Class = '" & Trim(mInString.Substring(149, 4)) & "',")
        mStrSQL.Append("Inter_SIC = '" & Trim(mInString.Substring(181, 4)) & "',")
        mStrSQL.Append("Dom_Canada = '" & Trim(mInString.Substring(213, 5)) & "',")
        mStrSQL.Append("CS_54_Code = '" & Trim(mInString.Substring(218, 2)) & "',")
        mStrSQL.Append("CS_54_Name = '" & Trim(mInString.Substring(220, 30)) & "',")
        mStrSQL.Append("Dereg_Flg = '" & Trim(mInString.Substring(250, 1)) & "',")
        If IsDate(mInString.Substring(255, 2) & "/" &
            mInString.Substring(257, 2) & "/" &
                mInString.Substring(251, 4)) Then
            mStrSQL.Append("Dereg_Date = '" & mInString.Substring(255, 2) & "/" &
            mInString.Substring(257, 2) & "/" &
            mInString.Substring(251, 4) & "',")
        Else
            mStrSQL.Append("Dereg_Date = '',")
        End If
        mStrSQL.Append("Car_Grade = '" & Trim(mInString.Substring(259, 1)) & "',") '
        If IsDate(mInString.Substring(260, 2) & "/" &
            mInString.Substring(262, 2) & "/" &
                mInString.Substring(264, 4)) Then
            mStrSQL.Append("Effective_Date = '" & mInString.Substring(260, 2) & "/" &
                mInString.Substring(262, 2) & "/" &
                mInString.Substring(264, 4) & "',")
        Else
            mStrSQL.Append("Effective Date = '',")
        End If
        mStrSQL.Append("Header1 = '" & Trim(mInString.Substring(268, 2)) & "',")
        mStrSQL.Append("Header2 = '" & Trim(mInString.Substring(270, 3)) & "',")
        mStrSQL.Append("Header3 = '" & Trim(mInString.Substring(273, 4)) & "',")
        mStrSQL.Append("Header4 = '" & Trim(mInString.Substring(277, 5)) & "',")
        mThisStr = Trim(RemoveSpaces(mInString.Substring(282, 2500)))
        mThisStr = Replace(mThisStr, "- ", "")
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("Alpha_Desc = '" & Left(mThisStr, 255) & "',")
        mThisStr = Trim(mInString.Substring(2782, 15))
        mThisStr = Replace(mThisStr, "'", "")
        mStrSQL.Append("STCC_Desc_15 = '" & mThisStr & "',")
        mStrSQL.Append("Alternate_Num = '" & Trim(mInString.Substring(2797, 2)) & "',")
        mStrSQL.Append("STCC_Repl_Code = '" & Trim(mInString.Substring(2799, 7)) & "',")
        If IsDate(mInString.Substring(2810, 2) & "/" &
            mInString.Substring(2812, 2) & "/" &
            mInString.Substring(2806, 4)) Then
            mStrSQL.Append("Expiration_Date = '" & mInString.Substring(2810, 2) & "/" &
                mInString.Substring(2812, 2) & "/" &
                mInString.Substring(2806, 4) & "',")
        Else
            mStrSQL.Append("Expiration_Date = '',")
        End If
        If IsDate(mInString.Substring(2818, 2) & "/" &
            mInString.Substring(2820, 2) & "/" &
            mInString.Substring(2814, 4)) Then
            mStrSQL.Append("Deletion_Date = '" & mInString.Substring(2818, 2) & "/" &
            mInString.Substring(2820, 2) & "/" &
            mInString.Substring(2814, 4) & "'")
        Else
            mStrSQL.Append("Deletion_Date = ''")
        End If
        mStrSQL.Append(" WHERE STCC = '" & Trim(mInString.Substring(0, 7)) & "'")

        Build_Update_Non_HazMat_STCC_Rec = mStrSQL

    End Function

    Function Build_Insert_CSM_Record(ByVal mInstring As String, ByVal mTablename As String) As String
        Dim mStrSQL As StringBuilder

        mStrSQL = New StringBuilder
        mStrSQL.Append("INSERT INTO " & mTablename & " (")
        mStrSQL.Append("Road_Mark,")
        mStrSQL.Append("FSAC,")
        mStrSQL.Append("Eff_Date,")
        mStrSQL.Append("Eff_Time,")
        mStrSQL.Append("Sta_Status,")
        mStrSQL.Append("R260,")
        mStrSQL.Append("SPLC,")
        mStrSQL.Append("OPSL,")
        mStrSQL.Append("Loc_Type,")
        mStrSQL.Append("AAR_Status,")
        mStrSQL.Append("Loc_Name,")
        mStrSQL.Append("OPSL_Name,")
        mStrSQL.Append("Loc_Geo_Name,")
        mStrSQL.Append("Loc_County,")
        mStrSQL.Append("Loc_St,")
        mStrSQL.Append("Loc_Cntry,")
        mStrSQL.Append("Loc_Zip,")
        'mStrSQL.Append("Loc_Rating_Zip,")
        mStrSQL.Append("Rate_Base_SPLC,")
        mStrSQL.Append("Rate_Base_City,")
        mStrSQL.Append("Rate_Base_St,")
        mStrSQL.Append("Rev_Sw_SPLC,")
        mStrSQL.Append("Rev_Sw_City,")
        mStrSQL.Append("Rev_Sw_St,")
        mStrSQL.Append("CIF_ID,")
        mStrSQL.Append("Imp_Exp_Flg,")
        mStrSQL.Append("Customs_Flg,")
        mStrSQL.Append("Grain_Flg,")
        mStrSQL.Append("Auto_Ramp_Flg,")
        mStrSQL.Append("Intermodal_Flg,")
        mStrSQL.Append("Embargo_Flg,")
        mStrSQL.Append("Oper_Plate,")
        mStrSQL.Append("Oper_Wght,")
        mStrSQL.Append("FIPS,")
        mStrSQL.Append("BEA,")
        mStrSQL.Append("BEA_Name,")
        mStrSQL.Append("CEA,")
        mStrSQL.Append("Lat,")
        mStrSQL.Append("Long,")
        mStrSQL.Append("Reload,")
        mStrSQL.Append("Geo_SPLC,")
        mStrSQL.Append("Customs_CIF,")
        mStrSQL.Append("Time_Zone,")
        mStrSQL.Append("Daylight_Ind,")
        mStrSQL.Append("OPSL_Notes,")
        mStrSQL.Append("OPSL_Ref,")
        mStrSQL.Append("Exp_Date,")
        mStrSQL.Append("Interswitch_Area,")
        mStrSQL.Append("Name_333,")
        mStrSQL.Append("New_Road_Name,")
        mStrSQL.Append("New_FSAC,")
        mStrSQL.Append("Last_Maint_Stamp,")
        mStrSQL.Append("Last_Trans_Type,")
        mStrSQL.Append("Last_Update_Type,")
        mStrSQL.Append("Last_Road_Mark,")
        mStrSQL.Append("Last_Date,")
        mStrSQL.Append("Last_Time")

        mStrSQL.Append(") VALUES (")

        mStrSQL.Append("'" & Trim(mInstring.Substring(0, 4)) & "',") 'Road_Mark
        mStrSQL.Append("'" & Trim(mInstring.Substring(4, 5)) & "',") 'FSAC
        If IsDate(mInstring.Substring(15, 2) & "/" & mInstring.Substring(13, 2) & "/" & mInstring.Substring(9, 4)) Then
            mStrSQL.Append("'" & mInstring.Substring(15, 2) & "/" &
            mInstring.Substring(13, 2) & "/" &
            mInstring.Substring(9, 4) & "',") 'Eff_Date
        Else
            mStrSQL.Append("'',")
        End If
        If IsDate(mInstring.Substring(17, 2) & ":" & mInstring.Substring(19, 2) & ":" & mInstring.Substring(21, 2)) Then
            mStrSQL.Append("'" & mInstring.Substring(17, 2) & ":" &
            mInstring.Substring(19, 2) & ":" &
            mInstring.Substring(21, 2) & "',") 'Eff_Time
        Else
            mStrSQL.Append("'',")
        End If
        mStrSQL.Append("'" & mInstring.Substring(23, 2) & "',") 'Sta_Status
        mStrSQL.Append("'" & Trim(mInstring.Substring(26, 5)) & "',") 'R260
        mStrSQL.Append("'" & Trim(mInstring.Substring(31, 9)) & "',") 'SPLC
        mStrSQL.Append("'" & Trim(mInstring.Substring(40, 7)) & "',") 'OPSL
        mStrSQL.Append("'" & Trim(mInstring.Substring(47, 5)) & "',") 'Loc_Type
        mStrSQL.Append("'" & Trim(mInstring.Substring(52, 1)) & "',") 'AAR_Status
        mStrSQL.Append("'" & Trim(mInstring.Substring(54, 30)).Replace("'", "''") & "',") 'Loc_Name
        mStrSQL.Append("'" & Trim(mInstring.Substring(84, 30)).Replace("'", "''") & "',") 'OPSL_Name
        mStrSQL.Append("'" & Trim(mInstring.Substring(114, 30)).Replace("'", "''") & "',") 'Loc_Geo_Name
        mStrSQL.Append("'" & Trim(mInstring.Substring(144, 30)).Replace("'", "''") & "',") 'Loc_County
        mStrSQL.Append("'" & Trim(mInstring.Substring(174, 2)) & "',") 'Loc_St
        mStrSQL.Append("'" & Trim(mInstring.Substring(176, 2)) & "',") 'Loc_Cntry
        mStrSQL.Append("'" & Trim(mInstring.Substring(178, 11)) & "',") 'Loc_Zip
        ' mStrSQL.Append("'" & Trim(mInstring.Substring(189, 11)) & "',") 'Loc_Rating_Zip
        mStrSQL.Append("'" & Trim(mInstring.Substring(200, 9)) & "',") 'Rate_Base_SPLC
        mStrSQL.Append("'" & Trim(mInstring.Substring(209, 30)).Replace("'", "''") & "',") 'Rate_Base_City
        mStrSQL.Append("'" & Trim(mInstring.Substring(239, 2)) & "',") 'Rate_Base_St
        mStrSQL.Append("'" & Trim(mInstring.Substring(241, 9)) & "',") 'Rev_Sw_SPLC
        mStrSQL.Append("'" & Trim(mInstring.Substring(250, 30)).Replace("'", "''") & "',") 'Rev_Sw_City
        mStrSQL.Append("'" & Trim(mInstring.Substring(280, 2)) & "',") 'Rev_Sw_St
        mStrSQL.Append("'" & Trim(mInstring.Substring(282, 2)) & "',") 'CIF_ID
        mStrSQL.Append("'" & Trim(mInstring.Substring(292, 1)) & "',") 'Imp_Exp_Flg
        mStrSQL.Append("'" & Trim(mInstring.Substring(293, 1)) & "',") 'Customs_Flg
        mStrSQL.Append("'" & Trim(mInstring.Substring(294, 1)) & "',") 'Grain_Flg
        mStrSQL.Append("'" & Trim(mInstring.Substring(295, 1)) & "',") 'Auto_Ramp_Flg
        mStrSQL.Append("'" & Trim(mInstring.Substring(296, 5)) & "',") 'Intermodal_Flg
        mStrSQL.Append("'" & Trim(mInstring.Substring(301, 1)) & "',") 'Embargo_Flg
        mStrSQL.Append("'" & Trim(mInstring.Substring(302, 1)) & "',") 'Oper_Plate
        mStrSQL.Append("'" & Trim(mInstring.Substring(303, 4)) & "',") 'Oper_Wght
        mStrSQL.Append("'" & Trim(mInstring.Substring(395, 5)) & "',") 'FIPS
        mStrSQL.Append("'" & Trim(mInstring.Substring(400, 3)) & "',") 'BEA
        mStrSQL.Append("'" & Trim(mInstring.Substring(403, 60)).Replace("'", "''") & "',") 'BEA_Name
        mStrSQL.Append("'" & Trim(mInstring.Substring(463, 4)) & "',") 'CEA
        mStrSQL.Append("'" & Trim(mInstring.Substring(470, 9)) & "',") 'Lat
        mStrSQL.Append("'" & Trim(mInstring.Substring(479, 18)) & "',") 'Long
        mStrSQL.Append("'" & Trim(mInstring.Substring(488, 5)) & "',") 'Reload
        mStrSQL.Append("'" & Trim(mInstring.Substring(493, 9)) & "',") 'Geo_SPLC
        mStrSQL.Append("'" & Trim(mInstring.Substring(502, 13)) & "',") 'Customs_CIF
        mStrSQL.Append("'" & Trim(mInstring.Substring(515, 2)) & "',") 'Time_Zone
        mStrSQL.Append("'" & Trim(mInstring.Substring(517, 1)) & "',") 'Daylight_Ind
        mStrSQL.Append("'" & Trim(mInstring.Substring(518, 40)) & "',") 'OPSL_Notes
        mStrSQL.Append("'" & Trim(mInstring.Substring(558, 3)) & "',") 'OPSL_Ref
        If IsDate(mInstring.Substring(677, 2) & "/" & mInstring.Substring(679, 2) & "/" & mInstring.Substring(673, 4)) Then
            mStrSQL.Append("'" & mInstring.Substring(677, 2) & "/" &
            mInstring.Substring(679, 2) & "/" &
            mInstring.Substring(673, 4) & "',") 'Exp_Date
        Else
            mStrSQL.Append("'',")
        End If
        mStrSQL.Append("'" & Trim(mInstring.Substring(681, 9)) & "',") 'Interswitch_Area
        mStrSQL.Append("'" & Trim(mInstring.Substring(690, 9)).Replace("'", "''") & "',") 'Name_333
        mStrSQL.Append("'" & Trim(mInstring.Substring(699, 4)) & "',") 'New_Road_Name
        mStrSQL.Append("'" & Trim(mInstring.Substring(703, 5)) & "',") 'New_FSAC
        mStrSQL.Append("'" & Trim(mInstring.Substring(754, 26)) & "',") 'Last_Maint_Stamp
        mStrSQL.Append("'" & Trim(mInstring.Substring(780, 1)) & "',") 'Last_Trans_Type
        mStrSQL.Append("'',") ' Last_Update_Type
        mStrSQL.Append("'" & Trim(mInstring.Substring(782, 4)) & "',") 'Last_Road_Mark
        If IsDate(mInstring.Substring(790, 2) & "/" & mInstring.Substring(792, 2) & "/" & mInstring.Substring(786, 4)) Then
            mStrSQL.Append("'" & mInstring.Substring(790, 2) & "/" &
            mInstring.Substring(792, 2) & "/" &
            mInstring.Substring(786, 4) & "',") 'Last_Date
        Else
            mStrSQL.Append("'',")
        End If
        If IsDate(mInstring.Substring(794, 2) & ":" & mInstring.Substring(796, 2) & ":" & mInstring.Substring(698, 2)) Then
            mStrSQL.Append("'" & mInstring.Substring(794, 2) & ":" &
            mInstring.Substring(796, 2) & ":" &
            mInstring.Substring(798, 2) & "'") 'Last_Time
        Else
            mStrSQL.Append("''")
        End If

        mStrSQL.Append(")")

        Build_Insert_CSM_Record = mStrSQL.ToString

    End Function

    Public Sub Insert_AuditTrail_Record(mDatabase As String, mActivity As String)
        Dim mSQLCommand As SqlCommand

        Gbl_AuditTrailLog_Tablename = "ActivityAuditLog"

        ' Verify table exists

        mSQLCommand = New SqlCommand
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = "INSERT INTO " & Gbl_AuditTrailLog_Tablename & " (" &
            "ExecutedDateTime, ActivityExecuted, ExecutedBy " &
            ") VALUES (" &
            "CAST('" & Now().ToString("MM/dd/yyyy HH:mm:ss.fffffff") & "' AS DATETIME2(7)), " &
            "'" & mActivity & "'," &
            "'" & My.User.Name & "')"

        ' Open the SQL connection
        OpenSQLConnection(mDatabase)

        ' execute the command
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.ExecuteNonQuery()

    End Sub

    Public Sub CreateProcedure(mProcedureYear As String, mProcedureName As String)
        Dim mCommand As New SqlCommand

        ' Open the SQL connection
        OpenSQLConnection("URCS" & mProcedureYear)

        mCommand.Connection = gbl_SQLConnection
        mCommand.CommandType = CommandType.Text

        If ProcedureExist(mProcedureYear, mProcedureName) = True Then
            mCommand.CommandText = "DROP PROCEDURE " & mProcedureName
            mCommand.ExecuteNonQuery()
            Insert_AuditTrail_Record("URCS" & mProcedureYear, "Dropped " & "URCS" & mProcedureYear & "->" & mProcedureName & " Procedure.")
        End If

        Select Case mProcedureName
            Case "usp_GenerateEValuesXML"
                mCommand.CommandText = Create_GenerateEValues_Procedure_Command()
            Case "usp_RunSubstitutions"
                mCommand.CommandText = Create_RunSubstitutions_Procedure_Command()
            Case "usp_WriteCRPSEG_Legacy"
                mCommand.CommandText = Create_Write_CRPSEG_Legacy_Procedure_Command()
        End Select

        mCommand.ExecuteNonQuery()
        Insert_AuditTrail_Record("URCS" & mProcedureYear, "Created " & "URCS" & mProcedureYear & "->" & mProcedureName & " Procedure.")

    End Sub

    Public Sub Create_ufn_EValues_Function(mFunctionYear As String)

        Dim mCommand As New SqlCommand

        ' Open the SQL connection
        OpenSQLConnection("URCS" & mFunctionYear)

        mCommand.Connection = gbl_SQLConnection
        mCommand.CommandType = CommandType.Text

        If FunctionExist(mFunctionYear) = True Then
            mCommand.CommandText = "DROP FUNCTION ufn_EValues"
            mCommand.ExecuteNonQuery()
            Insert_AuditTrail_Record("URCS" & mFunctionYear, "Dropped URCS" & mFunctionYear & "->ufn_EValues Function.")
        End If

        mCommand.CommandText = Create_EValues_Function_Command()
        mCommand.ExecuteNonQuery()

        Insert_AuditTrail_Record("URCS" & mFunctionYear, "Created URCS" & mFunctionYear & "->ufn_EValues Function.")

    End Sub
    Public Sub Insert_Productivity_Record(ByVal mYear As Integer,
                                              ByVal mLengthOfHaulStratum As Integer,
                                              ByVal mCarTypeStratum As Integer,
                                              ByVal mLadingWeightStratum As Integer,
                                              ByVal mCarsStratum As Integer,
                                              ByVal mRevenue As Long,
                                              ByVal mTonMiles As Long)
        Dim mSQLCommand As SqlCommand
        Dim mStrSQL As StringBuilder

        mStrSQL = New StringBuilder
        mStrSQL.Append("INSERT INTO " & Gbl_Productivity_TableName & " (")
        mStrSQL.Append("Year,")
        mStrSQL.Append("Length_Of_Haul_Stratum,")
        mStrSQL.Append("Car_Type_Stratum,")
        mStrSQL.Append("Lading_Weight_Stratum,")
        mStrSQL.Append("Cars_Stratum,")
        mStrSQL.Append("Revenue,")
        mStrSQL.Append("Ton_Miles,")
        mStrSQL.Append("Waybill_Records")

        mStrSQL.Append(") VALUES (")

        mStrSQL.Append(mYear.ToString & ",")
        mStrSQL.Append(mLengthOfHaulStratum.ToString & ",")
        mStrSQL.Append(mCarTypeStratum.ToString & ",")
        mStrSQL.Append(mLadingWeightStratum.ToString & ",")
        mStrSQL.Append(mCarsStratum.ToString & ",")
        mStrSQL.Append(mRevenue.ToString & ",")
        mStrSQL.Append(mTonMiles.ToString & ",")
        mStrSQL.Append("1)") ' If we're inserting, this will always be the first record

        ' Open the SQL connection
        OpenSQLConnection(Gbl_Waybill_Database_Name)

        ' execute the command
        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandText = mStrSQL.ToString
        mSQLCommand.ExecuteNonQuery()

    End Sub

    Public Sub Update_Productivity_Record(ByVal mYear As Integer,
                                              ByVal mLengthOfHaulStratum As Integer,
                                              ByVal mCarTypeStratum As Integer,
                                              ByVal mLadingWeightStratum As Integer,
                                              ByVal mCarsStratum As Integer,
                                              ByVal mRevenue As Long,
                                              ByVal mTonMiles As Long,
                                              ByVal mNumberOfWaybills As Integer)
        Dim mSQLCommand As SqlCommand
        Dim mStrSQL As StringBuilder

        mStrSQL = New StringBuilder
        mStrSQL.Append("Update " & Gbl_Productivity_TableName & " SET " &
            "Revenue = " & mRevenue.ToString & ", " &
            "Ton_Miles = " & mTonMiles.ToString & ", " &
            "Waybill_Records = " & mNumberOfWaybills.ToString &
            "WHERE Year = " & mYear.ToString & " AND " &
            "Length_Of_Haul_Stratum = " & mLengthOfHaulStratum.ToString & " AND " &
            "Car_Type_Stratum = " & mCarTypeStratum.ToString & " AND " &
            "Lading_Weight_Stratum = " & mLadingWeightStratum.ToString & " AND " &
            "Cars_Stratum = " & mCarsStratum.ToString)

        ' Open the SQL connection
        OpenSQLConnection(Gbl_Waybill_Database_Name)

        ' execute the command
        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandText = mStrSQL.ToString
        mSQLCommand.ExecuteNonQuery()

    End Sub

    Function Build_Insert_Marks_Record(ByVal mInstring As String, ByVal mTablename As String) As String
        Dim mStrSQL As StringBuilder

        mStrSQL = New StringBuilder
        mStrSQL.Append("INSERT INTO " & mTablename & " (")
        mStrSQL.Append("Road_Mark,")
        mStrSQL.Append("Owner,")
        mStrSQL.Append("Release_Ind,")
        mStrSQL.Append("Placement_Ind,")
        mStrSQL.Append("Mark_Type,")
        mStrSQL.Append("Exp_Date,")
        mStrSQL.Append("Mark_Name,")
        mStrSQL.Append("R260_Num,")
        mStrSQL.Append("Eff_Date,")
        mStrSQL.Append("CIF_ID,")
        mStrSQL.Append("Last_Maint_Stamp")

        mStrSQL.Append(") VALUES (")

        mStrSQL.Append("'" & Trim(mInstring.Substring(0, 4)) & "',") 'Road_Mark
        mStrSQL.Append("'" & Trim(mInstring.Substring(4, 4)) & "',") 'Owner
        mStrSQL.Append("'" & Trim(mInstring.Substring(10, 1)) & "',") 'Release_Ind
        mStrSQL.Append("'" & Trim(mInstring.Substring(11, 1)) & "',") 'Placement_Ind
        mStrSQL.Append("'" & Trim(mInstring.Substring(12, 3)) & "',") 'Mark_Type
        mStrSQL.Append("'" & mInstring.Substring(20, 10) & "',") 'Exp_Date
        mStrSQL.Append("'" & Trim(mInstring.Substring(30, 105)) & "',") 'Mark_Name
        mStrSQL.Append("'" & Trim(mInstring.Substring(136, 3)) & "',") 'R260_Num
        mStrSQL.Append("'" & mInstring.Substring(139, 10) & "',") 'Eff_Date
        mStrSQL.Append("'" & Trim(mInstring.Substring(149, 13)) & "',") 'CIF_ID
        mStrSQL.Append("'" & Trim(mInstring.Substring(162, 26)) & "'") 'Last_Maint_Stamp

        mStrSQL.Append(")")

        Build_Insert_Marks_Record = mStrSQL.ToString

    End Function

    Function Build_Update_CSM_Record(ByVal mInString As String, ByVal mTableName As String) As String
        Dim mStrSQL As StringBuilder

        mStrSQL = New StringBuilder
        mStrSQL.Append("UPDATE " & mTableName & " SET ")
        If IsDate(mInString.Substring(13, 2) & "/" & mInString.Substring(15, 2) & "/" & mInString.Substring(9, 4)) Then
            mStrSQL.Append("Eff_Date = '" & mInString.Substring(13, 2) & "/" &
            mInString.Substring(15, 2) & "/" &
            mInString.Substring(9, 4) & "',")
        Else
            mStrSQL.Append("Eff_Date ='',")
        End If
        If IsDate(mInString.Substring(17, 2) & ":" & mInString.Substring(19, 2) & ":" & mInString.Substring(21, 2)) Then
            mStrSQL.Append("Eff_Time = '" & mInString.Substring(17, 2) & ":" &
            mInString.Substring(19, 2) & ":" &
            mInString.Substring(21, 2) & "',")
        Else
            mStrSQL.Append("Eff_Time = '',")
        End If
        mStrSQL.Append("Sta_Status = '" & mInString.Substring(23, 2) & "',")
        mStrSQL.Append("R260 = '" & Trim(mInString.Substring(26, 5)) & "',")
        mStrSQL.Append("SPLC = '" & Trim(mInString.Substring(31, 9)) & "',")
        mStrSQL.Append("OPSL = '" & Trim(mInString.Substring(40, 7)) & "',")
        mStrSQL.Append("AAR_Status = '" & Trim(mInString.Substring(52, 1)) & "',")
        mStrSQL.Append("Loc_Name = '" & Trim(mInString.Substring(54, 30)).Replace("'", "''") & "',")
        mStrSQL.Append("OPSL_Name = '" & Trim(mInString.Substring(84, 30)).Replace("'", "''") & "',")
        mStrSQL.Append("Loc_Geo_Name ='" & Trim(mInString.Substring(114, 30)).Replace("'", "''") & "',")
        mStrSQL.Append("Loc_County = '" & Trim(mInString.Substring(144, 30)).Replace("'", "''") & "',")
        mStrSQL.Append("Loc_St = '" & Trim(mInString.Substring(174, 2)) & "',")
        mStrSQL.Append("Loc_Cntry = '" & Trim(mInString.Substring(176, 2)) & "',")
        mStrSQL.Append("Loc_Zip = '" & Trim(mInString.Substring(178, 11)) & "',")
        ' mStrSQL.Append("Loc_Rating_Zip = '" & Trim(mInString.Substring(189, 11)) & "',")
        mStrSQL.Append("Rate_Base_SPLC = '" & Trim(mInString.Substring(200, 9)) & "',")
        mStrSQL.Append("Rate_Base_City = '" & Trim(mInString.Substring(209, 30)).Replace("'", "''") & "',")
        mStrSQL.Append("Rate_Base_St = '" & Trim(mInString.Substring(239, 2)) & "',")
        mStrSQL.Append("Rev_Sw_SPLC = '" & Trim(mInString.Substring(241, 9)) & "',")
        mStrSQL.Append("Rev_Sw_City = '" & Trim(mInString.Substring(250, 30)).Replace("'", "''") & "',")
        mStrSQL.Append("Rev_Sw_St = '" & Trim(mInString.Substring(280, 2)) & "',")
        mStrSQL.Append("CIF_ID = '" & Trim(mInString.Substring(282, 2)) & "',")
        mStrSQL.Append("Imp_Exp_Flg = '" & Trim(mInString.Substring(292, 1)) & "',")
        mStrSQL.Append("Customs_Flg = '" & Trim(mInString.Substring(293, 1)) & "',")
        mStrSQL.Append("Grain_Flg = '" & Trim(mInString.Substring(294, 1)) & "',")
        mStrSQL.Append("Auto_Ramp_Flg = '" & Trim(mInString.Substring(295, 1)) & "',")
        mStrSQL.Append("Intermodal_Flg = '" & Trim(mInString.Substring(296, 5)) & "',")
        mStrSQL.Append("Embargo_Flg = '" & Trim(mInString.Substring(301, 1)) & "',")
        mStrSQL.Append("Oper_Plate = '" & Trim(mInString.Substring(302, 1)) & "',")
        mStrSQL.Append("Oper_Wght = '" & Trim(mInString.Substring(303, 4)) & "',")
        mStrSQL.Append("FIPS = '" & Trim(mInString.Substring(395, 5)) & "',")
        mStrSQL.Append("BEA = '" & Trim(mInString.Substring(400, 3)) & "',")
        mStrSQL.Append("BEA_Name = '" & Trim(mInString.Substring(403, 60)).Replace("'", "''") & "',")
        mStrSQL.Append("CEA = '" & Trim(mInString.Substring(463, 4)) & "',")
        mStrSQL.Append("Lat = '" & Trim(mInString.Substring(470, 9)) & "',")
        mStrSQL.Append("Long = '" & Trim(mInString.Substring(479, 9)) & "',")
        mStrSQL.Append("Reload = '" & Trim(mInString.Substring(488, 5)) & "',")
        mStrSQL.Append("Geo_SPLC = '" & Trim(mInString.Substring(493, 9)) & "',")
        mStrSQL.Append("Customs_CIF = '" & Trim(mInString.Substring(502, 13)) & "',")
        mStrSQL.Append("Time_Zone = '" & Trim(mInString.Substring(515, 2)) & "',")
        mStrSQL.Append("Daylight_Ind = '" & Trim(mInString.Substring(517, 1)) & "',")
        mStrSQL.Append("OPSL_Notes = '" & Trim(mInString.Substring(518, 40)) & "',")
        mStrSQL.Append("OPSL_Ref = '" & Trim(mInString.Substring(558, 3)) & "',")
        If IsDate(mInString.Substring(677, 2) & "/" & mInString.Substring(679, 2) & "/" & mInString.Substring(673, 4)) Then
            mStrSQL.Append("Exp_Date = '" & mInString.Substring(677, 2) & "/" &
            mInString.Substring(679, 2) & "/" &
            mInString.Substring(673, 4) & "',")
        Else
            mStrSQL.Append("Eff_Date ='',")
        End If
        mStrSQL.Append("Interswitch_Area = '" & Trim(mInString.Substring(681, 9)) & "',")
        mStrSQL.Append("Name_333 = '" & Trim(mInString.Substring(690, 9)).Replace("'", "''") & "',")
        mStrSQL.Append("New_Road_Name = '" & Trim(mInString.Substring(699, 4)) & "',")
        mStrSQL.Append("New_FSAC = '" & Trim(mInString.Substring(703, 5)) & "',")
        mStrSQL.Append("Last_Trans_Type = '" & Trim(mInString.Substring(780, 1)) & "',")
        mStrSQL.Append("Last_Update_Type = '" & Trim(mInString.Substring(781, 1)) & "',")
        mStrSQL.Append("Last_Road_Mark = '" & Trim(mInString.Substring(782, 4)) & "',")
        If IsDate(mInString.Substring(790, 2) & "/" & mInString.Substring(792, 2) & "/" & mInString.Substring(786, 4)) Then
            mStrSQL.Append("Last_Date = '" & mInString.Substring(790, 2) & "/" &
            mInString.Substring(792, 2) & "/" &
            mInString.Substring(786, 4) & "',")
        Else
            mStrSQL.Append("Last_Date ='',")
        End If
        If IsDate(mInString.Substring(794, 2) & ":" & mInString.Substring(796, 2) & ":" & mInString.Substring(798, 2)) Then
            mStrSQL.Append("Last_Time = '" & mInString.Substring(794, 2) & ":" &
            mInString.Substring(796, 2) & ":" &
            mInString.Substring(798, 2) & "'")
        Else
            mStrSQL.Append("Last_Time = ''")
        End If

        mStrSQL.Append(" WHERE Road_Mark = '" & Trim(mInString.Substring(0, 4)) & "' AND " &
            "FSAC = '" & Trim(mInString.Substring(4, 5)) & "' AND " &
            "Last_Maint_Stamp = '" & Trim(mInString.Substring(754, 26)) & "'")

        Build_Update_CSM_Record = mStrSQL.ToString

    End Function

    Function Build_Update_Marks_Record(ByVal mInString As String, ByVal mTableName As String) As String
        Dim mStrSQL As StringBuilder

        mStrSQL = New StringBuilder
        mStrSQL.Append("UPDATE " & mTableName & " SET ")
        mStrSQL.Append("Owner = '" & Trim(mInString.Substring(4, 4)) & "',")
        mStrSQL.Append("Release_Ind = '" & Trim(mInString.Substring(10, 1)) & "',")
        mStrSQL.Append("Placement_Ind = '" & Trim(mInString.Substring(11, 1)) & "',")
        mStrSQL.Append("Mark_Type = '" & Trim(mInString.Substring(12, 3)) & "',")
        mStrSQL.Append("Exp_Date = '" & mInString.Substring(20, 10) & "',")
        mStrSQL.Append("Mark_Name = '" & Trim(mInString.Substring(30, 105)) & "',")
        mStrSQL.Append("R260_Num = '" & Trim(mInString.Substring(136, 3)) & "',")
        mStrSQL.Append("Eff_Date = '" & mInString.Substring(139, 10) & "',")
        mStrSQL.Append("CIF_ID = '" & Trim(mInString.Substring(149, 13)) & "'")

        mStrSQL.Append(" WHERE Road_Mark = '" & Trim(mInString.Substring(0, 4)) & "' AND " &
            "Last_Maint_Stamp = '" & Trim(mInString.Substring(162, 26)) & "'")

        Build_Update_Marks_Record = mStrSQL.ToString

    End Function

    Function Count_Trans_Records(
        mYear As Integer,
        Optional ByVal mRRICC As Object = Nothing,
        Optional ByVal mSch As Object = Nothing,
        Optional ByVal mStartLine As Object = Nothing,
        Optional ByVal mEndLine As Object = Nothing) As Integer

        Dim mDataTable As DataTable
        Dim strSQL As String

        Count_Trans_Records = 0

        Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "TRANS")
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        mDataTable = New DataTable

        ' Set the default select statement
        strSQL = "Select COUNT(*) As MyCount from " & Global_Variables.Gbl_Trans_TableName &
            " WHERE Year = " & mYear.ToString

        If IsNothing(mRRICC) Then
            GoTo Execute
        Else
            strSQL = "Select COUNT(*) As MyCount from " & Global_Variables.Gbl_Trans_TableName &
            " WHERE Year = " & mYear.ToString &
            " And rricc = " & mRRICC.ToString
        End If

        If IsNothing(mSch) Then
            GoTo Execute
        Else
            strSQL = "Select COUNT(*) As MyCount from " & Global_Variables.Gbl_Trans_TableName &
            " WHERE Year = " & mYear.ToString &
            " And rricc = " & mRRICC.ToString &
            " And sch = " & mSch.ToString
        End If

        If IsNothing(mStartLine) Then
            GoTo Execute
        Else
            strSQL = "Select COUNT(*) As MyCount from " & Global_Variables.Gbl_Trans_TableName &
            " WHERE Year = " & mYear.ToString &
            " And rricc = " & mRRICC.ToString &
            " And sch = " & mSch.ToString &
            " And line = " & mStartLine.ToString
        End If

        If IsNothing(mEndLine) Then
            GoTo Execute
        Else
            strSQL = "Select COUNT(*) As MyCount from " & Global_Variables.Gbl_Trans_TableName &
            " WHERE Year = " & mYear.ToString &
            " And rricc = " & mRRICC.ToString &
            " And sch = " & mSch.ToString &
            " And line >= " & mStartLine.ToString &
            " And line <= " & mEndLine.ToString
        End If

Execute:

        Using daAdapter As New SqlDataAdapter(strSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        If mDataTable.Rows(0)("MyCount") > 0 Then
            Count_Trans_Records = mDataTable.Rows(0)("MyCount")
        Else
            Count_Trans_Records = 0
        End If

        mDataTable = Nothing

    End Function

    Function ConvertColumn(ByRef mColumn As String) As String

        ConvertColumn = mColumn

        Select Case mColumn
            Case "B"
                ConvertColumn = "C1"
            Case "C"
                ConvertColumn = "C2"
            Case "D"
                ConvertColumn = "C3"
            Case "E"
                ConvertColumn = "C4"
            Case "F"
                ConvertColumn = "C5"
            Case "G"
                ConvertColumn = "C6"
            Case "H"
                ConvertColumn = "C7"
            Case "I"
                ConvertColumn = "C8"
            Case "J"
                ConvertColumn = "C9"
            Case "K"
                ConvertColumn = "C10"
            Case "L"
                ConvertColumn = "C11"
            Case "M"
                ConvertColumn = "C12"
            Case "N"
                ConvertColumn = "C13"
            Case "O"
                ConvertColumn = "C14"
            Case "P"
                ConvertColumn = "C15"
        End Select

    End Function

    Function Create_Database_Command(ByVal mYear As String) As String

        Create_Database_Command = "CREATE DATABASE URCS" & mYear

    End Function

    Function Drop_Database_Command(ByVal mYear As String) As String

        Drop_Database_Command = "DROP DATABASE URCS" & mYear

    End Function

    Function Drop_Table_Command(ByVal mTable As String) As String

        Drop_Table_Command = "Drop Table " & mTable

    End Function

    Sub Create_AuditLog_Table(ByVal mDatabaseName As String, ByVal mTableName As String)
        Dim mSQLCommand As SqlCommand
        Dim mWorkStr As New StringBuilder

        mWorkStr.Append("CREATE TABLE [" & mTableName & "](")
        mWorkStr.Append("[ExecutedDateTime] [datetime] NOT NULL, ")
        mWorkStr.Append("[ActivityExecuted] [nvarchar](200) NOT NULL, ")
        mWorkStr.Append("[ExecutedBy] [nvarchar](50) NOT NULL, ")
        mWorkStr.Append("CONSTRAINT [PK_" & mTableName & "] PRIMARY KEY CLUSTERED ")
        mWorkStr.Append("([ExecutedDateTime] ASC, [ActivityExecuted] ASC, [ExecutedBy] ASC) ")
        mWorkStr.Append("WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ")
        mWorkStr.Append("ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ")
        mWorkStr.Append("ON [PRIMARY]) ON [PRIMARY]")

        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandText = mWorkStr.ToString
        mSQLCommand.ExecuteNonQuery()

        Insert_AuditTrail_Record(mDatabaseName, "Created " & mDatabaseName & "->" & mTableName & " Table.")

    End Sub

    Sub Create_AValues_Table(ByVal mDatabaseName As String, mTableName As String)
        Dim mSQLCommand As SqlCommand
        Dim mWorkStr As New StringBuilder

        mWorkStr.Append("CREATE TABLE [" & mTableName & "](")
        mWorkStr.Append("[Year] [int] Not NULL, ")
        mWorkStr.Append("[RR_Id] [int] Not NULL, ")
        mWorkStr.Append("[aCode_id] [int] Not NULL, ")
        mWorkStr.Append("[Value] [float] NULL, ")
        mWorkStr.Append("[entry_dt] [datetime2](7) NULL,")
        mWorkStr.Append("Constraint [PK_" & mTableName & "] PRIMARY KEY CLUSTERED")
        mWorkStr.Append("([Year] ASC, [RR_Id] ASC, [aCode_id] ASC)")
        mWorkStr.Append("With (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ")
        mWorkStr.Append("ALLOW_PAGE_LOCKS = On) On [PRIMARY]) On [PRIMARY]")

        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandText = mWorkStr.ToString
        mSQLCommand.ExecuteNonQuery()

        Insert_AuditTrail_Record(mDatabaseName, "Created " & mDatabaseName & "->" & mTableName & " Table.")

    End Sub

    Sub Create_CRPRES_Tables(ByVal mDatabaseName As String, ByVal mTableName As String)
        Dim mSQLCommand As SqlCommand
        Dim mWorkStr As New StringBuilder
        Dim mLooper As Integer

        mWorkStr.Append("CREATE TABLE [dbo].[" & mTableName & "](")

        For mLooper = 1 To 19
            mWorkStr.Append("[c" & mLooper.ToString & "] [int] Not NULL")
            If mLooper <> 19 Then
                mWorkStr.Append(",")
            End If
        Next

        mWorkStr.Append(") On [PRIMARY]")

        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mWorkStr.ToString
        mSQLCommand.ExecuteNonQuery()

        Insert_AuditTrail_Record(mDatabaseName, "Created " & mDatabaseName & "->" & mTableName & " Table.")

    End Sub

    'Public Sub Create_Table_CRPRES_Legacy(ByVal mDatabaseName As String, ByVal mTableName As String)
    '    Dim mSQLCommand As SqlCommand
    '    Dim mWorkStr As New StringBuilder
    '    Dim mLooper As Integer

    '    mWorkStr.Append("CREATE TABLE [dbo].[" & mTableName & "(")

    '    For mLooper = 1 To 19
    '        mWorkStr.Append("[c" & mLooper.ToString & "] [int] Not NULL")
    '        If mLooper <> 19 Then
    '            mWorkStr.Append(",")
    '        End If
    '    Next

    '    mWorkStr.Append(") On [PRIMARY]")

    '    mSQLCommand = New SqlCommand
    '    mSQLCommand.Connection = gbl_SQLConnection
    '    mSQLCommand.CommandType = CommandType.Text
    '    mSQLCommand.CommandText = mWorkStr.ToString
    '    mSQLCommand.ExecuteNonQuery()

    '    Insert_AuditTrail_Record(mDatabaseName, "Created " & mDatabaseName & "->" & mTableName & " Table.")

    'End Sub

    Public Sub Create_CRPSEG_Tables(ByVal mDatabaseName As String, ByVal mTableName As String)
        Dim mSQLCommand As SqlCommand
        Dim mWorkStr As New StringBuilder
        Dim mLooper As Integer

        mWorkStr.Append("CREATE TABLE [dbo].[" & mTableName & "](" &
            "[SerialNum] [Int] Not NULL, [SegmentNumber] [Int] Not NULL, [ShipmentSize] [Char](2) Not NULL, " &
            "[SegmentType] [Char](2) Not NULL, [CarType] [Int] Not NULL, [Ownership] [Char](1) Not NULL, " &
            "[RR_ID] [Int] Not NULL, [RR_Region] [Int] Not NULL, [Num_Cars] [Int] Not NULL, [Num_Cars_Expanded] [Int] Not NULL,")

        For mLooper = 101 To 104
            mWorkStr.Append("[L" & mLooper.ToString & "] [Int] Not NULL,")
        Next

        For mLooper = 105 To 111
            mWorkStr.Append("[L" & mLooper.ToString & "] [float] Not NULL,")
        Next

        mWorkStr.Append("[L201] [Int] Not NULL,")

        For mLooper = 202 To 216
            mWorkStr.Append("[L" & mLooper.ToString & "] [float] Not NULL,")
        Next

        For mLooper = 217 To 218
            mWorkStr.Append("[L" & mLooper.ToString & "] [Int] Not NULL,")
        Next

        For mLooper = 219 To 248
            mWorkStr.Append("[L" & mLooper.ToString & "] [float] Not NULL,")
        Next

        For mLooper = 250 To 290
            mWorkStr.Append("[L" & mLooper.ToString & "] [float] Not NULL,")
        Next

        For mLooper = 301 To 334
            mWorkStr.Append("[L" & mLooper.ToString & "] [float] Not NULL,")
        Next

        For mLooper = 401 To 499
            mWorkStr.Append("[L" & mLooper.ToString & "] [float] Not NULL,")
        Next

        mWorkStr.Append("[L499A] [float] Not NULL,")
        mWorkStr.Append("[L499B] [float] Not NULL,")
        mWorkStr.Append("[L499C] [float] Not NULL,")
        mWorkStr.Append("[L499D] [float] Not NULL,")

        For mLooper = 501 To 587
            mWorkStr.Append("[L" & mLooper.ToString & "] [float] Not NULL,")
        Next

        For mLooper = 601 To 699
            mWorkStr.Append("[L" & mLooper.ToString & "] [float] Not NULL,")
        Next

        mWorkStr.Append("[L700] [float] Not NULL) On [PRIMARY]")

        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mWorkStr.ToString
        mSQLCommand.ExecuteNonQuery()

        Insert_AuditTrail_Record(mDatabaseName, "Created " & mDatabaseName & "->" & mTableName & " Table.")


    End Sub

    Function Create_Unmasked_Rev_Table_Command(ByVal mTableName As String) As String
        Dim mworkstr As New StringBuilder

        mworkstr.Append("CREATE TABLE [dbo].[" & mTableName & "](")
        mworkstr.Append("[UnMasked_Serial_No] [int] Not NULL, ")
        mworkstr.Append("[Total_UnMasked_Rev] [int] Not NULL, ")
        mworkstr.Append("[ORR_UnMasked_Rev] [int] Not NULL, ")
        mworkstr.Append("[JRR1_UnMasked_Rev] [int] Not NULL, ")
        mworkstr.Append("[JRR2_UnMasked_Rev] [int] Not NULL, ")
        mworkstr.Append("[JRR3_UnMasked_Rev] [int] Not NULL, ")
        mworkstr.Append("[JRR4_UnMasked_Rev] [int] Not NULL, ")
        mworkstr.Append("[JRR5_UnMasked_Rev] [int] Not NULL, ")
        mworkstr.Append("[JRR6_UnMasked_Rev] [int] Not NULL, ")
        mworkstr.Append("[TRR_UnMasked_Rev] [int] Not NULL, ")
        mworkstr.Append("[U_Rev_unmasked] [int] Not NULL,")
        mworkstr.Append("Constraint [PK__" & mTableName & "] PRIMARY KEY CLUSTERED ")
        mworkstr.Append("([UnMasked_Serial_No] ASC) With (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF,")
        mworkstr.Append("IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On) On [PRIMARY]")
        mworkstr.Append(") On [PRIMARY]")

        Create_Unmasked_Rev_Table_Command = mworkstr.ToString

    End Function

    'Updated to address 913 waybilll record.  11/16/2020
    Sub Create_Masked_445_Table(ByVal mDatabaseName As String, ByVal mTableName As String)
        Dim mCommand As New SqlCommand
        Dim mWorkStr As New StringBuilder

        mWorkStr.Append("CREATE TABLE " & mTableName & "(")
        mWorkStr.Append("[Serial_No] [Varchar](6) Not NULL, ")
        mWorkStr.Append("[Wb_Num] [Int] NULL, ")
        mWorkStr.Append("[Wb_Date] [datetime] NULL,")
        mWorkStr.Append("[Acct_Period] [nvarchar](6) NULL,")
        mWorkStr.Append("[U_Cars] [Int] NULL,")
        mWorkStr.Append("[U_Car_Init] [nvarchar](4) NULL,")
        mWorkStr.Append("[U_Car_Num] [Int] NULL,")
        mWorkStr.Append("[TOFC_Serv_Code] [nvarchar](3) NULL,")
        mWorkStr.Append("[U_TC_Units] [Int] NULL,")
        mWorkStr.Append("[U_TC_Init] [nvarchar](4) NULL,")
        mWorkStr.Append("[U_TC_Num] [Int] NULL,")
        mWorkStr.Append("[STCC_W49] [nvarchar](7) NULL,")
        mWorkStr.Append("[Bill_Wght] [Int] NULL,")
        mWorkStr.Append("[Act_Wght] [Int] NULL,")
        mWorkStr.Append("[U_Rev] [Int] NULL,")
        mWorkStr.Append("[Tran_Chrg] [Int] NULL,")
        mWorkStr.Append("[Misc_Chrg] [Int] NULL,")
        mWorkStr.Append("[Intra_State_Code] [Int] NULL,")
        mWorkStr.Append("[Transit_Code] [Int] NULL,")
        mWorkStr.Append("[All_Rail_Code] [Int] NULL,")
        mWorkStr.Append("[Type_Move] [Int] NULL,")
        mWorkStr.Append("[Move_Via_Water] [Int] NULL,")
        mWorkStr.Append("[Truck_For_Rail] [Int] NULL,")
        mWorkStr.Append("[Shortline_Miles] [Int] NULL,")
        mWorkStr.Append("[Rebill] [Int] NULL,")
        mWorkStr.Append("[Stratum] [Int] NULL,")
        mWorkStr.Append("[Subsample] [Int] NULL,")
        mWorkStr.Append("[Transborder_Flg] [Int] NULL,")
        mWorkStr.Append("[Rate_Flg] [Int] NULL,")
        mWorkStr.Append("[Wb_ID] [nvarchar](25) NULL,")
        mWorkStr.Append("[Report_RR] [smallint] NULL,")
        mWorkStr.Append("[O_FSAC] [Int] NULL,")
        mWorkStr.Append("[ORR] [smallint] NULL,")
        mWorkStr.Append("[JCT1] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR1] [smallint] NULL,")
        mWorkStr.Append("[JCT2] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR2] [smallint] NULL,")
        mWorkStr.Append("[JCT3] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR3] [smallint] NULL,")
        mWorkStr.Append("[JCT4] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR4] [smallint] NULL,")
        mWorkStr.Append("[JCT5] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR5] [smallint] NULL,")
        mWorkStr.Append("[JCT6] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR6] [smallint] NULL,")
        mWorkStr.Append("[JCT7] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR7] [smallint] NULL,")
        mWorkStr.Append("[JCT8] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR8] [smallint] NULL,")
        mWorkStr.Append("[JCT9] [nvarchar](5) NULL,")
        mWorkStr.Append("[TRR] [smallint] NULL,")
        mWorkStr.Append("[T_FSAC] [Int] NULL,")
        mWorkStr.Append("[Pop_Cnt] [Int] NULL,")
        mWorkStr.Append("[Stratum_Cnt] [Int] NULL,")
        mWorkStr.Append("[Report_Period] [tinyint] NULL,")
        mWorkStr.Append("[Car_Own_Mark] [nvarchar](4) NULL,")
        mWorkStr.Append("[Car_Lessee_Mark] [nvarchar](4) NULL,")
        mWorkStr.Append("[Car_Cap] [Int] NULL,")
        mWorkStr.Append("[Nom_Car_Cap] [smallint] NULL,")
        mWorkStr.Append("[Tare] [smallint] NULL,")
        mWorkStr.Append("[Outside_L] [Int] NULL,")
        mWorkStr.Append("[Outside_W] [smallint] NULL,")
        mWorkStr.Append("[Outside_H] [smallint] NULL,")
        mWorkStr.Append("[Ex_Outside_H] [smallint] NULL,")
        mWorkStr.Append("[Type_Wheel] [nvarchar](1) NULL,")
        mWorkStr.Append("[No_Axles] [nvarchar](1) NULL,")
        mWorkStr.Append("[Draft_Gear] [tinyint] NULL,")
        mWorkStr.Append("[Art_Units] [Int] NULL,")
        mWorkStr.Append("[Pool_Code] [Int] NULL,")
        mWorkStr.Append("[Car_Typ] [nvarchar](4) NULL,")
        mWorkStr.Append("[Mech] [nvarchar](4) NULL,")
        mWorkStr.Append("[Lic_St] [nvarchar](2) NULL,")
        mWorkStr.Append("[Mx_Wght_Rail] [smallint] NULL,")
        mWorkStr.Append("[O_SPLC] [nvarchar](6) NULL,")
        mWorkStr.Append("[T_SPLC] [nvarchar](6) NULL,")
        mWorkStr.Append("[U_Fuel_Surchg] [numeric](13,0) NULL,")
        mWorkStr.Append("[Err_Code1] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code2] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code3] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code4] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code5] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code6] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code7] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code8] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code9] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code10] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code11] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code12] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code13] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code14] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code15] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code16] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code17] [tinyint] NULL,")
        mWorkStr.Append("[Car_Own] [nvarchar](1) NULL,")
        mWorkStr.Append("[TOFC_Unit_Type] [nvarchar](4) NULL,")
        mWorkStr.Append("[ALK_Flg] [tinyint] NULL,")
        mWorkStr.Append("[Tracking_No] [numeric](13, 0) NULL,")
        mWorkStr.Append("Constraint [PK__" & mTableName & "] PRIMARY KEY CLUSTERED ")
        mWorkStr.Append("([Serial_No] Asc) With ")
        mWorkStr.Append("(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ")
        mWorkStr.Append("ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On) On [PRIMARY]) On [PRIMARY]")

        mCommand.Connection = gbl_SQLConnection
        mCommand.CommandType = CommandType.Text
        mCommand.CommandText = mWorkStr.ToString
        mCommand.ExecuteNonQuery()
        Insert_AuditTrail_Record(mDatabaseName, "Created " & mTableName & " Table.")

    End Sub

    Public Sub Create_STCC_Table()
        Dim mStrSQL As StringBuilder

        mStrSQL = New StringBuilder
        mStrSQL.Append("CREATE TABLE " & Gbl_STCC_Codes_TableName)
        mStrSQL.Append(" ( ")
        mStrSQL.Append("[STCC] [nvarchar](7) Not NULL PRIMARY KEY,")
        mStrSQL.Append("[STCC_Code2] [nvarchar](7) NULL,")
        mStrSQL.Append("[Transaction_Date] [date] NULL,")
        mStrSQL.Append("[Transaction_Time] [time](0) NULL,")
        mStrSQL.Append("[Status_Code] [nvarchar](1) NULL,")
        mStrSQL.Append("[Effective_Date] [date] NULL,")
        mStrSQL.Append("[Header1] [nvarchar](2) NULL,")
        mStrSQL.Append("[Header2] [nvarchar](3) NULL,")
        mStrSQL.Append("[Header3] [nvarchar](4) NULL,")
        mStrSQL.Append("[Header4] [nvarchar](5) NULL,")
        mStrSQL.Append("[STCC_Desc_15] [nvarchar](15) NULL,")
        mStrSQL.Append("[Alternate_Num] [nvarchar](2) NULL,")
        mStrSQL.Append("[Expiration_Date] [date] NULL,")
        mStrSQL.Append("[Description_US] [nvarchar](300) NULL,")
        mStrSQL.Append("[Primary_Hazard_Class_Intl] [varchar](4) NULL,")
        mStrSQL.Append("[NOS_Flg_Intl] [nvarchar](1) NULL,")
        mStrSQL.Append("[Technical_Name_Intl] [nvarchar](125) NULL,")
        mStrSQL.Append("[UN_Number_Intl] [nvarchar](6) NULL,")
        mStrSQL.Append("[Packing_Grp_Intl] [nvarchar](1) NULL,")
        mStrSQL.Append("[Poison_Flg_intl] [nvarchar](1) NULL,")
        mStrSQL.Append("[Primary_Placard_Intl] [nvarchar](2) NULL,")
        mStrSQL.Append("[Shipping_Name_Intl] [nvarchar](125) NULL,")
        mStrSQL.Append("[Primary_Class_CN] [nvarchar](4) NULL,")
        mStrSQL.Append("[Sub_Class_CN] [nvarchar](9) NULL,")
        mStrSQL.Append("[CN_Orig_US_Dest_Flg] [nvarchar](1) NULL,")
        mStrSQL.Append("[Emer_Resp_Asst_Plan_CN_Flg] [nvarchar](4) NULL,")
        mStrSQL.Append("[Primary_Placard_CN] [nvarchar](2) NULL,")
        mStrSQL.Append("[Special_Commodity_CN_Flg] [nvarchar](1) NULL,")
        mStrSQL.Append("[NOS_Flag_CN] [nvarchar](1) NULL,")
        mStrSQL.Append("[Sub_Placard_CN] [nvarchar](2) NULL,")
        mStrSQL.Append("[Technical_Name_CN] [nvarchar](125) NULL,")
        mStrSQL.Append("[UN_Number_CN] [nvarchar](6) NULL,")
        mStrSQL.Append("[Packing_Grp_CN] [nvarchar](1) NULL,")
        mStrSQL.Append("[Poison_Flg_CN] [nvarchar](1) NULL,")
        mStrSQL.Append("[Shipping_Name_CN] [nvarchar](125) NULL,")
        mStrSQL.Append("[EPA_Waste_Stream_Num_US] [nvarchar](18) NULL,")
        mStrSQL.Append("[Haz_Placard_US] [nvarchar](2) NULL,")
        mStrSQL.Append("[Primary_Hazard_Class_US] [nvarchar](4) NULL,")
        mStrSQL.Append("[Sub_Hazard_US] [nvarchar](6) NULL,")
        mStrSQL.Append("[Hazard_Zone_US] [nvarchar](1) NULL,")
        mStrSQL.Append("[NOS_Flag_US] [nvarchar](1) NULL,")
        mStrSQL.Append("[Sub_Placard_US] [nvarchar](2) NULL,")
        mStrSQL.Append("[Technical_Name_US] [nvarchar](125) NULL,")
        mStrSQL.Append("[UN_NA_ID_Num_US] [nvarchar](6) NULL,")
        mStrSQL.Append("[US_Orig_CN_Dest_Flg] [nvarchar](1) NULL,")
        mStrSQL.Append("[Packing_Grp_US] [nvarchar](1) NULL,")
        mStrSQL.Append("[Poison_Flg_US] [nvarchar](1) NULL,")
        mStrSQL.Append("[Primary_Placard_US] [nvarchar](2) NULL,")
        mStrSQL.Append("[Shipping_Name_US] [nvarchar](125) NULL,")
        mStrSQL.Append("[OT55_Flg] [nvarchar](1) NULL,")
        mStrSQL.Append("[Reportable_Quantity_Flg_US] [nvarchar](1) NULL,")
        mStrSQL.Append("[Marine_Pollutant_Flg_US] [nvarchar](1) NULL,")
        mStrSQL.Append("[HazMat_Name_US] [nvarchar](125) NULL,")
        mStrSQL.Append("[Marine_Pollutant_Name_US] [nvarchar](125) NULL,")
        mStrSQL.Append("[Special_Shipping_Name_Flg_CN] [nvarchar](1) NULL,")
        mStrSQL.Append("[Spcl_Proper_Ship_Name_Flg_Intl] [nvarchar](1) NULL,")
        mStrSQL.Append("[Spcl_Proper_Ship_Name_Flg_US] [nvarchar](1) NULL,")
        mStrSQL.Append("[Intermodal_Flg_CN] [nvarchar](1) NULL,")
        mStrSQL.Append("[Internodal_Flg_Intl] [nvarchar](1) NULL,")
        mStrSQL.Append("[Intermodal_Flg_US] [nvarchar](1) NULL,")
        mStrSQL.Append("[RSSM_Flg_US] [nvarchar](2) NULL,")
        mStrSQL.Append("[Alpha_Desc] [nvarchar](277) NULL, ")
        mStrSQL.Append("[STCC_Desc] [nvarchar](250) NULL, ")
        mStrSQL.Append("[Sub_Placard_Intl] [nvarchar](2) NULL, ")
        mStrSQL.Append("[HazMat_Class_Intl] [nvarchar](9) NULL, ")
        mStrSQL.Append("[Deletion_Date] [date] NULL, ")
        mStrSQL.Append("[Alt_Shipping_Name_CN] [nvarchar](625) NULL, ")
        mStrSQL.Append("[Alt_Proper_Ship_Name_Intl] [nvarchar](625) NULL, ")
        mStrSQL.Append("[Alt_Proper_Ship_Name_US] [nvarchar](625) NULL, ")
        mStrSQL.Append("[Int_Harm_Code] [nvarchar](12) NULL, ")
        mStrSQL.Append("[Indus_Class] [nvarchar](4) NULL, ")
        mStrSQL.Append("[Inter_SIC] [nvarchar](4) NULL, ")
        mStrSQL.Append("[Dom_Canada] [nvarchar](5) NULL, ")
        mStrSQL.Append("[CS_54_Code] [nvarchar](2) NULL, ")
        mStrSQL.Append("[CS_54_Name] [nvarchar](30) NULL, ")
        mStrSQL.Append("[Dereg_Flg] [nvarchar](1) NULL, ")
        mStrSQL.Append("[Dereg_Date] [Date] NULL, ")
        mStrSQL.Append("[Car_Grade] [nvarchar](1) NULL, ")
        mStrSQL.Append("[STCC_Repl_Code] [nvarchar](7) NULL ) On [PRIMARY]")

        ' Check/Open the SQL connection
        OpenADOConnection(Gbl_Controls_Database_Name)

        ' Create the table in SQL
        gbl_ADOConnection.Execute(mStrSQL.ToString)

        ' Write the command to create the index for the new table
        'mStrSQL = New StringBuilder
        'mStrSQL.Append("ALTER TABLE " & Gbl_STCC_Codes_TableName)
        'mStrSQL.Append(" ADD CONSTRAINT pk_" & Gbl_STCC_Codes_TableName)
        'mStrSQL.Append(" PRIMARY KEY CLUSTERED ")
        'mStrSQL.Append(" (STCC) With (STATISTICS_NORECOMPUTE = OFF,")
        'mStrSQL.Append(" IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ")
        'mStrSQL.Append(" ALLOW_PAGE_LOCKS = On) On [PRIMARY]")

        ' Create the index In SQL
        'gbl_ADOConnection.Execute(mStrSQL.ToString)

    End Sub

    'Public Sub Create_Waybill_Table_455()
    '    ' Commented out as this is not correct (missing fuel surcharge field) and it is not used (0 references)
    '    Dim mStrSQL As StringBuilder

    '    mStrSQL = New StringBuilder
    '    mStrSQL.Append("CREATE TABLE " & Gbl_RailInc_455_TableName)
    '    mStrSQL.Append(" ( ")
    '    mStrSQL.Append("[Serial_No] [int] Not NULL,")
    '    mStrSQL.Append("[Wb_Num] [int] NULL,")
    '    mStrSQL.Append("[Wb_Date] [datetime] NULL,")
    '    mStrSQL.Append("[Acct_Period] [nvarchar](6) NULL,")
    '    mStrSQL.Append("[U_Cars] [int] NULL,")
    '    mStrSQL.Append("[U_Car_Init] [nvarchar](4) NULL,")
    '    mStrSQL.Append("[U_Car_Num] [int] NULL,")
    '    mStrSQL.Append("[TOFC_Serv_Code] [nvarchar](3) NULL,")
    '    mStrSQL.Append("[U_TC_Units] [int] NULL,")
    '    mStrSQL.Append("[U_TC_Init] [nvarchar](4) NULL,")
    '    mStrSQL.Append("[U_TC_Num] [int] NULL,")
    '    mStrSQL.Append("[STCC_W49] [nvarchar](7) NULL,")
    '    mStrSQL.Append("[Bill_Wght] [int] NULL,")
    '    mStrSQL.Append("[Act_Wght] [int] NULL,")
    '    mStrSQL.Append("[U_Rev] [int] NULL,")
    '    mStrSQL.Append("[Tran_Chrg] [int] NULL,")
    '    mStrSQL.Append("[Misc_Chrg] [int] NULL,")
    '    mStrSQL.Append("[Intra_State_Code] [int] NULL,")
    '    mStrSQL.Append("[Transit_Code] [int] NULL,")
    '    mStrSQL.Append("[All_Rail_Code] [int] NULL,")
    '    mStrSQL.Append("[Type_Move] [int] NULL,")
    '    mStrSQL.Append("[Move_Via_Water] [int] NULL,")
    '    mStrSQL.Append("[Truck_For_Rail] [int] NULL,")
    '    mStrSQL.Append("[Shortline_Miles] [int] NULL,")
    '    mStrSQL.Append("[Rebill] [int] NULL,")
    '    mStrSQL.Append("[Stratum] [int] NULL,")
    '    mStrSQL.Append("[Subsample] [int] NULL,")
    '    mStrSQL.Append("[Rate_Flg] [int] NULL,")
    '    mStrSQL.Append("[Wb_ID] [nvarchar](25) NULL,")
    '    mStrSQL.Append("[Report_RR] [smallint] NULL,")
    '    mStrSQL.Append("[O_FSAC] [int] NULL,")
    '    mStrSQL.Append("[ORR] [smallint] NULL,")
    '    mStrSQL.Append("[JCT1] [nvarchar](5) NULL,")
    '    mStrSQL.Append("[JRR1] [smallint] NULL,")
    '    mStrSQL.Append("[JCT2] [nvarchar](5) NULL,")
    '    mStrSQL.Append("[JRR2] [smallint] NULL,")
    '    mStrSQL.Append("[JCT3] [nvarchar](5) NULL,")
    '    mStrSQL.Append("[JRR3] [smallint] NULL,")
    '    mStrSQL.Append("[JCT4] [nvarchar](5) NULL,")
    '    mStrSQL.Append("[JRR4] [smallint] NULL,")
    '    mStrSQL.Append("[JCT5] [nvarchar](5) NULL,")
    '    mStrSQL.Append("[JRR5] [smallint] NULL,")
    '    mStrSQL.Append("[JCT6] [nvarchar](5) NULL,")
    '    mStrSQL.Append("[JRR6] [smallint] NULL,")
    '    mStrSQL.Append("[JCT7] [nvarchar](5) NULL,")
    '    mStrSQL.Append("[JRR7] [smallint] NULL,")
    '    mStrSQL.Append("[JCT8] [nvarchar](5) NULL,")
    '    mStrSQL.Append("[JRR8] [smallint] NULL,")
    '    mStrSQL.Append("[JCT9] [nvarchar](5) NULL,")
    '    mStrSQL.Append("[TRR] [smallint] NULL,")
    '    mStrSQL.Append("[T_FSAC] [int] NULL,")
    '    mStrSQL.Append("[Pop_Cnt] [int] NULL,")
    '    mStrSQL.Append("[Stratum_Cnt] [int] NULL,")
    '    mStrSQL.Append("[Report_Period] [tinyint] NULL,")
    '    mStrSQL.Append("[Car_Own_Mark] [nvarchar](4) NULL,")
    '    mStrSQL.Append("[Car_Lessee_Mark] [nvarchar](4) NULL,")
    '    mStrSQL.Append("[Car_Lessee_Mark] [nvarchar](4) NULL,")
    '    mStrSQL.Append("[Nom_Car_Cap] [smallint] NULL,")
    '    mStrSQL.Append("[Tare] [smallint] NULL,")
    '    mStrSQL.Append("[Outside_L] [int] NULL,")
    '    mStrSQL.Append("[Outside_W] [smallint] NULL,")
    '    mStrSQL.Append("[Outside_H] [smallint] NULL,")
    '    mStrSQL.Append("[Ex_Outside_H] [smallint] NULL,")
    '    mStrSQL.Append("[Type_Wheel] [nvarchar](1) NULL,")
    '    mStrSQL.Append("[No_Axles] [nvarchar](1) NULL,")
    '    mStrSQL.Append("[Draft_Gear] [tinyint] NULL,")
    '    mStrSQL.Append("[Art_Units] [int] NULL,")
    '    mStrSQL.Append("[Pool_Code] [int] NULL,")
    '    mStrSQL.Append("[Car_Typ] [nvarchar](4) NULL,")
    '    mStrSQL.Append("[Mech] [nvarchar](4) NULL,")
    '    mStrSQL.Append("[Lic_St] [nvarchar](2) NULL,")
    '    mStrSQL.Append("[Mx_Wght_Rail] [smallint] NULL,")
    '    mStrSQL.Append("[O_SPLC] [nvarchar](6) NULL,")
    '    mStrSQL.Append("[T_SPLC] [nvarchar](6) NULL,")
    '    mStrSQL.Append("[Err_Code1] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code2] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code3] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code4] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code5] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code6] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code7] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code8] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code9] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code10] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code11] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code12] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code13] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code14] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code15] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code16] [tinyint] NULL,")
    '    mStrSQL.Append("[Err_Code17] [tinyint] NULL,")
    '    mStrSQL.Append("[Car_Own] [nvarchar](1) NULL,")
    '    mStrSQL.Append("[TOFC_Unit_Type] [nvarchar](4) NULL,")
    '    mStrSQL.Append("[U_Fuel_SurChrg] [numeric](13,0) NULL,")
    '    mStrSQL.Append("[ALK_Flag] [tinyint] NULL,")
    '    mStrSQL.Append("[Tracking_No] [nvarchar](13),")
    '    mStrSQL.Append(" CONSTRAINT [PK_" & Gbl_RailInc_455_TableName & "] PRIMARY KEY CLUSTERED ")
    '    mStrSQL.Append("([Serial_No] ASC) With (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF,")
    '    mStrSQL.Append("IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On)")
    '    mStrSQL.Append(" [PRIMARY] ) On [PRIMARY]")

    '    ' Check/Open the SQL connection
    '    OpenADOConnection(Gbl_Controls_Database_Name)

    '    ' Create the table in SQL
    '    gbl_ADOConnection.Execute(mStrSQL.ToString)

    'End Sub

    Public Sub Create_CSM_Table()
        Dim mStrSQL As StringBuilder

        mStrSQL = New StringBuilder
        mStrSQL.Append("CREATE TABLE " & Gbl_CSM_TableName)
        mStrSQL.Append(" ( ")
        mStrSQL.Append("[Road_Mark] [nvarchar](4) Not NULL,")
        mStrSQL.Append("[FSAC] [nvarchar](5) Not NULL,")
        mStrSQL.Append("[Eff_Date] [date],")
        mStrSQL.Append("[Eff_Time] [time],")
        mStrSQL.Append("[Sta_Status] [nvarchar](2),")
        mStrSQL.Append("[R260] [nvarchar](5),")
        mStrSQL.Append("[SPLC] [nvarchar](9),")
        mStrSQL.Append("[OPSL] [nvarchar](7),")
        mStrSQL.Append("[Loc_Type] [nvarchar](5),")
        mStrSQL.Append("[AAR_Status] [nvarchar](1),")
        mStrSQL.Append("[Loc_Name] [nvarchar](30),")
        mStrSQL.Append("[OPSL_Name] [nvarchar](30),")
        mStrSQL.Append("[Loc_Geo_Name] [nvarchar](30),")
        mStrSQL.Append("[Loc_County] [nvarchar](30),")
        mStrSQL.Append("[Loc_St] [varchar](2),")
        mStrSQL.Append("[Loc_Cntry] [nvarchar](2),")
        mStrSQL.Append("[Loc_Zip] [nvarchar](11),")
        mStrSQL.Append("[Rate_Base_SPLC] [nvarchar](9),")
        mStrSQL.Append("[Rate_Base_City] [nvarchar](30),")
        mStrSQL.Append("[Rate_Base_St] [nvarchar](2),")
        mStrSQL.Append("[Rev_Sw_SPLC] [nvarchar](9),")
        mStrSQL.Append("[Rev_Sw_City] [nvarchar](30),")
        mStrSQL.Append("[Rev_Sw_St] [nvarchar](2),")
        mStrSQL.Append("[CIF_ID] [nvarchar](2),")
        mStrSQL.Append("[Imp_Exp_Flg] [nvarchar](1),")
        mStrSQL.Append("[Customs_Flg] [nvarchar](1),")
        mStrSQL.Append("[Grain_Flg] [nvarchar](1),")
        mStrSQL.Append("[Auto_Ramp_Flg] [nvarchar](1),")
        mStrSQL.Append("[Intermodal_Flg] [nvarchar](5),")
        mStrSQL.Append("[Embargo_Flg] [nvarchar](1),")
        mStrSQL.Append("[Oper_Plate] [nvarchar](1),")
        mStrSQL.Append("[Oper_Wght] [nvarchar](4),")
        mStrSQL.Append("[FIPS] [nvarchar](5),")
        mStrSQL.Append("[BEA] [nvarchar](3),")
        mStrSQL.Append("[BEA_Name] [nvarchar](60),")
        mStrSQL.Append("[CEA] [nvarchar](4),")
        mStrSQL.Append("[Lat] [nvarchar](9),")
        mStrSQL.Append("[Long] [nvarchar](18),")
        mStrSQL.Append("[Reload] [nvarchar](5),")
        mStrSQL.Append("[Geo_SPLC] [nvarchar](9),")
        mStrSQL.Append("[Customs_CIF] [nvarchar](13),")
        mStrSQL.Append("[Time_Zone] [nvarchar](2),")
        mStrSQL.Append("[Daylight_Ind] [nvarchar](1),")
        mStrSQL.Append("[OPSL_Notes] [nvarchar](40),")
        mStrSQL.Append("[OPSL_Ref] [nvarchar](3),")
        mStrSQL.Append("[Exp_Date] [date],")
        mStrSQL.Append("[Interswitch_Area] [nvarchar](9),")
        mStrSQL.Append("[Name_333] [nvarchar](9),")
        mStrSQL.Append("[New_Road_Name] [nvarchar](4),")
        mStrSQL.Append("[New_FSAC] [nvarchar](5),")
        mStrSQL.Append("[Last_Maint_Stamp] [nvarchar](26) Not NULL,")
        mStrSQL.Append("[Last_Trans_Type] [nvarchar](1),")
        mStrSQL.Append("[Last_Update_Type] [nvarchar](1),")
        mStrSQL.Append("[Last_Road_Mark] [nvarchar](4),")
        mStrSQL.Append("[Last_Date] [date],")
        mStrSQL.Append("[Last_Time] [date])")

        ' Check/Open the SQL connection
        OpenADOConnection(Gbl_Controls_Database_Name)

        ' Create the table in SQL
        gbl_ADOConnection.Execute(mStrSQL.ToString)

        ' Write the command to create the index for the new table
        mStrSQL = New StringBuilder
        mStrSQL.Append("ALTER TABLE " & Gbl_CSM_TableName)
        mStrSQL.Append(" ADD CONSTRAINT pk_" & Gbl_CSM_TableName)
        mStrSQL.Append(" PRIMARY KEY CLUSTERED ")
        mStrSQL.Append(" (Road_Mark Asc, FSAC Asc, Last_Maint_Stamp Asc) With (STATISTICS_NORECOMPUTE = OFF,")
        mStrSQL.Append(" IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ")
        mStrSQL.Append(" ALLOW_PAGE_LOCKS = On) On [PRIMARY]")

        ' Create the index In SQL
        gbl_ADOConnection.Execute(mStrSQL.ToString)

    End Sub

    Public Sub Create_Marks_Table()
        Dim mStrSQL As StringBuilder

        mStrSQL = New StringBuilder
        mStrSQL.Append("CREATE TABLE " & Gbl_Marks_Tablename)
        mStrSQL.Append(" ( ")
        mStrSQL.Append("[Road_Mark] [nvarchar](4) Not NULL,")
        mStrSQL.Append("[Owner] [nvarchar](4) NULL,")
        mStrSQL.Append("[Release_Ind] [nvarchar](1) NULL,")
        mStrSQL.Append("[Placement_Ind] [nvarchar](1) NULL,")
        mStrSQL.Append("[Mark_Type] [nvarchar](3) NULL,")
        mStrSQL.Append("[Exp_Date] [date] NULL,")
        mStrSQL.Append("[Mark_Name] [nvarchar](105) NULL,")
        mStrSQL.Append("[R260_Num] [nvarchar](3) NULL,")
        mStrSQL.Append("[Eff_Date] [date] NULL,")
        mStrSQL.Append("[CIF_ID] [nvarchar](13) NULL,")
        mStrSQL.Append("[Last_Maint_Stamp] [nvarchar](26) Not NULL ) On [PRIMARY]")

        ' Check/Open the SQL connection
        OpenADOConnection(Gbl_Waybill_Database_Name)

        ' Create the table in SQL
        gbl_ADOConnection.Execute(mStrSQL.ToString)

        ' Write the command to create the index for the new table
        mStrSQL = New StringBuilder
        mStrSQL.Append("ALTER TABLE " & Gbl_Marks_Tablename)
        mStrSQL.Append(" ADD CONSTRAINT pk_" & Gbl_Marks_Tablename)
        mStrSQL.Append(" PRIMARY KEY CLUSTERED ")
        mStrSQL.Append(" (Road_Mark ASC, Last_Maint_Stamp ASC) With (STATISTICS_NORECOMPUTE = OFF,")
        mStrSQL.Append(" IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ")
        mStrSQL.Append(" ALLOW_PAGE_LOCKS = On) On [PRIMARY]")

        ' Create the index In SQL
        gbl_ADOConnection.Execute(mStrSQL.ToString)

    End Sub

    Public Sub Create_Productivity_Table()
        Dim mSQLCommand As SqlCommand
        Dim mStrSQL As StringBuilder

        mStrSQL = New StringBuilder
        mStrSQL.Append("CREATE TABLE " & Gbl_Productivity_TableName)
        mStrSQL.Append(" ( ")
        mStrSQL.Append("[Year] [smallint] Not NULL,")
        mStrSQL.Append("[Length_Of_Haul_Stratum] [smallint] Not NULL,")
        mStrSQL.Append("[Car_Type_Stratum] [smallint] Not NULL,")
        mStrSQL.Append("[Lading_Weight_Stratum] [smallint] Not NULL,")
        mStrSQL.Append("[Cars_Stratum] [smallint] Not NULL,")
        mStrSQL.Append("[Revenue] [bigint] Not NULL,")
        mStrSQL.Append("[Ton_Miles] [int] Not NULL,")
        mStrSQL.Append("[Waybill_Records] [int] Not NULL)")

        ' Check/Open the SQL connection
        OpenSQLConnection(Gbl_Waybill_Database_Name)

        ' Create the table in SQL
        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandText = mStrSQL.ToString
        mSQLCommand.ExecuteNonQuery()

        ' Write the command to create the index for the new table
        mStrSQL = New StringBuilder
        mStrSQL.Append("ALTER TABLE " & Gbl_Productivity_TableName)
        mStrSQL.Append(" ADD CONSTRAINT pk_" & Gbl_Productivity_TableName)
        mStrSQL.Append(" PRIMARY KEY CLUSTERED ")
        mStrSQL.Append(" ([Year] ASC,[Length_Of_Haul_Stratum] ASC,[Car_Type_Stratum] ASC," &
                       "[Lading_Weight_Stratum] ASC,[Cars_Stratum] ASC) With (STATISTICS_NORECOMPUTE = OFF,")
        mStrSQL.Append(" IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ")
        mStrSQL.Append(" ALLOW_PAGE_LOCKS = On) On [PRIMARY]")

        ' Create the index In SQL
        mSQLCommand.CommandText = mStrSQL.ToString
        mSQLCommand.ExecuteNonQuery()

    End Sub

    Sub Create_Errors_Table(mDatabaseName As String, mTableName As String)
        Dim mSQLCommand As SqlCommand
        Dim mWorkStr As New StringBuilder

        mWorkStr.Append("CREATE TABLE [dbo].[U_ERRORS](")
        mWorkStr.Append("[error_seq] [Int] IDENTITY(1, 1) Not NULL,")
        mWorkStr.Append("[error_data] [nvarchar](255) NULL,")
        mWorkStr.Append("[error_timestamp] [datetime2](7) NULL,")
        mWorkStr.Append("[error_message] [nvarchar](255) NULL,")
        mWorkStr.Append("[error_stack] [ntext] NULL,")
        mWorkStr.Append("[error_location] [nvarchar](255) NULL,")
        mWorkStr.Append("PRIMARY KEY CLUSTERED ")
        mWorkStr.Append("([error_seq] Asc) With ")
        mWorkStr.Append("(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On)")
        mWorkStr.Append("On [PRIMARY] ) On [PRIMARY] TEXTIMAGE_ON [PRIMARY]")

        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mWorkStr.ToString
        mSQLCommand.ExecuteNonQuery()

        Insert_AuditTrail_Record(mDatabaseName, "Created " & mDatabaseName & "->" & mTableName & " Table.")

    End Sub

    Public Sub Create_EValues_Table(mDatabaseName As String, mTableName As String)
        Dim mSQLCommand As SqlCommand
        Dim mWorkStr As New StringBuilder

        mWorkStr.Append("CREATE TABLE [dbo].[U_EVALUES](")
        mWorkStr.Append("[Year] [Int] Not NULL,")
        mWorkStr.Append("[RR_Id] [Int] Not NULL,")
        mWorkStr.Append("[eCode_id] [Int] Not NULL,")
        mWorkStr.Append("[Value] [float] NULL,")
        mWorkStr.Append("[entry_dt] [datetime2](7) NULL,")
        mWorkStr.Append("Constraint [PK_U_EVALUES] PRIMARY KEY CLUSTERED")
        mWorkStr.Append("([Year] Asc, [RR_Id] ASC, [eCode_id] Asc)")
        mWorkStr.Append("With (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On)")
        mWorkStr.Append("On [PRIMARY]) On [PRIMARY]")

        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mWorkStr.ToString
        mSQLCommand.ExecuteNonQuery()

        Insert_AuditTrail_Record(mDatabaseName, "Created " & mDatabaseName & "->" & mTableName & " Table.")

    End Sub

    Public Sub Create_Substitutions_Table(mDatabaseName As String, mTableName As String)
        Dim mWorkStr As New StringBuilder
        Dim mSQLCommand As SqlCommand

        mWorkStr.Append("CREATE TABLE [dbo]." & mTableName & "(")
        mWorkStr.Append("[Year] [Int] Not NULL,")
        mWorkStr.Append("[RR_Id] [Int] Not NULL,")
        mWorkStr.Append("[eCode_id] [Int] Not NULL,")
        mWorkStr.Append("[Value] [float] NULL,")
        mWorkStr.Append("[Final_Value] [float] NULL,")
        mWorkStr.Append("[Notes] [nvarchar](100) NULL,")
        mWorkStr.Append("[entry_dt] [datetime2](7) NULL,")
        mWorkStr.Append("Constraint [PK_U_SUBSTITUTIONS] PRIMARY KEY CLUSTERED")
        mWorkStr.Append("([Year] Asc, [RR_Id] ASC, [eCode_id] Asc)")
        mWorkStr.Append("With (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On)")
        mWorkStr.Append("On [PRIMARY] ) On [PRIMARY]")

        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mWorkStr.ToString
        mSQLCommand.ExecuteNonQuery()

        Insert_AuditTrail_Record(mDatabaseName, "Created " & mDatabaseName & "->" & mTableName & " Table.")

    End Sub

    Public Sub Create_MakeWhole_Factors_Table(mDatabaseName As String, mTableName As String)
        Dim mWorkStr As New StringBuilder
        Dim mSQLCommand As SqlCommand

        mWorkStr.Append("CREATE TABLE [dbo].[UT_MAKEWHOLE_FACTORS](")
        mWorkStr.Append("[Year] [varchar](4) Not NULL,")
        mWorkStr.Append("[RR_ID] [Int] Not NULL,")
        mWorkStr.Append("[eCode_ID] [Int] Not NULL,")
        mWorkStr.Append("[eCode_Value] [float] Not NULL")
        mWorkStr.Append(") On [PRIMARY]")

        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mWorkStr.ToString
        mSQLCommand.ExecuteNonQuery()

    End Sub

    Public Sub Create_PUWS_Masked_Table(mYear As String)
        Dim mWorkStr As New StringBuilder
        Dim mSQLCommand As SqlCommand

        mWorkStr.Append("CREATE TABLE [dbo].[PUWS")
        mWorkStr.Append(mYear)
        mWorkStr.Append("_Masked] (")
        mWorkStr.Append("[WB_Date] [nvarchar](6) NULL,")
        mWorkStr.Append("[acct_period] [nvarchar](4) NULL,")
        mWorkStr.Append("[u_cars] [Int] NULL,")
        mWorkStr.Append("[car_own] [nvarchar](1) NULL,")
        mWorkStr.Append("[car_typ] [nvarchar](4) NULL,")
        mWorkStr.Append("[mech] [nvarchar](4) NULL,")
        mWorkStr.Append("[stb_car_typ] [tinyint] NULL,")
        mWorkStr.Append("[tofc_serv_code] [nvarchar](3) NULL,")
        mWorkStr.Append("[u_tc_units] [Int] NULL,")
        mWorkStr.Append("[tofc_own_code] [nvarchar](1) NULL,")
        mWorkStr.Append("[tofc_unit_type] [nvarchar](1) NULL,")
        mWorkStr.Append("[haz_bulk] [nvarchar](1) NULL,")
        mWorkStr.Append("[stcc] [nvarchar](7) NULL,")
        mWorkStr.Append("[bill_wght_tons] [Int] NULL,")
        mWorkStr.Append("[act_wght] [Int] NULL,")
        mWorkStr.Append("[u_rev] [Int] NULL,")
        mWorkStr.Append("[tran_chrg] [Int] NULL,")
        mWorkStr.Append("[misc_chrg] [Int] NULL,")
        mWorkStr.Append("[intra_state_code] [Int] NULL,")
        mWorkStr.Append("[type_move] [Int] NULL,")
        mWorkStr.Append("[all_rail_code] [Int] NULL,")
        mWorkStr.Append("[move_via_water] [Int] NULL,")
        mWorkStr.Append("[transit_code] [Int] NULL,")
        mWorkStr.Append("[truck_for_rail] [Int] NULL,")
        mWorkStr.Append("[rebill] [Int] NULL,")
        mWorkStr.Append("[shortline_miles] [Int] NULL,")
        mWorkStr.Append("[stratum] [Int] NULL,")
        mWorkStr.Append("[subsample] [Int] NULL,")
        mWorkStr.Append("[exp_factor] [Int] NULL,")
        mWorkStr.Append("[exp_factor_th] [Int] NULL,")
        mWorkStr.Append("[jf] [Int] NULL,")
        mWorkStr.Append("[o_bea] [smallint] NULL,")
        mWorkStr.Append("[o_ft] [smallint] NULL,")
        mWorkStr.Append("[jct1_st] [nvarchar](2) NULL,")
        mWorkStr.Append("[jct2_st] [nvarchar](2) NULL,")
        mWorkStr.Append("[jct3_st] [nvarchar](2) NULL,")
        mWorkStr.Append("[jct4_st] [nvarchar](2) NULL,")
        mWorkStr.Append("[jct5_st] [nvarchar](2) NULL,")
        mWorkStr.Append("[jct6_st] [nvarchar](2) NULL,")
        mWorkStr.Append("[jct7_st] [nvarchar](2) NULL,")
        mWorkStr.Append("[jct8_st] [nvarchar](2) NULL,")
        mWorkStr.Append("[jct9_st] [nvarchar](2) NULL,")
        mWorkStr.Append("[t_bea] [smallint] NULL,")
        mWorkStr.Append("[t_ft] [smallint] NULL,")
        mWorkStr.Append("[report_period] [tinyint] NULL,")
        mWorkStr.Append("[car_cap] [Int] NULL,")
        mWorkStr.Append("[nom_car_cap] [smallint] NULL,")
        mWorkStr.Append("[tare] [smallint] NULL,")
        mWorkStr.Append("[outside_l] [Int] NULL,")
        mWorkStr.Append("[outside_w] [smallint] NULL,")
        mWorkStr.Append("[outside_h] [smallint] NULL,")
        mWorkStr.Append("[ex_outside_h] [smallint] NULL,")
        mWorkStr.Append("[type_wheel] [nvarchar](1) NULL,")
        mWorkStr.Append("[no_axles] [nvarchar](1) NULL,")
        mWorkStr.Append("[draft_gear] [tinyint] NULL,")
        mWorkStr.Append("[art_units] [Int] NULL,")
        mWorkStr.Append("[err_code1] [tinyint] NULL,")
        mWorkStr.Append("[err_code2] [tinyint] NULL,")
        mWorkStr.Append("[error_flg] [nvarchar](1) NULL,")
        mWorkStr.Append("[cars] [Int] NULL,")
        mWorkStr.Append("[tons] [Int] NULL,")
        mWorkStr.Append("[total_rev] [Decimal](18, 0) NULL,")
        mWorkStr.Append("[tc_units] [Int] NULL,")
        mWorkStr.Append("[serial_no] [Int] Not NULL,")
        mWorkStr.Append("Constraint [pk_PUWS")
        mWorkStr.Append(mYear)
        mWorkStr.Append("_Masked] ")
        mWorkStr.Append("PRIMARY KEY CLUSTERED ([serial_no] Asc) ")
        mWorkStr.Append("With (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ")
        mWorkStr.Append("ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On) On [PRIMARY] ) On [PRIMARY]")

        mSQLCommand = New SqlCommand
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mWorkStr.ToString
        mSQLCommand.ExecuteNonQuery()

    End Sub

    Function Create_PUWS_Masking_Factors_Table_Command(mYear As String) As String
        Dim mWorkStr As New StringBuilder

        mWorkStr.Append("CREATE TABLE [dbo].[PUWS")
        mWorkStr.Append(mYear)
        mWorkStr.Append("_Masking_Factors]")
        mWorkStr.Append("([Serial_No] [Int] Not NULL,")
        mWorkStr.Append("[Masking_Factor] [Decimal](3, 2) Not NULL,")
        mWorkStr.Append("Constraint [pk_PUWS")
        mWorkStr.Append(mYear)
        mWorkStr.Append("_Masking_Factors] ")
        mWorkStr.Append("PRIMARY KEY CLUSTERED ([Serial_No] Asc) ")
        mWorkStr.Append("With (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ")
        mWorkStr.Append("ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On) On [PRIMARY] ) On [PRIMARY]")

        Create_PUWS_Masking_Factors_Table_Command = mWorkStr.ToString

    End Function

    Function Create_EValues_Function_Command() As String
        Dim mWorkStr As New StringBuilder

        mWorkStr.Append("CREATE Function [dbo].[ufn_EValues] ")
        mWorkStr.Append("(@Year As Char(4), @RRID As TINYINT) ")
        mWorkStr.Append("RETURNS TABLE ")
        mWorkStr.Append("As ")
        mWorkStr.Append("Return ")
        mWorkStr.Append("With PivotData As ")
        mWorkStr.Append("( ")
        mWorkStr.Append("Select XML_Epart + LineA As Node, Code, Value ")
        mWorkStr.Append("From U_EVALUES v INNER JOIN URCS_Controls.dbo.R_ECODES c On v.eCode_id = c.eCode_id ")
        mWorkStr.Append("WHERE Year = @Year And RR_Id = @RRID ")
        mWorkStr.Append(") ")
        mWorkStr.Append("Select Node, [C1],[C2],[C3],[C4],[C5],[C6],[C7],[C8],[C9],[C10],")
        mWorkStr.Append("[C11], [C12], [C13], [C14], [C15], [C16], [C17], [C18], [C19], [C20],")
        mWorkStr.Append("[C21], [C22], [C23], [C24], [C25], [C26], [C27], [C28], [C29] ")
        mWorkStr.Append("From PivotData ")
        mWorkStr.Append("PIVOT (MAX(Value) For Code In ([C1],[C2],[C3],[C4],[C5],[C6],[C7],[C8],")
        mWorkStr.Append("[C9], [C10], [C11], [C12], [C13], [C14], [C15], [C16], [C17], [C18], [C19], ")
        mWorkStr.Append("[C20], [C21], [C22], [C23], [C24], [C25], [C26], [C27], [C28], [C29])) As P")

        Create_EValues_Function_Command = mWorkStr.ToString

    End Function

    Function Create_GenerateEValues_Procedure_Command() As String
        Dim mWorkStr As New StringBuilder

        mWorkStr.Append("CREATE PROCEDURE [dbo].[usp_GenerateEValuesXML] (@Year As Char(4), @Output As XML OUTPUT)")
        mWorkStr.Append(vbCrLf)
        mWorkStr.Append("As BEGIN ")
        mWorkStr.Append(vbCrLf)
        mWorkStr.Append("Declare @comment NVARCHAR(100) ")
        mWorkStr.Append(vbCrLf)
        mWorkStr.Append("Declare @xml_var NVARCHAR(MAX) ")
        mWorkStr.Append(vbCrLf)
        mWorkStr.Append("Set @comment = '<!-- This unit cost file created was on ' + ")
        mWorkStr.Append("Convert(VARCHAR(30),SYSDATETIME(),101) + ' at ' + ")
        mWorkStr.Append("Convert(VARCHAR(30), SYSDATETIME(), 108) + ' -->' ")
        mWorkStr.Append(vbCrLf)
        mWorkStr.Append("Set @xml_var = ")
        mWorkStr.Append("(SELECT SHORT_NAME AS Name, Name as Title, ")
        mWorkStr.Append("(SELECT * FROM ufn_EValues (@Year,l.RR_ID) FOR XML RAW('X'),TYPE) ")
        mWorkStr.Append("From URCS_Controls.dbo.R_Class1_Rail_List l ")
        mWorkStr.Append("WHERE RR_ID <> 0 For XML RAW ('Railroad'), ROOT('UnitCostData')) ")
        mWorkStr.Append("Set @Output = CAST(@comment + REPLACE(REPLACE(@xml_var, 'X Node=")
        mWorkStr.Append(Chr(34) & "',''),'" & Chr(34) & "C1=',' C1=') AS XML) ")
        mWorkStr.Append(vbCrLf)
        mWorkStr.Append("Return 0 ")
        mWorkStr.Append(vbCrLf)
        mWorkStr.Append("End")

        Create_GenerateEValues_Procedure_Command = mWorkStr.ToString

    End Function

    Function Create_RunSubstitutions_Procedure_Command() As String
        Dim mWorkstr As New StringBuilder

        '--=============================================
        '-- Modified By: Michael Sanders, STB
        '-- Mod Date:    2014-12-10 10:21
        '-- Description: Removed reference to fixed database name (URCS)
        '--=========================================

        mWorkstr.Append("CREATE PROCEDURE [dbo].[usp_RunSubstitutions](@Year CHAR(4), @RR_ID TINYINT) ")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("AS BEGIN")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("Declare @NotesText CHAR(55) = 'No substitution - initial value within reasonable range'; ")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("Declare @RegionID TINYINT; ")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("Declare @RR_NAME VARCHAR(5); ")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("Declare @Region_NAME VARCHAR(5); ")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("Select @RegionID = CASE WHEN RR.RR_ID = RG.RR_ID THEN RG.Other_RR_ID ELSE RG.RR_ID END ")
        mWorkstr.Append("From URCS_Controls.dbo.R_Class1_Rail_List RR INNER JOIN ")
        mWorkstr.Append("(SELECT Region_ID, RR_ID, ")
        mWorkstr.Append("(SELECT TOP 1 RR_ID FROM URCS_Controls.dbo.R_Class1_Rail_List ")
        mWorkstr.Append("WHERE RRICC > 900000 And RR_ID > 0 And RR_ID <> S.RR_ID) Other_RR_ID ")
        mWorkstr.Append("From URCS_Controls.dbo.R_Class1_Rail_List S ")
        mWorkstr.Append("WHERE RRICC > 900000 And RR_ID > 0) RG On RR.REGION_ID = RG.REGION_ID ")
        mWorkstr.Append("WHERE RR.RR_ID = @RR_ID ")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("Select TOP 1 @RR_NAME = SHORT_NAME ")
        mWorkstr.Append("From URCS_Controls.dbo.R_Class1_Rail_List Where RR_ID = @RR_ID ")
        mWorkstr.Append("Select TOP 1 @Region_NAME = SHORT_NAME FROM URCS_Controls.dbo.R_Class1_Rail_List ")
        mWorkstr.Append("WHERE RR_ID = @RegionID ")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("-- Run rules 1 through 13 ")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("UPDATE OriginalValues Set OriginalValues.Final_Value = ReplaceValues.Final_Value, ")
        mWorkstr.Append("Notes = 'Updated with ' + ReplaceCodes.eCode + ' from ' + @RR_NAME + ' (RR_ID = ' + ")
        mWorkstr.Append("CAST(ReplaceValues.RR_id As VARCHAR(10)) + ')' ")
        mWorkstr.Append("FROM (U_Substitutions OriginalValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON OriginalValues.eCode_id = OriginalCodes.eCode_ID ")
        mWorkstr.Append("And OriginalCodes.eLine=201) INNER JOIN ")
        mWorkstr.Append("(U_Substitutions ReplaceValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID ")
        mWorkstr.Append("And ReplaceCodes.eLine=202) ")
        mWorkstr.Append("On OriginalValues.RR_id = ReplaceValues.RR_id And OriginalValues.Year = ReplaceValues.Year ")
        mWorkstr.Append("And OriginalCodes.eCode = REPLACE(ReplaceCodes.eCode,'2C','1C') ")
        mWorkstr.Append("WHERE OriginalValues.RR_id = @RR_ID And OriginalValues.Year = @Year ")
        mWorkstr.Append("And OriginalCodes.ePart = 'E1' ")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("-- Run rules 14 through 17 ")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("UPDATE OriginalValues Set OriginalValues.Final_Value = ReplaceValues.Final_Value, ")
        mWorkstr.Append("Notes = 'Updated with ' + ReplaceCodes.eCode + ' from ' + @RR_NAME + ' (RR_ID = ' + ")
        mWorkstr.Append("CAST(ReplaceValues.RR_id As VARCHAR(10)) + ')' ")
        mWorkstr.Append("FROM (U_Substitutions OriginalValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON OriginalValues.eCode_id = OriginalCodes.eCode_ID ")
        mWorkstr.Append("And OriginalCodes.eLine=101) INNER JOIN ")
        mWorkstr.Append("(U_Substitutions ReplaceValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID ")
        mWorkstr.Append("And ReplaceCodes.eLine=102) On OriginalValues.RR_id = ReplaceValues.RR_id ")
        mWorkstr.Append("And OriginalValues.Year = ReplaceValues.Year And ")
        mWorkstr.Append("OriginalCodes.eCode = Replace(ReplaceCodes.eCode,'2C','1C') ")
        mWorkstr.Append("WHERE OriginalValues.RR_id = @RR_ID And OriginalValues.Year = @Year And ")
        mWorkstr.Append("OriginalCodes.ePart = 'E2' AND OriginalCodes.eColumn IN (2,3,4,24)")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("-- Run rules 18 through 38")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("UPDATE OriginalValues Set OriginalValues.Final_Value = Case When ")
        mWorkstr.Append("OriginalValues.Final_Value <= 0 Then ReplaceValues.Final_Value Else ")
        mWorkstr.Append("OriginalValues.Final_Value End, Notes = CASE WHEN OriginalValues.Final_Value <= 0 ")
        mWorkstr.Append("THEN 'Updated with ' + ReplaceCodes.eCode + ' from ' + @Region_NAME + ' ")
        mWorkstr.Append("(RR_ID = ' + CAST(ReplaceValues.RR_id as VARCHAR(10)) + ')' ")
        mWorkstr.Append("Else Case When OriginalValues.Notes Is NULL Then ")
        mWorkstr.Append("@NotesText Else OriginalValues.Notes End End ")
        mWorkstr.Append("FROM (U_Substitutions OriginalValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON OriginalValues.eCode_id = OriginalCodes.eCode_ID) ")
        mWorkstr.Append("INNER Join (U_Substitutions ReplaceValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID) ")
        mWorkstr.Append("On OriginalValues.Year = ReplaceValues.Year And ")
        mWorkstr.Append("OriginalCodes.eCode_id = ReplaceCodes.eCode_id ")
        mWorkstr.Append("WHERE OriginalValues.RR_id = @RR_ID And ReplaceValues.RR_id = @RegionID And ")
        mWorkstr.Append("OriginalValues.Year = @Year And OriginalCodes.eLine BETWEEN 200 And 300 And ")
        mWorkstr.Append("OriginalCodes.eColumn = 13 And OriginalCodes.ePart = 'E1'")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("-- Run rules 39 through 56")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("UPDATE OriginalValues Set OriginalValues.Final_Value = Case ")
        mWorkstr.Append("When OriginalValues.Final_Value < 1 Or OriginalValues.Final_Value > 3 ")
        mWorkstr.Append("Then ReplaceValues.Final_Value Else OriginalValues.Final_Value End, ")
        mWorkstr.Append("Notes = CASE WHEN OriginalValues.Final_Value < 1 Or ")
        mWorkstr.Append("OriginalValues.Final_Value > 3 THEN 'Updated with ' + ReplaceCodes.eCode + ' ")
        mWorkstr.Append("from ' + @Region_NAME + ' (RR_ID = ' + CAST(ReplaceValues.RR_id as VARCHAR(10)) + ')' ")
        mWorkstr.Append("Else Case When OriginalValues.Notes Is NULL Then @NotesText Else OriginalValues.Notes End End ")
        mWorkstr.Append("FROM (U_Substitutions OriginalValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON OriginalValues.eCode_id = OriginalCodes.eCode_ID) ")
        mWorkstr.Append("INNER Join (U_Substitutions ReplaceValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID) ")
        mWorkstr.Append("On OriginalValues.Year = ReplaceValues.Year And OriginalCodes.eCode_id = ReplaceCodes.eCode_id ")
        mWorkstr.Append("WHERE OriginalValues.RR_id = @RR_ID And ReplaceValues.RR_id = @RegionID And ")
        mWorkstr.Append("OriginalValues.Year = @Year And OriginalCodes.eLine BETWEEN 100 And 200 And ")
        mWorkstr.Append("OriginalCodes.eColumn = 3 And OriginalCodes.ePart = 'E2'")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("-- Run rules 58 through 59")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("UPDATE OriginalValues Set OriginalValues.Final_Value = ReplaceValues.Final_Value, ")
        mWorkstr.Append("Notes = 'Updated with ' + ReplaceCodes.eCode + ' from ' + @RR_NAME + ' (RR_ID = ' + ")
        mWorkstr.Append("CAST(ReplaceValues.RR_id As VARCHAR(10)) + ')' ")
        mWorkstr.Append("FROM (U_Substitutions OriginalValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON ")
        mWorkstr.Append("OriginalValues.eCode_id = OriginalCodes.eCode_ID And OriginalCodes.eColumn = 2) ")
        mWorkstr.Append("INNER Join (U_Substitutions ReplaceValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes ReplaceCodes ON ")
        mWorkstr.Append("ReplaceValues.eCode_id = ReplaceCodes.eCode_ID And ReplaceCodes.eColumn = 13 ) ")
        mWorkstr.Append("On OriginalValues.Year = ReplaceValues.Year And OriginalValues.RR_id = ReplaceValues.RR_id ")
        mWorkstr.Append("And OriginalCodes.eLine = ReplaceCodes.eLine ")
        mWorkstr.Append("WHERE OriginalValues.RR_id = @RR_ID And OriginalValues.Year = @Year ")
        mWorkstr.Append("And OriginalCodes.eLine IN (215,216) And OriginalCodes.ePart = 'E1'")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("-- Run rules 61 through 113 (excluding 87)")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("UPDATE OriginalValues Set OriginalValues.Final_Value = ReplaceValues.Final_Value, ")
        mWorkstr.Append("Notes = 'Updated with ' + ReplaceCodes.eCode + ' from ' + @RR_NAME + ' (RR_ID = ' + ")
        mWorkstr.Append("CAST(ReplaceValues.RR_id As VARCHAR(10)) + ')' ")
        mWorkstr.Append("FROM (U_Substitutions OriginalValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON ")
        mWorkstr.Append("OriginalValues.eCode_id = OriginalCodes.eCode_ID And OriginalCodes.eLine In (115,116)) ")
        mWorkstr.Append("INNER Join (U_Substitutions ReplaceValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID ")
        mWorkstr.Append("And ReplaceCodes.eLine = 117) ")
        mWorkstr.Append("On OriginalValues.Year = ReplaceValues.Year And OriginalValues.RR_id = ReplaceValues.RR_id And ")
        mWorkstr.Append("OriginalCodes.eColumn = ReplaceCodes.eColumn And OriginalCodes.EPart = ReplaceCodes.EPart ")
        mWorkstr.Append("WHERE OriginalValues.RR_id = @RR_ID And OriginalValues.Year = @Year And ")
        mWorkstr.Append("OriginalCodes.ePart = 'E2' And ")
        mWorkstr.Append("(OriginalCodes.eColumn = 2 Or OriginalCodes.eColumn BETWEEN 5 And 29)")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("-- Run rule 114")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("UPDATE OriginalValues Set OriginalValues.Final_Value = Case When OriginalValues.Final_Value = 1")
        mWorkstr.Append("Then ReplaceValues.Final_Value Else OriginalValues.Final_Value End, ")
        mWorkstr.Append("Notes = CASE WHEN OriginalValues.Final_Value = 1 ")
        mWorkstr.Append("THEN 'Updated with ' + ReplaceCodes.eCode + ' from ' + @RR_NAME + ' (RR_ID = ' + ")
        mWorkstr.Append("CAST(ReplaceValues.RR_id As VARCHAR(10)) + ')' ")
        mWorkstr.Append("Else Case When OriginalValues.Notes Is NULL Then @NotesText Else OriginalValues.Notes End End ")
        mWorkstr.Append("FROM (U_Substitutions OriginalValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON OriginalValues.eCode_id = OriginalCodes.eCode_ID ")
        mWorkstr.Append("And OriginalCodes.eColumn = 8) INNER JOIN ")
        mWorkstr.Append("(U_Substitutions ReplaceValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes ReplaceCodes ON ReplaceValues.eCode_id = ReplaceCodes.eCode_ID ")
        mWorkstr.Append("And ReplaceCodes.eColumn = 4) ")
        mWorkstr.Append("On OriginalValues.Year = ReplaceValues.Year And OriginalValues.RR_id = ReplaceValues.RR_id ")
        mWorkstr.Append("And OriginalCodes.eLine = ReplaceCodes.eLine ")
        mWorkstr.Append("WHERE OriginalValues.RR_id = @RR_ID And OriginalValues.Year = @Year And ")
        mWorkstr.Append("OriginalCodes.ePart = 'E2' AND OriginalCodes.eLine = 111")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("-- Run rule 115")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("UPDATE OriginalValues Set OriginalValues.Final_Value = Case ")
        mWorkstr.Append("When OriginalValues.Final_Value <> 4162 Then 4162 Else OriginalValues.Final_Value End, ")
        mWorkstr.Append("Notes = CASE WHEN OriginalValues.Final_Value <> 4162 ")
        mWorkstr.Append("THEN 'Updated with value 4162' ELSE @NotesText END ")
        mWorkstr.Append("From U_Substitutions OriginalValues INNER JOIN ")
        mWorkstr.Append("URCS_Controls.dbo.R_ECodes OriginalCodes ON ")
        mWorkstr.Append("OriginalValues.eCode_id = OriginalCodes.eCode_ID And OriginalCodes.eColumn = 23 ")
        mWorkstr.Append("WHERE OriginalValues.RR_id = @RR_ID And OriginalValues.Year = @Year And ")
        mWorkstr.Append("OriginalCodes.ePart = 'E2' AND OriginalCodes.eLine = 111")
        mWorkstr.Append(vbCrLf)
        mWorkstr.Append("End")

        Create_RunSubstitutions_Procedure_Command = mWorkstr.ToString

    End Function

    Function Create_Write_CRPSEG_Legacy_Procedure_Command() As String
        Dim mWorkstr As New StringBuilder
        Dim mLooper As Integer

        mWorkstr.Append("CREATE PROCEDURE [dbo].[usp_WriteCRPSEG_Legacy](")
        mWorkstr.Append("@loopSegment Integer, @SerialNum float, @SegmentNumber float, @ShipmentSize char(2), ")
        mWorkstr.Append("@SegmentType char(2), @CarType integer, @Ownership char(1), @RR_ID int, @RR_Region int, ")
        mWorkstr.Append("@Num_Cars int, @Num_Cars_Expanded int, ")

        For mLooper = 101 To 104
            mWorkstr.Append("@L" & mLooper.ToString & " int, " & vbCrLf)
        Next

        For mLooper = 105 To 111
            mWorkstr.Append("@L" & mLooper.ToString & " float, " & vbCrLf)
        Next

        mWorkstr.Append("@L201 int, " & vbCrLf)

        For mLooper = 202 To 216
            mWorkstr.Append("@L" & mLooper.ToString & " float, " & vbCrLf)
        Next

        mWorkstr.Append("@L217 int, " & vbCrLf & "@L218 int, " & vbCrLf)

        For mLooper = 219 To 248
            mWorkstr.Append("@L" & mLooper.ToString & " float, " & vbCrLf)
        Next

        mWorkstr.Append("@L250 int, " & vbCrLf)

        For mLooper = 251 To 290
            mWorkstr.Append("@L" & mLooper.ToString & " float, " & vbCrLf)
        Next

        For mLooper = 301 To 334
            mWorkstr.Append("@L" & mLooper.ToString & " float, " & vbCrLf)
        Next

        For mLooper = 401 To 499
            mWorkstr.Append("@L" & mLooper.ToString & " float, " & vbCrLf)
        Next

        mWorkstr.Append("@L499A float, @L499B float, @L499C float, @L499D float, " & vbCrLf)

        For mLooper = 501 To 587
            mWorkstr.Append("@L" & mLooper.ToString & " float, " & vbCrLf)
        Next

        For mLooper = 601 To 699
            mWorkstr.Append("@L" & mLooper.ToString & " float, " & vbCrLf)
        Next

        mWorkstr.Append("@L700 float) AS SET NOCOUNT ON; ")
        mWorkstr.Append("If @loopSegment = 1 ")
        mWorkstr.Append("BEGIN ")
        mWorkstr.Append("Insert Into U_CRPSEG_Legacy_1 Values(@SerialNum, @SegmentNumber, @ShipmentSize, ")
        mWorkstr.Append("@SegmentType, @CarType, @Ownership, @RR_ID, @RR_Region, @Num_Cars, @Num_Cars_Expanded, ")

        For mLooper = 101 To 111
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 201 To 248
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 250 To 290
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 301 To 334
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 401 To 499
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        mWorkstr.Append("@L499A, @L499B, @L499C, @L499D, " & vbCrLf)

        For mLooper = 501 To 587
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 601 To 699
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        mWorkstr.Append("@L700) ")
        mWorkstr.Append("End ")
        mWorkstr.Append("Else If @loopSegment = 2 ")
        mWorkstr.Append("BEGIN ")
        mWorkstr.Append("Insert Into U_CRPSEG_Legacy_2 ")
        mWorkstr.Append("Values(@SerialNum, @SegmentNumber, @ShipmentSize, @SegmentType, @CarType, @Ownership, ")
        mWorkstr.Append("@RR_ID, @RR_Region, @Num_Cars, @Num_Cars_Expanded, ")

        For mLooper = 101 To 111
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 201 To 248
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 250 To 290
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 301 To 334
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 401 To 499
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        mWorkstr.Append("@L499A, @L499B, @L499C, @L499D, " & vbCrLf)

        For mLooper = 501 To 587
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 601 To 699
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        mWorkstr.Append("@L700) " & vbCrLf)
        mWorkstr.Append("End ")
        mWorkstr.Append("Else If @loopSegment = 3 ")
        mWorkstr.Append("BEGIN ")
        mWorkstr.Append("Insert Into U_CRPSEG_Legacy ")
        mWorkstr.Append("Values(@SerialNum, @SegmentNumber, @ShipmentSize, @SegmentType, @CarType, @Ownership, ")
        mWorkstr.Append("@RR_ID, @RR_Region, @Num_Cars, @Num_Cars_Expanded, ")

        For mLooper = 101 To 111
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 201 To 248
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 250 To 290
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 301 To 334
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 401 To 499
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        mWorkstr.Append("@L499A, @L499B, @L499C, @L499D, " & vbCrLf)

        For mLooper = 501 To 587
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        For mLooper = 601 To 699
            mWorkstr.Append("@L" & mLooper.ToString & ", " & vbCrLf)
        Next

        mWorkstr.Append("@L700) ")
        mWorkstr.Append("End ")

        Create_Write_CRPSEG_Legacy_Procedure_Command = mWorkstr.ToString

    End Function

    Public Sub Create_URCS_Year_Entry(ByVal mYear As String)
        Dim mSQLCmd As String
        Dim mDatatable As New DataTable
        Dim mCommand As New SqlCommand

        gbl_SQLConnection = New SqlConnection
        gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(Gbl_Controls_Database_Name)
        gbl_SQLConnection.Open()

        mSQLCmd = "Select URCS_Year from URCS_Years WHERE URCS_Year = " & mYear

        Using daAdapter As New SqlDataAdapter(mSQLCmd, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        If mDatatable.Rows.Count = 0 Then
            mCommand.Connection = gbl_SQLConnection
            mCommand.CommandType = CommandType.Text
            mCommand.CommandText = "INSERT INTO URCS_Years VALUES (" & mYear & ")"
            mCommand.ExecuteNonQuery()
        End If

        mDatatable = Nothing

    End Sub

    Public Sub Create_Waybill_Year_Entry(ByVal mYear As String)
        Dim mSQLCmd As String
        Dim mDatatable As New DataTable
        Dim mCommand As New SqlCommand

        gbl_SQLConnection = New SqlConnection
        gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(Gbl_Controls_Database_Name)
        gbl_SQLConnection.Open()

        mSQLCmd = "Select WB_Year from WB_Years WHERE WB_Year = " & mYear

        Using daAdapter As New SqlDataAdapter(mSQLCmd, gbl_SQLConnection)
            daAdapter.Fill(mDatatable)
        End Using

        If mDatatable.Rows.Count = 0 Then
            mCommand.Connection = gbl_SQLConnection
            mCommand.CommandType = CommandType.Text
            mCommand.CommandText = "INSERT INTO WB_Years VALUES (" & mYear & ")"
            mCommand.ExecuteNonQuery()
        End If

        mDatatable = Nothing

    End Sub

    Public Sub Create_Unmasked_Rev_Table(ByVal mDatabase_Name As String, ByVal mTableName As String)
        Dim mCommand As New SqlCommand
        Dim mworkstr As New StringBuilder

        mworkstr.Append("CREATE TABLE [dbo].[" & mTableName & "](")
        mworkstr.Append("[UnMasked_Serial_No] [VarChar](6) Not NULL, ")
        mworkstr.Append("[Total_UnMasked_Rev] [int] NULL, ")
        mworkstr.Append("[ORR_UnMasked_Rev] [int] NULL, ")
        mworkstr.Append("[JRR1_UnMasked_Rev] [int] NULL, ")
        mworkstr.Append("[JRR2_UnMasked_Rev] [int] NULL, ")
        mworkstr.Append("[JRR3_UnMasked_Rev] [int] NULL, ")
        mworkstr.Append("[JRR4_UnMasked_Rev] [int] NULL, ")
        mworkstr.Append("[JRR5_UnMasked_Rev] [int] NULL, ")
        mworkstr.Append("[JRR6_UnMasked_Rev] [int] NULL, ")
        mworkstr.Append("[TRR_UnMasked_Rev] [int] NULL, ")
        mworkstr.Append("[U_Rev_unmasked] [int] NULL,")
        mworkstr.Append("Constraint [PK__" & mTableName & "] PRIMARY KEY CLUSTERED ")
        mworkstr.Append("([UnMasked_Serial_No] ASC) With (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF,")
        mworkstr.Append("IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On) On [PRIMARY]")
        mworkstr.Append(") On [PRIMARY]")

        mCommand.Connection = gbl_SQLConnection
        mCommand.CommandType = CommandType.Text
        mCommand.CommandText = mworkstr.ToString
        mCommand.ExecuteNonQuery()
        Insert_AuditTrail_Record(mDatabase_Name, "Created Waybills->" & mTableName & " Table.")

    End Sub

    Public Sub Create_PUWS_Masked_Rev_Table(ByVal mDatabase_Name As String, ByVal mTableName As String)
        Dim mCommand As New SqlCommand
        Dim mworkstr As New StringBuilder

        mworkstr.Append("CREATE TABLE [dbo].[" & mTableName & "](")
        mworkstr.Append("[PUWS_Serial_No] [int] Not NULL, ")
        mworkstr.Append("[PUWS_Masking_Factor] [decimal(3,2)] NULL, ")
        mworkstr.Append("[PUWS_U_Rev] [int] NULL, ")
        mworkstr.Append("[Total_Rev] [decimal(13,0)] NULL, ")
        mworkstr.Append("Constraint [PK__" & mTableName & "] PRIMARY KEY CLUSTERED ")
        mworkstr.Append("([PUWS_Serial_No] ASC) With (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF,")
        mworkstr.Append("IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On) On [PRIMARY]")
        mworkstr.Append(") On [PRIMARY]")

        mCommand.Connection = gbl_SQLConnection
        mCommand.CommandType = CommandType.Text
        mCommand.CommandText = mworkstr.ToString
        mCommand.ExecuteNonQuery()
        Insert_AuditTrail_Record(mDatabase_Name, "Created Waybills->" & mTableName & " Table.")

    End Sub

    Public Sub Create_Segments_Table(ByVal mDatabaseName As String, ByVal mTableName As String)
        Dim mCommand As New SqlCommand
        Dim mworkstr As New StringBuilder

        mworkstr.Append("CREATE TABLE [dbo].[" & mTableName & "](")
        mworkstr.Append("[Serial_No] [VarChar](6) Not NULL, ")
        mworkstr.Append("[Seg_no] [tinyint] Not NULL, ")
        mworkstr.Append("[Total_Segs] [tinyint] NULL, ")
        mworkstr.Append("[RR_Num] [smallint] NULL, ")
        mworkstr.Append("[RR_Alpha] [nvarchar](5) NULL,")
        mworkstr.Append("[RR_Dist] [int] NULL,")
        mworkstr.Append("[RR_Cntry] [nvarchar](2) NULL,")
        mworkstr.Append("[RR_Rev] [decimal](18, 0) NULL,")
        mworkstr.Append("[RR_VC] [int] NULL,")
        mworkstr.Append("[Seg_Type] [nvarchar](2) NULL,")
        mworkstr.Append("[From_Node] [int] NULL,")
        mworkstr.Append("[To_Node] [int] NULL,")
        mworkstr.Append("[From_Loc] [nvarchar](9) NULL,")
        mworkstr.Append("[From_St] [nvarchar](5) NULL,")
        mworkstr.Append("[To_Loc] [nvarchar](9) NULL,")
        mworkstr.Append("[To_St] [nvarchar](5) NULL,")
        mworkstr.Append("[From_Latitude] [decimal](8,5) NULL,")
        mworkstr.Append("[From_Longitude] [decimal](8,5) NULL,")
        mworkstr.Append("[To_Latitude] [decimal](8,5) NULL,")
        mworkstr.Append("[To_Longitude] [decimal](8,5) NULL,")
        mworkstr.Append("Constraint [pk_" & mTableName & "] PRIMARY KEY CLUSTERED ")
        mworkstr.Append("([Serial_No] ASC, [Seg_no] ASC)With (PAD_INDEX = OFF, ")
        mworkstr.Append("STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On)")
        mworkstr.Append("On [PRIMARY]) On [PRIMARY]")

        mCommand.Connection = gbl_SQLConnection
        mCommand.CommandType = CommandType.Text
        mCommand.CommandText = mworkstr.ToString
        mCommand.ExecuteNonQuery()
        Insert_AuditTrail_Record(mDatabaseName, "Created " & mTableName & " Table.")

    End Sub

    Public Sub Create_Locator_Entry(mYear As String, mTableName As String)
        Dim mSQLCmd As New StringBuilder
        Dim mCommand As New SqlCommand
        Dim mdataTable As New DataTable

        Gbl_Controls_Database_Name = My.Settings.Controls_DB
        OpenSQLConnection(Gbl_Controls_Database_Name)

        mSQLCmd.Append("Select Count(*) as MyCount FROM U_TABLE_LOCATOR WHERE Year = " & mYear & " AND Data_Type = '" & mTableName & "'")

        Using daAdapter As New SqlDataAdapter(mSQLCmd.ToString, gbl_SQLConnection)
            daAdapter.Fill(mdataTable)
        End Using

        mSQLCmd = New StringBuilder

        If mdataTable.Rows(0)("MyCount") = 0 Then
            mSQLCmd.Append("INSERT INTO U_TABLE_LOCATOR (" &
                "Year, Data_Type, Database_Name, Table_Name, Description " &
                ") VALUES (")
            mSQLCmd.Append(mYear & ",")
            mSQLCmd.Append("'" & mTableName & "',")

            Select Case mTableName
                Case "Makewhole_Factors"
                    mSQLCmd.Append("'URCS" & mYear & "','")
                    mSQLCmd.Append("UT_" & mTableName)
                Case "Masked", "Unmasked_Rev", "Segments", "Unmasked_Segments"
                    mSQLCmd.Append("'WAYBILLS','")
                    mSQLCmd.Append("WB" & mYear & "_" & mTableName)
                Case "PUWS_Masked", "PUWS_Masking_Factors"
                    mSQLCmd.Append("'WAYBILLS','PUWS")
                    If mTableName = "PUWS_Masked" Then
                        mSQLCmd.Append(mYear & "_Masked")
                    Else
                        mSQLCmd.Append(mYear & "_Masking_Factors")
                    End If
                Case Else
                    mSQLCmd.Append("'URCS" & mYear & "','")
                    mSQLCmd.Append("U_" & mTableName)
            End Select

            mSQLCmd.Append("','Added " & Date.Now.Date.Date.ToString("MM-dd-yyyy") & "')")
        Else
            mSQLCmd.Append("Update U_TABLE_LOCATOR SET ")
            mSQLCmd.Append("Year = " & mYear & ", ")
            mSQLCmd.Append("Data_Type = '" & mTableName & "',")

            Select Case mTableName
                Case "Makewhole_Factors"
                    mSQLCmd.Append("Database_Name = 'URCS" & mYear & "',")
                    mSQLCmd.Append("Table_Name = 'UT_" & mTableName & "', ")
                Case "Masked", "Unmasked_Rev", "Segments", "Unmasked_Segments"
                    mSQLCmd.Append("Database_Name = 'WAYBILLS',")
                    mSQLCmd.Append("Table_Name = 'WB" & mYear & "_" & mTableName & "', ")
                Case "PUWS_Masked", "PUWS_Masking_Factors"
                    mSQLCmd.Append("Database_Name = 'WAYBILLS',")
                    If mTableName = "PUWS_Masked" Then
                        mSQLCmd.Append("Table_Name = 'PUWS" & mYear & "_Masked" & "', ")
                    Else
                        mSQLCmd.Append("Table_Name = 'PUWS" & mYear & "_Masking_Factors" & "', ")
                    End If
                Case Else
                    mSQLCmd.Append("Database_Name = 'URCS" & mYear & "',")
                    mSQLCmd.Append("Table_Name = 'U_" & mTableName & "', ")
            End Select

            mSQLCmd.Append("Description = '" & Date.Now.Date.Date.ToString("MM-dd-yyyy") & "' ")
            mSQLCmd.Append("WHERE YEAR = " & CStr(mYear) & " ")
            mSQLCmd.Append("AND TABLE_NAME = '" & mTableName & "'")
        End If

        mCommand.Connection = gbl_SQLConnection
        mCommand.CommandType = CommandType.Text
        mCommand.CommandText = mSQLCmd.ToString

        mCommand.ExecuteNonQuery()
        mdataTable = Nothing

    End Sub

    Public Function GetWaybillData(ByVal mYear As Integer, ByVal mUnmasked As Boolean) As DataTable
        ' This function builds a table of either Masked or Unmasked Waybill data in memory
        ' Parameters are the Waybill Year and a True/false condition for requiring Unmasked Revenue data

        Dim cmdCommand As SqlCommand
        Dim daAdapter As New SqlDataAdapter
        Dim dsDataSet As New DataSet

        Try
            cmdCommand = New SqlCommand
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.CommandText = Build_Select_Waybills(mYear, mUnmasked)
            OpenSQLConnection(Gbl_Waybill_Database_Name)
            cmdCommand.Connection = gbl_SQLConnection

            'fill the dataset
            daAdapter = New SqlDataAdapter
            daAdapter.SelectCommand = cmdCommand
            daAdapter.Fill(dsDataSet)

            'Name the dataset Tables
            dsDataSet.Tables(0).TableName = "Waybills"

        Catch SqlEx As SqlException
            Throw New System.Exception("Error when retrieving Waybill table values", SqlEx)
        End Try

        GetWaybillData = dsDataSet.Tables(0)

        'clean up
        cmdCommand.Dispose()
        daAdapter.Dispose()
        dsDataSet.Dispose()

    End Function

    Public Function Build_Select_Waybills(ByVal mYear As Integer, ByVal mUnMasked As Boolean) As String
        ' This function builds a SQL command to retrieve either Masked or unmasked Waybills
        ' Parameters are the Waybill Year and a True/false condition for requiring Unmasked Revenue data
        ' Utilizes the Table_Locator table

        Dim mDuplicateSerials As Boolean

        Build_Select_Waybills = ""

        ' Get the name of the Database and tables in the Waybill database
        gbl_Database_Name = Get_Database_Name_From_SQL(mYear, "MASKED")
        Gbl_Masked_TableName = Get_Table_Name_From_SQL(mYear, "MASKED")
        Gbl_Unmasked_Rev_TableName = Get_Table_Name_From_SQL(mYear, "UnMASKED_REV")

        ' Determine if the masked table is one that is known to contain duplicate serial numbers
        If InStr(Gbl_Masked_TableName, "1984") > 0 Or
            InStr(Gbl_Masked_TableName, "1985") > 0 Or
            InStr(Gbl_Masked_TableName, "1994") > 0 Or
            InStr(Gbl_Masked_TableName, "1995") > 0 Then
            mDuplicateSerials = True
        Else
            mDuplicateSerials = False
        End If

        If mUnMasked = False Then
            Build_Select_Waybills = "SELECT * FROM " & Gbl_Masked_TableName & " ORDER BY serial_no ASC"
        Else
            Select Case mDuplicateSerials
                Case True
                    Build_Select_Waybills = "SELECT " & Gbl_Masked_TableName & ".*, " &
                            Gbl_Unmasked_Rev_TableName & ".* " &
                            "FROM " & Gbl_Masked_TableName & " INNER JOIN " &
                            Gbl_Unmasked_Rev_TableName & " ON " &
                            Gbl_Masked_TableName & ".Serial_No = " &
                            Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_no AND " &
                            Gbl_Masked_TableName & ".WB_Num = " &
                            Gbl_Unmasked_Rev_TableName & ".Unmasked_WB_Num" &
                            " ORDER BY " & Gbl_Masked_TableName & ".Serial_No"
                Case False
                    Build_Select_Waybills = "SELECT " & Gbl_Masked_TableName & ".*, " &
                            Gbl_Unmasked_Rev_TableName & ".* " &
                            "FROM " & Gbl_Masked_TableName & " INNER JOIN " &
                            Gbl_Unmasked_Rev_TableName & " ON " &
                            Gbl_Masked_TableName & ".Serial_No = " &
                            Gbl_Unmasked_Rev_TableName & ".Unmasked_Serial_no" &
                            " ORDER BY " & Gbl_Masked_TableName & ".Serial_No"
            End Select
        End If

    End Function

    Function Get_Railroads_For_Costing() As DataTable
        Dim mSQLstr As String

        Get_Railroads_For_Costing = New DataTable

        ' Get the table name from the Table Locator table
        Gbl_URCS_WAYRRR_TableName = Get_Table_Name_From_SQL("1", "WAYRRR")

        ' Build the SQL statement
        mSQLstr = "select * from " & Gbl_URCS_WAYRRR_TableName

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(Get_Railroads_For_Costing)
        End Using

    End Function

    Function Get_Productivity_Data_Record(ByVal mYear As String,
                                     ByVal mLengthOfHaulStratum As String,
                                     ByVal mCarTypeStratum As String,
                                     ByVal mLadingWeightStratum As String,
                                     ByVal mTonsPerCarStratum As String) As DataTable
        Dim mSQLstr As String

        Get_Productivity_Data_Record = New DataTable

        ' Get data from the Table Locator table
        Gbl_Waybill_Database_Name = Get_Database_Name_From_SQL("1", "Productivity")
        Gbl_Productivity_TableName = Get_Table_Name_From_SQL("1", "Productivity")

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        ' Build the SQL statement
        mSQLstr = "select * from " & Gbl_Productivity_TableName & " WHERE " &
            "Year = " & mYear & " AND " &
            "Length_Of_Haul_Stratum = " & mLengthOfHaulStratum & " AND " &
            "Car_Type_Stratum = " & mCarTypeStratum & " AND " &
            "Lading_Weight_Stratum = " & mLadingWeightStratum & " AND " &
            "Cars_Stratum = " & mTonsPerCarStratum


        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(Get_Productivity_Data_Record)
        End Using

    End Function

    Function Get_Productivity_DataTable(ByVal mYear As String) As DataTable
        Dim mSQLstr As String

        Get_Productivity_DataTable = New DataTable

        ' Get data from the Table Locator table
        Gbl_Waybill_Database_Name = Get_Database_Name_From_SQL("1", "Productivity")
        Gbl_Productivity_TableName = Get_Table_Name_From_SQL("1", "Productivity")

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        ' Build the SQL statement
        mSQLstr = "select * from " & Gbl_Productivity_TableName & " WHERE Year = " & mYear & " ORDER BY " &
            "Year, Length_Of_Haul_Stratum, Car_Type_Stratum, Lading_Weight_Stratum, Cars_Stratum ASC"

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLstr, gbl_SQLConnection)
            daAdapter.Fill(Get_Productivity_DataTable)
        End Using

    End Function

    Function Get_Misc_Report_Line(ByVal mReport As String, ByVal mLine As String) As String
        Dim mSQLStr As String
        Dim mTable As New DataTable

        gbl_Table_Name = Get_Table_Name_From_SQL("1", "R_Misc_Report_Lines")

        OpenSQLConnection(Gbl_Controls_Database_Name)

        ' Build the sql statement
        mSQLStr = "select rpt_Text from " & gbl_Table_Name & " where rpt_name = 'AAR_Index' AND Line_Num = " & mLine

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLStr, gbl_SQLConnection)
            daAdapter.Fill(mTable)
        End Using

        If IsDBNull(mTable.Rows(0)("Rpt_Text")) = True Then
            Get_Misc_Report_Line = ""
        Else
            Get_Misc_Report_Line = mTable.Rows(0)("Rpt_Text")
        End If

    End Function

    Function Column_Exist(ByVal mTableName As String, mColumnName As String) As Boolean
        Dim mTable As New DataTable
        Dim mSQLStr As StringBuilder

        Column_Exist = False
        mSQLStr = New StringBuilder

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        ' Build the sql statement
        mSQLStr.Append("SELECT * FROM Information_Schema.Columns ")
        mSQLStr.Append("WHERE Table_Name = '" & mTableName & "' ")
        mSQLStr.Append("AND Column_Name = '" & mColumnName & "'")

        ' Fill thedatatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLStr.ToString, gbl_SQLConnection)
            daAdapter.Fill(mTable)
        End Using

        If mTable.Rows.Count > 0 Then
            Column_Exist = True
        Else
            Column_Exist = False
        End If

        mTable = Nothing

    End Function

    Public Sub Column_Add(ByVal mTableName As String, mColumnName As String, mType As String)
        Dim mSQLStr As New StringBuilder
        Dim mSQLCommand As SqlCommand

        OpenSQLConnection(Gbl_Waybill_Database_Name)

        ' Build the sql statement
        mSQLStr.Append("ALTER TABLE " & mTableName & " ")
        mSQLStr.Append("ADD " & mColumnName & " " & mType)

        mSQLCommand = New SqlCommand
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mSQLStr.ToString
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.ExecuteNonQuery()

    End Sub

    Public Sub Create_Masked_913_Table(ByVal mDatabaseName As String, ByVal mTableName As String)
        Dim mCommand As New SqlCommand
        Dim mWorkStr As New StringBuilder

        mWorkStr.Append("CREATE TABLE " & mTableName & " (")
        mWorkStr.Append("[Serial_No] [Varchar](6) Not NULL, ")
        mWorkStr.Append("[Wb_Num] [Int] NULL, ")
        mWorkStr.Append("[Wb_Date] [DateTime] NULL, ")
        mWorkStr.Append("[Acct_Period] [nvarchar](6) NULL,")
        mWorkStr.Append("[U_Cars] [Int] NULL,")
        mWorkStr.Append("[U_Car_Init] [nvarchar](4) NULL,")
        mWorkStr.Append("[U_Car_Num] [Int] NULL,")
        mWorkStr.Append("[TOFC_Serv_Code] [nvarchar](3) NULL,")
        mWorkStr.Append("[U_TC_Units] [Int] NULL,")
        mWorkStr.Append("[U_TC_Init] [nvarchar](4) NULL,")
        mWorkStr.Append("[U_TC_Num] [Int] NULL,")
        mWorkStr.Append("[STCC_W49] [nvarchar](7) NULL,")
        mWorkStr.Append("[Bill_Wght] [Int] NULL,")
        mWorkStr.Append("[Act_Wght] [Int] NULL,")
        mWorkStr.Append("[U_Rev] [Int] NULL,")
        mWorkStr.Append("[Tran_Chrg] [Int] NULL,")
        mWorkStr.Append("[Misc_Chrg] [Int] NULL,")
        mWorkStr.Append("[Intra_State_Code] [Int] NULL,")
        mWorkStr.Append("[Transit_Code] [Int] NULL,")
        mWorkStr.Append("[All_Rail_Code] [Int] NULL,")
        mWorkStr.Append("[Type_Move] [Int] NULL,")
        mWorkStr.Append("[Move_Via_Water] [Int] NULL,")
        mWorkStr.Append("[Truck_For_Rail] [Int] NULL,")
        mWorkStr.Append("[Shortline_Miles] [Int] NULL,")
        mWorkStr.Append("[Rebill] [Int] NULL,")
        mWorkStr.Append("[Stratum] [Int] NULL,")
        mWorkStr.Append("[Subsample] [Int] NULL,")
        mWorkStr.Append("[Int_Eq_Flg] [Int] NULL,")
        mWorkStr.Append("[Rate_Flg] [Int] NULL,")
        mWorkStr.Append("[Wb_ID] [nvarchar](25) NULL,")
        mWorkStr.Append("[Report_RR] [smallint] NULL,")
        mWorkStr.Append("[O_FSAC] [Int] NULL,")
        mWorkStr.Append("[ORR] [smallint] NULL,")
        mWorkStr.Append("[JCT1] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR1] [smallint] NULL,")
        mWorkStr.Append("[JCT2] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR2] [smallint] NULL,")
        mWorkStr.Append("[JCT3] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR3] [smallint] NULL,")
        mWorkStr.Append("[JCT4] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR4] [smallint] NULL,")
        mWorkStr.Append("[JCT5] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR5] [smallint] NULL,")
        mWorkStr.Append("[JCT6] [nvarchar](5) NULL,")
        mWorkStr.Append("[JRR6] [smallint] NULL,")
        mWorkStr.Append("[JCT7] [nvarchar](5) NULL,")
        mWorkStr.Append("[TRR] [smallint] NULL,")
        mWorkStr.Append("[T_FSAC] [Int] NULL,")
        mWorkStr.Append("[Pop_Cnt] [Int] NULL,")
        mWorkStr.Append("[Stratum_Cnt] [Int] NULL,")
        mWorkStr.Append("[Report_Period] [tinyint] NULL,")
        mWorkStr.Append("[Car_Own_Mark] [nvarchar](4) NULL,")
        mWorkStr.Append("[Car_Lessee_Mark] [nvarchar](4) NULL,")
        mWorkStr.Append("[Car_Cap] [Int] NULL,")
        mWorkStr.Append("[Nom_Car_Cap] [smallint] NULL,")
        mWorkStr.Append("[Tare] [smallint] NULL,")
        mWorkStr.Append("[Outside_L] [Int] NULL,")
        mWorkStr.Append("[Outside_W] [smallint] NULL,")
        mWorkStr.Append("[Outside_H] [smallint] NULL,")
        mWorkStr.Append("[Ex_Outside_H] [smallint] NULL,")
        mWorkStr.Append("[Type_Wheel] [nvarchar](1) NULL,")
        mWorkStr.Append("[No_Axles] [nvarchar](1) NULL,")
        mWorkStr.Append("[Draft_Gear] [tinyint] NULL,")
        mWorkStr.Append("[Art_Units] [Int] NULL,")
        mWorkStr.Append("[Pool_Code] [Int] NULL,")
        mWorkStr.Append("[Car_Typ] [nvarchar](4) NULL,")
        mWorkStr.Append("[Mech] [nvarchar](4) NULL,")
        mWorkStr.Append("[Lic_St] [nvarchar](2) NULL,")
        mWorkStr.Append("[Mx_Wght_Rail] [smallint] NULL,")
        mWorkStr.Append("[O_SPLC] [nvarchar](6) NULL,")
        mWorkStr.Append("[T_SPLC] [nvarchar](6) NULL,")
        mWorkStr.Append("[STCC] [nvarchar](7) NULL,")
        mWorkStr.Append("[ORR_Alpha] [nvarchar](4) NULL,")
        mWorkStr.Append("[JRR1_Alpha] [nvarchar](4) NULL,")
        mWorkStr.Append("[JRR2_Alpha] [nvarchar](4) NULL,")
        mWorkStr.Append("[JRR3_Alpha] [nvarchar](4) NULL,")
        mWorkStr.Append("[JRR4_Alpha] [nvarchar](4) NULL,")
        mWorkStr.Append("[JRR5_Alpha] [nvarchar](4) NULL,")
        mWorkStr.Append("[JRR6_Alpha] [nvarchar](4) NULL,")
        mWorkStr.Append("[TRR_Alpha] [nvarchar](4) NULL,")
        mWorkStr.Append("[JF] [tinyint] NULL,")
        mWorkStr.Append("[Exp_Factor_Th] [smallint] NULL,")
        mWorkStr.Append("[Error_Flg] [nvarchar](1) NULL,")
        mWorkStr.Append("[STB_Car_typ] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code1] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code2] [tinyint] NULL,")
        mWorkStr.Append("[Err_Code3] [tinyint] NULL,")
        mWorkStr.Append("[Car_Own] [nvarchar](1) NULL,")
        mWorkStr.Append("[TOFC_Unit_Type] [nvarchar](4) NULL,")
        mWorkStr.Append("[Dereg_Date] [DateTime] NULL,")
        mWorkStr.Append("[Dereg_Flg] [Int] NULL,")
        mWorkStr.Append("[Service_Type] [tinyint] NULL,")
        mWorkStr.Append("[Cars] [Int] NULL,")
        mWorkStr.Append("[Bill_Wght_Tons] [Int] NULL,")
        mWorkStr.Append("[Tons] [Int] NULL,")
        mWorkStr.Append("[TC_Units] [Int] NULL,")
        mWorkStr.Append("[Total_Rev] [Decimal](18, 0) NULL,")
        mWorkStr.Append("[ORR_Rev] [Decimal](18, 0) NULL,")
        mWorkStr.Append("[JRR1_Rev] [Decimal](18, 0) NULL,")
        mWorkStr.Append("[JRR2_Rev] [Decimal](18, 0) NULL,")
        mWorkStr.Append("[JRR3_Rev] [Decimal](18, 0) NULL,")
        mWorkStr.Append("[JRR4_Rev] [Decimal](18, 0) NULL,")
        mWorkStr.Append("[JRR5_Rev] [Decimal](18, 0) NULL,")
        mWorkStr.Append("[JRR6_Rev] [Decimal](18, 0) NULL,")
        mWorkStr.Append("[TRR_Rev] [Decimal](18, 0) NULL,")
        mWorkStr.Append("[ORR_Dist] [Int] NULL,")
        mWorkStr.Append("[JRR1_Dist] [Int] NULL,")
        mWorkStr.Append("[JRR2_Dist] [Int] NULL,")
        mWorkStr.Append("[JRR3_Dist] [Int] NULL,")
        mWorkStr.Append("[JRR4_Dist] [Int] NULL,")
        mWorkStr.Append("[JRR5_Dist] [Int] NULL,")
        mWorkStr.Append("[JRR6_Dist] [Int] NULL,")
        mWorkStr.Append("[TRR_Dist] [Int] NULL,")
        mWorkStr.Append("[Total_Dist] [Int] NULL,")
        mWorkStr.Append("[O_ST] [nvarchar](2) NULL,")
        mWorkStr.Append("[JCT1_ST] [nvarchar](2) NULL,")
        mWorkStr.Append("[JCT2_ST] [nvarchar](2) NULL,")
        mWorkStr.Append("[JCT3_ST] [nvarchar](2) NULL,")
        mWorkStr.Append("[JCT4_ST] [nvarchar](2) NULL,")
        mWorkStr.Append("[JCT5_ST] [nvarchar](2) NULL,")
        mWorkStr.Append("[JCT6_ST] [nvarchar](2) NULL,")
        mWorkStr.Append("[JCT7_ST] [nvarchar](2) NULL,")
        mWorkStr.Append("[T_ST] [nvarchar](2) NULL,")
        mWorkStr.Append("[O_BEA] [smallint] NULL,")
        mWorkStr.Append("[T_BEA] [smallint] NULL,")
        mWorkStr.Append("[O_FIPS] [Int] NULL,")
        mWorkStr.Append("[T_FIPS] [Int] NULL,")
        mWorkStr.Append("[O_FA] [tinyint] NULL,")
        mWorkStr.Append("[T_FA] [tinyint] NULL,")
        mWorkStr.Append("[O_FT] [tinyint] NULL,")
        mWorkStr.Append("[T_FT] [tinyint] NULL,")
        mWorkStr.Append("[O_SMSA] [smallint] NULL,")
        mWorkStr.Append("[T_SMSA] [smallint] NULL,")
        mWorkStr.Append("[ONET] [Int] NULL,")
        mWorkStr.Append("[NET1] [Int] NULL,")
        mWorkStr.Append("[NET2] [Int] NULL,")
        mWorkStr.Append("[NET3] [Int] NULL,")
        mWorkStr.Append("[NET4] [Int] NULL,")
        mWorkStr.Append("[NET5] [Int] NULL,")
        mWorkStr.Append("[NET6] [Int] NULL,")
        mWorkStr.Append("[NET7] [Int] NULL,")
        mWorkStr.Append("[TNET] [Int] NULL,")
        mWorkStr.Append("[AL_Flg] [tinyint] NULL,")
        mWorkStr.Append("[AR_Flg] [tinyint] NULL,")
        mWorkStr.Append("[AZ_Flg] [tinyint] NULL,")
        mWorkStr.Append("[CA_Flg] [tinyint] NULL,")
        mWorkStr.Append("[CO_Flg] [tinyint] NULL,")
        mWorkStr.Append("[CT_Flg] [tinyint] NULL,")
        mWorkStr.Append("[DE_Flg] [tinyint] NULL,")
        mWorkStr.Append("[DC_Flg] [tinyint] NULL,")
        mWorkStr.Append("[FL_Flg] [tinyint] NULL,")
        mWorkStr.Append("[GA_Flg] [tinyint] NULL,")
        mWorkStr.Append("[ID_Flg] [tinyint] NULL,")
        mWorkStr.Append("[IL_Flg] [tinyint] NULL,")
        mWorkStr.Append("[IN_Flg] [tinyint] NULL,")
        mWorkStr.Append("[IA_Flg] [tinyint] NULL,")
        mWorkStr.Append("[KS_Flg] [tinyint] NULL,")
        mWorkStr.Append("[KY_Flg] [tinyint] NULL,")
        mWorkStr.Append("[LA_Flg] [tinyint] NULL,")
        mWorkStr.Append("[ME_Flg] [tinyint] NULL,")
        mWorkStr.Append("[MD_Flg] [tinyint] NULL,")
        mWorkStr.Append("[MA_Flg] [tinyint] NULL,")
        mWorkStr.Append("[MI_Flg] [tinyint] NULL,")
        mWorkStr.Append("[MN_Flg] [tinyint] NULL,")
        mWorkStr.Append("[MS_Flg] [tinyint] NULL,")
        mWorkStr.Append("[MO_Flg] [tinyint] NULL,")
        mWorkStr.Append("[MT_Flg] [tinyint] NULL,")
        mWorkStr.Append("[NE_Flg] [tinyint] NULL,")
        mWorkStr.Append("[NV_Flg] [tinyint] NULL,")
        mWorkStr.Append("[NH_Flg] [tinyint] NULL,")
        mWorkStr.Append("[NJ_Flg] [tinyint] NULL,")
        mWorkStr.Append("[NM_Flg] [tinyint] NULL,")
        mWorkStr.Append("[NY_Flg] [tinyint] NULL,")
        mWorkStr.Append("[NC_Flg] [tinyint] NULL,")
        mWorkStr.Append("[ND_Flg] [tinyint] NULL,")
        mWorkStr.Append("[OH_Flg] [tinyint] NULL,")
        mWorkStr.Append("[OK_Flg] [tinyint] NULL,")
        mWorkStr.Append("[OR_Flg] [tinyint] NULL,")
        mWorkStr.Append("[PA_Flg] [tinyint] NULL,")
        mWorkStr.Append("[RI_Flg] [tinyint] NULL,")
        mWorkStr.Append("[SC_Flg] [tinyint] NULL,")
        mWorkStr.Append("[SD_Flg] [tinyint] NULL,")
        mWorkStr.Append("[TN_Flg] [tinyint] NULL,")
        mWorkStr.Append("[TX_Flg] [tinyint] NULL,")
        mWorkStr.Append("[UT_Flg] [tinyint] NULL,")
        mWorkStr.Append("[VT_Flg] [tinyint] NULL,")
        mWorkStr.Append("[VA_Flg] [tinyint] NULL,")
        mWorkStr.Append("[WA_Flg] [tinyint] NULL,")
        mWorkStr.Append("[WV_Flg] [tinyint] NULL,")
        mWorkStr.Append("[WI_Flg] [tinyint] NULL,")
        mWorkStr.Append("[WY_Flg] [tinyint] NULL,")
        mWorkStr.Append("[CD_Flg] [tinyint] NULL,")
        mWorkStr.Append("[MX_Flg] [tinyint] NULL,")
        mWorkStr.Append("[Othr_ST_Flg] [tinyint] NULL,")
        mWorkStr.Append("[Int_Harm_Code] [nvarchar](12) NULL,")
        mWorkStr.Append("[Indus_Class] [nvarchar](4) NULL,")
        mWorkStr.Append("[Inter_SIC] [nvarchar](4) NULL,")
        mWorkStr.Append("[Dom_Canada] [nvarchar](3) NULL,")
        mWorkStr.Append("[CS_54] [nvarchar](2) NULL,")
        mWorkStr.Append("[O_FS_Type] [nvarchar](4) NULL,")
        mWorkStr.Append("[T_FS_Type] [nvarchar](4) NULL,")
        mWorkStr.Append("[O_FS_RateZip] [nvarchar](9) NULL,")
        mWorkStr.Append("[T_FS_RateZip] [nvarchar](9) NULL,")
        mWorkStr.Append("[O_Rate_SPLC] [nvarchar](9) NULL,")
        mWorkStr.Append("[T_Rate_SPLC] [nvarchar](9) NULL,")
        mWorkStr.Append("[O_SwLimit_SPLC] [nvarchar](9) NULL,")
        mWorkStr.Append("[T_SwLimit_SPLC] [nvarchar](9) NULL,")
        mWorkStr.Append("[O_Customs_Flg] [nvarchar](1) NULL,")
        mWorkStr.Append("[T_Customs_Flg] [nvarchar](1) NULL,")
        mWorkStr.Append("[O_Grain_Flg] [nvarchar](1) NULL,")
        mWorkStr.Append("[T_Grain_Flg] [nvarchar](1) NULL,")
        mWorkStr.Append("[O_Ramp_Code] [nvarchar](1) NULL,")
        mWorkStr.Append("[T_Ramp_Code] [nvarchar](1) NULL,")
        mWorkStr.Append("[O_IM_Flg] [nvarchar](1) NULL,")
        mWorkStr.Append("[T_IM_Flg] [nvarchar](1) NULL,")
        mWorkStr.Append("[Transborder_Flg] [nvarchar](1) NULL,")
        mWorkStr.Append("[ORR_Cntry] [nvarchar](2) NULL,")
        mWorkStr.Append("[JRR1_Cntry] [nvarchar](2) NULL,")
        mWorkStr.Append("[JRR2_Cntry] [nvarchar](2) NULL,")
        mWorkStr.Append("[JRR3_Cntry] [nvarchar](2) NULL,")
        mWorkStr.Append("[JRR4_Cntry] [nvarchar](2) NULL,")
        mWorkStr.Append("[JRR5_Cntry] [nvarchar](2) NULL,")
        mWorkStr.Append("[JRR6_Cntry] [nvarchar](2) NULL,")
        mWorkStr.Append("[TRR_Cntry] [nvarchar](2) NULL,")
        mWorkStr.Append("[U_Fuel_SurChrg] [Int] NULL,")
        mWorkStr.Append("[O_Census_Reg] [nvarchar](4) NULL,")
        mWorkStr.Append("[T_Census_Reg] [nvarchar](4) NULL,")
        mWorkStr.Append("[Exp_Factor] [real] NULL,")
        mWorkStr.Append("[Total_VC] [Int] NULL,")
        mWorkStr.Append("[RR1_VC] [Int] NULL,")
        mWorkStr.Append("[RR2_VC] [Int] NULL,")
        mWorkStr.Append("[RR3_VC] [Int] NULL,")
        mWorkStr.Append("[RR4_VC] [Int] NULL,")
        mWorkStr.Append("[RR5_VC] [Int] NULL,")
        mWorkStr.Append("[RR6_VC] [Int] NULL,")
        mWorkStr.Append("[RR7_VC] [Int] NULL,")
        mWorkStr.Append("[RR8_VC] [Int] NULL,")
        mWorkStr.Append("[Tracking_No] [Numeric](13,0) NULL,")
        mWorkStr.Append("Constraint [PK__" & mTableName & "] PRIMARY KEY CLUSTERED ")
        mWorkStr.Append("([Serial_No] Asc) With ")
        mWorkStr.Append("(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ")
        mWorkStr.Append("ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On) On [PRIMARY]) On [PRIMARY]")

        mCommand.Connection = gbl_SQLConnection
        mCommand.CommandType = CommandType.Text
        mCommand.CommandText = mWorkStr.ToString
        mCommand.ExecuteNonQuery()
        'Insert_AuditTrail_Record(mDatabaseName, "Created Waybills->" & mTableName & " Table.")

    End Sub

    Public Sub Create_Unmasked_Segments_Table(ByVal mDatabaseName As String, ByVal mTableName As String)
        Dim mCommand As New SqlCommand
        Dim mworkstr As New StringBuilder

        mworkstr.Append("CREATE TABLE [dbo].[" & mTableName & "](")
        mworkstr.Append("[Serial_No] [VarChar](6) Not NULL,")
        mworkstr.Append("[Seg_no] [tinyint] Not NULL,")
        mworkstr.Append("[RR_Unmasked_Rev] [decimal](18, 0) NULL,")
        mworkstr.Append("Constraint [pk_" & mTableName & "] PRIMARY KEY CLUSTERED ")
        mworkstr.Append("([Serial_No] ASC,[Seg_no] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF,")
        mworkstr.Append("IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On) On [PRIMARY]) ON [PRIMARY]")

        mCommand.Connection = gbl_SQLConnection
        mCommand.CommandType = CommandType.Text
        mCommand.CommandText = mworkstr.ToString
        mCommand.ExecuteNonQuery()
        Insert_AuditTrail_Record(mDatabaseName, "Created " & mTableName & " Table.")

    End Sub

    Public Sub Create_Interim_PUWS_Table(ByVal mDatabaseName As String, ByVal mTableName As String)
        Dim mCommand As New SqlCommand
        Dim mworkstr As New StringBuilder

        mworkstr.Append("CREATE TABLE [dbo].[" & mTableName & "](")
        mworkstr.Append("[PUWS_Serial_No] [VarChar](6) Not NULL,")
        mworkstr.Append("[PUWS_Total_Rev] [decimal](18,0) Not NULL,")
        mworkstr.Append("[PUWS_Masking_Factor] [float] NULL,")
        mworkstr.Append("Constraint [pk_" & mTableName & "] PRIMARY KEY CLUSTERED ")
        mworkstr.Append("([PUWS_Serial_No] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF,")
        mworkstr.Append("IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On) On [PRIMARY]) ON [PRIMARY]")

        mCommand.Connection = gbl_SQLConnection
        mCommand.CommandType = CommandType.Text
        mCommand.CommandText = mworkstr.ToString
        mCommand.ExecuteNonQuery()
        Insert_AuditTrail_Record(mDatabaseName, "Created " & mTableName & " Table.")

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Creates batch pro all miled table. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mTableName">   The table to use. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Sub Create_BatchPro_ALL_Miled_Table(mDatabaseName As String, mTableName As String)
        Dim mworkstr As New StringBuilder
        Dim mSQLconn As New SqlConnection
        Dim mSQLCommand As New SqlCommand

        mworkstr.Append("CREATE TABLE [dbo].[" & mTableName & "](")
        mworkstr.Append("[Serial_No] [VarChar](6) Not NULL,")
        mworkstr.Append("[Seg_no] [tinyint] Not NULL,")
        mworkstr.Append("[Total_Segs] [tinyint] Not NULL,")
        mworkstr.Append("[Seg_Type] [nvarchar](2) Not NULL,")
        mworkstr.Append("[Route_Formula] [tinyint] Not NULL,")
        mworkstr.Append("[RR_Num] [int] Not NULL,")
        mworkstr.Append("[RR_Alpha] [nvarchar](4) Not NULL,")
        mworkstr.Append("[From_GeoCodeType] [tinyint] Not NULL,")
        mworkstr.Append("[From_GeoCodeValue] [nVarchar](5) Not NULL,")
        mworkstr.Append("[To_GeoCodeType] [tinyint] Not NULL,")
        mworkstr.Append("[To_GeoCodeValue] [nVarchar](5) Not NULL,")
        mworkstr.Append("[RR_Dist] [int] Not NULL,")
        mworkstr.Append("[ErrorMsg] [nvarchar](150) Not NULL,")
        mworkstr.Append("Constraint [pk_" & mTableName & "] PRIMARY KEY CLUSTERED ")
        mworkstr.Append("([Serial_No] ASC,[Seg_no] ASC,[Total_Segs] ASC,[Seg_Type] ASC,[Route_Formula] ASC) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF,")
        mworkstr.Append("IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = On, ALLOW_PAGE_LOCKS = On) On [PRIMARY]) ON [PRIMARY]")

        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mworkstr.ToString
        mSQLCommand.ExecuteNonQuery()

        mSQLconn = Nothing
        mworkstr = Nothing

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Loads datatable from SQL. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mSQL_Database">    The SQL database. </param>
    ''' <param name="mSQL_Table_Name">  Name of the SQL table. </param>
    ''' <param name="mData_Table">      The data table. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub Load_Datatable_From_SQL(ByVal mSQL_Database As String, ByVal mSQL_Table_Name As String, ByVal mData_Table As DataTable)
        Dim mSQLStr As String

        OpenSQLConnection(mSQL_Database)
        mSQLStr = "Select * from " & mSQL_Table_Name

        Using daAdapter As New SqlDataAdapter(mSQLStr, gbl_SQLConnection)
            daAdapter.Fill(mData_Table)
        End Using

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Writes a 913 SQL from 445 record to the database. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mDatabase_Name">   Name of the database. </param>
    ''' <param name="mWBTableName">     Table name to use. </param>
    ''' <param name="strInline">        The 445 length text record in which to parse fields from. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub Write913SQLFrom445(ByVal mDatabase_Name As String,
                                  ByVal mWBTableName As String,
                                  ByVal strInline As String)

        Dim mSQLStr As New StringBuilder
        Dim workstr As String
        Dim mSTCC_W49 As String
        Dim mDate As Date
        Dim mDate_Str As String

        Dim mWaybill As New Class_913_Waybill
        Dim rand As New Random()
        Dim msqlCommand As New SqlCommand

        'Get the values needed to do the masking routine

        ' STCC_w49
        workstr = Trim(Mid(strInline, 54, 7))
        Do While Len(workstr) < 7
            workstr = "0" & workstr
        Loop
        mSTCC_W49 = workstr

        ' Date
        workstr = Mid(strInline, 13, 2) & "/" & Mid(strInline, 15, 2) & "/" &
                Mid(strInline, 17, 2)
        workstr = Replace(workstr, " ", "0")
        mDate_Str = workstr
        If IsDate(workstr) Then
            mDate = CDate(workstr)
        Else
            mDate_Str = ""
        End If

        mSQLStr = New StringBuilder

        mSQLStr.Append("INSERT INTO [dbo].[" & mWBTableName & "] (")
        mSQLStr.Append("[serial_no],")
        mSQLStr.Append("[wb_num],")
        mSQLStr.Append("[wb_date],")
        mSQLStr.Append("[acct_period],")
        mSQLStr.Append("[u_cars],")
        mSQLStr.Append("[u_car_init],")
        mSQLStr.Append("[u_car_num],")
        mSQLStr.Append("[tofc_serv_code],")
        mSQLStr.Append("[u_tc_units],")
        mSQLStr.Append("[u_tc_init],")
        mSQLStr.Append("[u_tc_num],")
        mSQLStr.Append("[stcc_w49],")
        mSQLStr.Append("[bill_wght],")
        mSQLStr.Append("[act_wght],")
        mSQLStr.Append("[u_rev],")
        mSQLStr.Append("[tran_chrg],")
        mSQLStr.Append("[misc_chrg],")
        mSQLStr.Append("[intra_state_code],")
        mSQLStr.Append("[transit_code],")
        mSQLStr.Append("[all_rail_code],")
        mSQLStr.Append("[type_move],")
        mSQLStr.Append("[move_via_water],")
        mSQLStr.Append("[truck_for_rail],")
        mSQLStr.Append("[Shortline_Miles],")
        mSQLStr.Append("[rebill],")
        mSQLStr.Append("[stratum],")
        mSQLStr.Append("[subsample],")
        mSQLStr.Append("[transborder_flg],")
        mSQLStr.Append("[rate_flg],")
        mSQLStr.Append("[wb_id],")
        mSQLStr.Append("[report_rr],")
        mSQLStr.Append("[o_fsac],")
        mSQLStr.Append("[orr],")
        mSQLStr.Append("[jct1],")
        mSQLStr.Append("[jrr1],")
        mSQLStr.Append("[jct2],")
        mSQLStr.Append("[jrr2],")
        mSQLStr.Append("[jct3],")
        mSQLStr.Append("[jrr3],")
        mSQLStr.Append("[jct4],")
        mSQLStr.Append("[jrr4],")
        mSQLStr.Append("[jct5],")
        mSQLStr.Append("[jrr5],")
        mSQLStr.Append("[jct6],")
        mSQLStr.Append("[jrr6],")
        mSQLStr.Append("[jct7],")
        mSQLStr.Append("[trr],")
        mSQLStr.Append("[t_fsac],")
        mSQLStr.Append("[pop_cnt],")
        mSQLStr.Append("[stratum_cnt],")
        mSQLStr.Append("[report_period],")
        mSQLStr.Append("[car_own_mark],")
        mSQLStr.Append("[car_lessee_mark],")
        mSQLStr.Append("[car_cap],")
        mSQLStr.Append("[nom_car_cap],")
        mSQLStr.Append("[tare],")
        mSQLStr.Append("[outside_l],")
        mSQLStr.Append("[outside_w],")
        mSQLStr.Append("[outside_h],")
        mSQLStr.Append("[ex_outside_h],")
        mSQLStr.Append("[type_wheel],")
        mSQLStr.Append("[no_axles],")
        mSQLStr.Append("[draft_gear],")
        mSQLStr.Append("[art_units],")
        mSQLStr.Append("[pool_code],")
        mSQLStr.Append("[car_typ],")
        mSQLStr.Append("[mech],")
        mSQLStr.Append("[lic_st],")
        mSQLStr.Append("[mx_wght_rail],")
        mSQLStr.Append("[o_splc],")
        mSQLStr.Append("[t_splc],")
        mSQLStr.Append("[u_fuel_surchrg],")
        mSQLStr.Append("[err_code1],")
        mSQLStr.Append("[err_code2],")
        mSQLStr.Append("[err_code3],")
        mSQLStr.Append("[car_own],")
        mSQLStr.Append("[tofc_unit_type],")
        mSQLStr.Append("[Tracking_No]")

        mSQLStr.Append(") VALUES (")            'do not delete

        ' serial_no
        mSQLStr.Append("'" & Mid(strInline, 1, 6) & "',")
        ' wb_num
        mSQLStr.Append(Mid(strInline, 7, 6) & ",")
        ' wb_date
        mSQLStr.Append("'" & mDate_Str & "', ")
        ' acct_period
        ' This gets the month and year value
        workstr = Trim(Mid(strInline, 19, 4))
        workstr = Replace(workstr, " ", "0")
        Do While Len(workstr) < 4
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' u_cars
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 23, 4)) & ",")
        ' u_car_init
        mSQLStr.Append("'" & Trim(Mid(strInline, 27, 4)) & "', ")
        ' u_car_num
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 31, 6)) & ",")
        ' tofc_serv_code
        mSQLStr.Append("'" & Trim(Mid(strInline, 37, 3)) & "', ")
        ' u_tc_units
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 40, 4)) & ",")
        ' u_tc_init
        mSQLStr.Append("'" & Trim(Mid(strInline, 44, 4)) & "', ")
        ' u_tc_num
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 48, 6)) & ",")
        ' stcc_w49
        mSQLStr.Append("'" & mSTCC_W49 & "', ")
        ' bill_wght
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 61, 9)) & ",")
        ' act_wght
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 70, 9)) & ",")
        ' u_rev
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 79, 9)) & ",")
        ' tran_chrg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 88, 9)) & ",")
        ' misc_chrg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 97, 9)) & ",")
        ' intra_state_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 106, 1)) & ",")
        ' transit_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 107, 1)) & ",")
        ' all_rail_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 108, 1)) & ",")
        ' type_move
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 109, 1)) & ",")
        ' move_via_water
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 110, 1)) & ",")
        ' truck_for_rail
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 111, 1)) & ",")
        ' Shortline_Miles
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 112, 4)) & ",")
        ' rebill
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 116, 1)) & ",")
        ' stratum
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 117, 1)) & ",")
        ' subsample
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 118, 1)) & ",")
        ' tramsborder_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 119, 1)) & ",")
        ' rate_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 120, 1)) & ",")
        ' wb_id
        mSQLStr.Append("'" & Trim(Mid(strInline, 121, 25)) & "', ")
        ' report_rr
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 146, 3)) & ",")
        ' o_fsac
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 149, 5)) & ",")
        ' orr
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 154, 3)) & ",")
        ' jct1
        mSQLStr.Append("'" & Trim(Mid(strInline, 157, 5)) & "', ")
        ' jrr1
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 162, 3)) & ",")
        ' jct2
        mSQLStr.Append("'" & Trim(Mid(strInline, 165, 5)) & "', ")
        ' jrr2
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 170, 3)) & ",")
        ' jct3
        mSQLStr.Append("'" & Trim(Mid(strInline, 173, 5)) & "', ")
        ' jrr3
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 178, 3)) & ",")
        ' jct4
        mSQLStr.Append("'" & Trim(Mid(strInline, 181, 5)) & "', ")
        ' jrr4
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 186, 3)) & ",")
        ' jct5
        mSQLStr.Append("'" & Trim(Mid(strInline, 189, 5)) & "', ")
        ' jrr5
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 194, 3)) & ",")
        ' jct6
        mSQLStr.Append("'" & Trim(Mid(strInline, 197, 5)) & "', ")
        ' jrr6
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 202, 3)) & ",")
        ' jct7
        mSQLStr.Append("'" & Trim(Mid(strInline, 205, 5)) & "', ")
        ' trr
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 226, 3)) & ",")
        ' t_fsac
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 229, 5)) & ",")
        ' pop_cnt
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 234, 8)) & ",")
        ' stratum_cnt
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 242, 6)) & ",")
        ' report_period
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 248, 1)) & ",")
        ' car_own_mark
        mSQLStr.Append("'" & Trim(Mid(strInline, 249, 4)) & "', ")
        ' car_lessee_mark
        mSQLStr.Append("'" & Trim(Mid(strInline, 253, 4)) & "', ")
        ' car_cap
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 257, 5)) & ",")
        ' nom_car_cap
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 262, 3)) & ",")
        ' tare
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 265, 4)) & ",")
        ' outside_l
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 269, 5)) & ",")
        ' outside_w
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 274, 4)) & ",")
        ' outside_h
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 278, 4)) & ",")
        ' ex_outside_h
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 282, 4)) & ",")
        ' type_wheel
        mSQLStr.Append("'" & Trim(Mid(strInline, 286, 1)) & "', ")
        ' no_axles
        mSQLStr.Append("'" & Trim(Mid(strInline, 287, 1)) & "', ")
        ' draft_gear
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 288, 2)) & ",")
        ' art_units
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 290, 1)) & ",")
        ' pool_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 291, 7)) & ",")
        ' car_typ        
        mSQLStr.Append("'" & Trim(Mid(strInline, 298, 4)) & "', ")
        ' mech
        mSQLStr.Append("'" & Trim(Mid(strInline, 302, 4)) & "', ")
        ' lic_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 306, 2)) & "', ")
        ' mx_wght_rail
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 308, 3)) & ",")
        ' o_splc
        workstr = Trim(Mid(strInline, 311, 6))
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' t_splc
        workstr = Trim(Mid(strInline, 317, 6))
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' u_fuel_surchg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 323, 9)) & ",")
        ' err_code1
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 383, 2)) & ",")
        ' err_code2
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 385, 2)) & ",")
        ' err_code3
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 387, 2)) & ",")
        'car_own
        mSQLStr.Append("'" & Trim(Mid(strInline, 417, 1)) & "', ")
        'tofc_unit_type
        mSQLStr.Append("'" & Mid(strInline, 419, 4) & "', ")
        'Tracking_No
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 433, 13)))

        mSQLStr.Append(")")

        OpenSQLConnection(mDatabase_Name)

        msqlCommand = New SqlCommand
        msqlCommand.Connection = gbl_SQLConnection
        msqlCommand.CommandType = CommandType.Text
        msqlCommand.CommandText = mSQLStr.ToString
        msqlCommand.ExecuteNonQuery()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Writes a 913 SQL to the database. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mDatabaseName">    Database name to open connection to. </param>
    ''' <param name="mWBTableName">     Table name to use. </param>
    ''' <param name="strInline">        The 445 length text record in which to parse fields from. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub Write913SQL(ByVal mDatabaseName As String,
                           ByVal mWBTableName As String,
                           ByVal strInline As String)

        Dim mSQLStr As New StringBuilder
        Dim workstr As String
        Dim mRate_Flg As Integer
        Dim mSTCC_W49 As String
        Dim mDate As Date
        Dim mDate_Str As String
        Dim mYear As Integer
        Dim mReport_RR As Integer
        Dim mU_Rev As Single, mNew_U_Rev As Single
        Dim mTotal_Rev As Single, mNew_Total_Rev As Single
        Dim mORR_Rev As Single, mNew_ORR_Rev As Single
        Dim mJRR1_Rev As Single, mNew_JRR1_Rev As Single
        Dim mJRR2_Rev As Single, mNew_JRR2_Rev As Single
        Dim mJRR3_Rev As Single, mNew_JRR3_Rev As Single
        Dim mJRR4_Rev As Single, mNew_JRR4_Rev As Single
        Dim mJRR5_Rev As Single, mNew_JRR5_Rev As Single
        Dim mJRR6_Rev As Single, mNew_JRR6_Rev As Single
        Dim mTRR_Rev As Single, mNew_TRR_Rev As Single
        Dim mJF As Integer, mExp_Factor_Th As Integer
        Dim mWaybill As Class_913_Waybill
        Dim rand As New Random()
        Dim msqlCommand As New SqlCommand

        mWaybill = New Class_913_Waybill

        'Get the values needed to do the masking routine

        ' STCC_w49
        workstr = Trim(Mid(strInline, 58, 7))
        Do While Len(workstr) < 7
            workstr = "0" & workstr
        Loop
        mSTCC_W49 = workstr

        ' Date
        workstr = Mid(strInline, 13, 2) & "/" & Mid(strInline, 15, 2) & "/" &
                Mid(strInline, 17, 4)
        workstr = Replace(workstr, " ", "0")
        mDate_Str = workstr
        mDate = CDate(workstr)

        '' Year
        'mYear = Val(mYearStr)

        ' JF
        mJF = CInt(Mid(strInline, 350, 1))

        ' Rate_Flg
        mRate_Flg = CInt(Mid(strInline, 124, 1))

        ' Exp_Factor_Th
        mExp_Factor_Th = CInt(Mid(strInline, 351, 3))

        ' Report_rr
        mReport_RR = Val(Mid(strInline, 150, 3))

        'Load original data for revenue values to memvars
        ' u_rev
        mU_Rev = ReturnValidNumber(Mid(strInline, 83, 9))
        ' total_rev
        mTotal_Rev = ReturnValidNumber(Mid(strInline, 405, 10))
        ' orr_rev
        mORR_Rev = ReturnValidNumber(Mid(strInline, 415, 10))
        ' jrr1_rev
        mJRR1_Rev = ReturnValidNumber(Mid(strInline, 425, 10))
        ' jrr2_rev
        mJRR2_Rev = ReturnValidNumber(Mid(strInline, 435, 10))
        ' jrr3_rev
        mJRR3_Rev = ReturnValidNumber(Mid(strInline, 445, 10))
        ' jrr4_rev
        mJRR4_Rev = ReturnValidNumber(Mid(strInline, 455, 10))
        ' jrr5_rev
        mJRR5_Rev = ReturnValidNumber(Mid(strInline, 465, 10))
        ' jrr6_rev
        mJRR6_Rev = ReturnValidNumber(Mid(strInline, 475, 10))
        ' trr_rev
        mTRR_Rev = ReturnValidNumber(Mid(strInline, 485, 10))

        'now, load the original values to the new memvars
        mNew_U_Rev = mU_Rev
        mNew_Total_Rev = mTotal_Rev
        mNew_ORR_Rev = mORR_Rev
        mNew_JRR1_Rev = mJRR1_Rev
        mNew_JRR2_Rev = mJRR2_Rev
        mNew_JRR3_Rev = mJRR3_Rev
        mNew_JRR4_Rev = mJRR4_Rev
        mNew_JRR5_Rev = mJRR5_Rev
        mNew_JRR6_Rev = mJRR6_Rev
        mNew_TRR_Rev = mTRR_Rev

        'Check for need to mask. - Do not mask class 1 railroads
        Select Case mReport_RR
            Case 190, 131, 712, 555, 105, 482, 777, 802
                ' Do nothing for records by:
                ' 190 - Conrail
                ' 131 - CNW (part of UP)
                ' 712 - CSXT
                ' 555 - NS
                ' 105 - CP
                ' 482 - SOO (part of CPRS)
                ' 777 - BNSF
                ' 802 - UP
            Case Else
                If (mRate_Flg = 1) And (mU_Rev > 0) Then
                    'mask the u_rev field
                    mNew_U_Rev = CStr(MaskGenericValue(mSTCC_W49, mDate, mYear, Val(mU_Rev)))
                    'calculate the new total rev
                    mNew_Total_Rev = mNew_U_Rev * mExp_Factor_Th
                    'calculate and round the new orr_rev
                    mNew_ORR_Rev = Math.Round(mNew_Total_Rev * (mORR_Rev / mTotal_Rev))
                    'calculate and round the new jrr1_rev
                    mNew_JRR1_Rev = Math.Round(mNew_Total_Rev * (mJRR1_Rev / mTotal_Rev))
                    'calculate and round the new jrr2_rev
                    mNew_JRR2_Rev = Math.Round(mNew_Total_Rev * (mJRR2_Rev / mTotal_Rev))
                    'calculate and round the new jrr3_rev
                    mNew_JRR3_Rev = Math.Round(mNew_Total_Rev * (mJRR3_Rev / mTotal_Rev))
                    'calculate and math.round the new jrr4_rev
                    mNew_JRR4_Rev = Math.Round(mNew_Total_Rev * (mJRR4_Rev / mTotal_Rev))
                    'calculate and math.round the new jrr5_rev
                    mNew_JRR5_Rev = Math.Round(mNew_Total_Rev * (mJRR5_Rev / mTotal_Rev))
                    'calculate and math.round the new jrr6_rev
                    mNew_JRR6_Rev = Math.Round(mNew_Total_Rev * (mJRR6_Rev / mTotal_Rev))
                    'calculate and math.round the new trr_rev
                    mNew_TRR_Rev = Math.Round(mNew_Total_Rev * (mTRR_Rev / mTotal_Rev))
                    If mJF > 0 Then
                        mNew_TRR_Rev = mNew_TRR_Rev + (mNew_Total_Rev -
                            (mNew_ORR_Rev + mNew_JRR1_Rev + mNew_JRR2_Rev +
                            mNew_JRR3_Rev + mNew_JRR4_Rev + mNew_JRR5_Rev +
                            mNew_JRR6_Rev + mNew_TRR_Rev))
                    End If
                End If
        End Select

        mSQLStr = New StringBuilder

        ' If this table does not contain a Tracking_No field, add it to the record
        If Column_Exist(mDatabaseName, mWBTableName, "Tracking_No") = False Then
            Column_Add(mDatabaseName, mWBTableName, "Tracking_No", "BigInt")
        End If

        mSQLStr.Append("INSERT INTO [dbo].[" & mWBTableName & "] (")
        mSQLStr.Append("[serial_no],")
        mSQLStr.Append("[wb_num],")
        mSQLStr.Append("[wb_date],")
        mSQLStr.Append("[acct_period],")
        mSQLStr.Append("[u_cars],")
        mSQLStr.Append("[u_car_init],")
        mSQLStr.Append("[u_car_num],")
        mSQLStr.Append("[tofc_serv_code],")
        mSQLStr.Append("[u_tc_units],")
        mSQLStr.Append("[u_tc_init],")
        mSQLStr.Append("[u_tc_num],")
        mSQLStr.Append("[stcc_w49],")
        mSQLStr.Append("[bill_wght],")
        mSQLStr.Append("[act_wght],")
        mSQLStr.Append("[u_rev],")
        mSQLStr.Append("[tran_chrg],")
        mSQLStr.Append("[misc_chrg],")
        mSQLStr.Append("[intra_state_code],")
        mSQLStr.Append("[transit_code],")
        mSQLStr.Append("[all_rail_code],")
        mSQLStr.Append("[type_move],")
        mSQLStr.Append("[move_via_water],")
        mSQLStr.Append("[truck_for_rail],")
        mSQLStr.Append("[Shortline_Miles],")
        mSQLStr.Append("[rebill],")
        mSQLStr.Append("[stratum],")
        mSQLStr.Append("[subsample],")
        mSQLStr.Append("[int_eq_flg],")
        mSQLStr.Append("[rate_flg],")
        mSQLStr.Append("[wb_id],")
        mSQLStr.Append("[report_rr],")
        mSQLStr.Append("[o_fsac],")
        mSQLStr.Append("[orr],")
        mSQLStr.Append("[jct1],")
        mSQLStr.Append("[jrr1],")
        mSQLStr.Append("[jct2],")
        mSQLStr.Append("[jrr2],")
        mSQLStr.Append("[jct3],")
        mSQLStr.Append("[jrr3],")
        mSQLStr.Append("[jct4],")
        mSQLStr.Append("[jrr4],")
        mSQLStr.Append("[jct5],")
        mSQLStr.Append("[jrr5],")
        mSQLStr.Append("[jct6],")
        mSQLStr.Append("[jrr6],")
        mSQLStr.Append("[jct7],")
        mSQLStr.Append("[trr],")
        mSQLStr.Append("[t_fsac],")
        mSQLStr.Append("[pop_cnt],")
        mSQLStr.Append("[stratum_cnt],")
        mSQLStr.Append("[report_period],")
        mSQLStr.Append("[car_own_mark],")
        mSQLStr.Append("[car_lessee_mark],")
        mSQLStr.Append("[car_cap],")
        mSQLStr.Append("[nom_car_cap],")
        mSQLStr.Append("[tare],")
        mSQLStr.Append("[outside_l],")
        mSQLStr.Append("[outside_w],")
        mSQLStr.Append("[outside_h],")
        mSQLStr.Append("[ex_outside_h],")
        mSQLStr.Append("[type_wheel],")
        mSQLStr.Append("[no_axles],")
        mSQLStr.Append("[draft_gear],")
        mSQLStr.Append("[art_units],")
        mSQLStr.Append("[pool_code],")
        mSQLStr.Append("[car_typ],")
        mSQLStr.Append("[mech],")
        mSQLStr.Append("[lic_st],")
        mSQLStr.Append("[mx_wght_rail],")
        mSQLStr.Append("[o_splc],")
        mSQLStr.Append("[t_splc],")
        mSQLStr.Append("[stcc],")
        mSQLStr.Append("[orr_alpha],")
        mSQLStr.Append("[jrr1_alpha],")
        mSQLStr.Append("[jrr2_alpha],")
        mSQLStr.Append("[jrr3_alpha],")
        mSQLStr.Append("[jrr4_alpha],")
        mSQLStr.Append("[jrr5_alpha],")
        mSQLStr.Append("[jrr6_alpha],")
        mSQLStr.Append("[trr_alpha],")
        mSQLStr.Append("[jf],")
        mSQLStr.Append("[exp_factor_th],")
        mSQLStr.Append("[error_flg],")
        mSQLStr.Append("[stb_car_typ],")
        mSQLStr.Append("[err_code1],")
        mSQLStr.Append("[err_code2],")
        mSQLStr.Append("[err_code3],")
        mSQLStr.Append("[car_own],")
        mSQLStr.Append("[tofc_unit_type],")
        'check to see if we need to insert the deregulation date
        If Val(Mid(strInline, 372, 2)) <> 0 Then
            mSQLStr.Append("[dereg_date],")
        End If
        mSQLStr.Append("[dereg_flg],")
        mSQLStr.Append("[service_type],")
        mSQLStr.Append("[cars],")
        mSQLStr.Append("[bill_wght_tons],")
        mSQLStr.Append("[tons],")
        mSQLStr.Append("[tc_units],")
        mSQLStr.Append("[total_rev],")
        mSQLStr.Append("[orr_rev],")
        mSQLStr.Append("[jrr1_rev],")
        mSQLStr.Append("[jrr2_rev],")
        mSQLStr.Append("[jrr3_rev],")
        mSQLStr.Append("[jrr4_rev],")
        mSQLStr.Append("[jrr5_rev],")
        mSQLStr.Append("[jrr6_rev],")
        mSQLStr.Append("[trr_rev],")
        mSQLStr.Append("[orr_dist],")
        mSQLStr.Append("[jrr1_dist],")
        mSQLStr.Append("[jrr2_dist],")
        mSQLStr.Append("[jrr3_dist],")
        mSQLStr.Append("[jrr4_dist],")
        mSQLStr.Append("[jrr5_dist],")
        mSQLStr.Append("[jrr6_dist],")
        mSQLStr.Append("[trr_dist],")
        mSQLStr.Append("[total_dist],")
        mSQLStr.Append("[o_st],")
        mSQLStr.Append("[jct1_st],")
        mSQLStr.Append("[jct2_st],")
        mSQLStr.Append("[jct3_st],")
        mSQLStr.Append("[jct4_st],")
        mSQLStr.Append("[jct5_st],")
        mSQLStr.Append("[jct6_st],")
        mSQLStr.Append("[jct7_st],")
        mSQLStr.Append("[t_st],")
        mSQLStr.Append("[o_bea],")
        mSQLStr.Append("[t_bea],")
        mSQLStr.Append("[o_fips],")
        mSQLStr.Append("[t_fips],")
        mSQLStr.Append("[o_fa],")
        mSQLStr.Append("[t_fa],")
        mSQLStr.Append("[o_ft],")
        mSQLStr.Append("[t_ft],")
        mSQLStr.Append("[o_smsa],")
        mSQLStr.Append("[t_smsa],")
        mSQLStr.Append("[onet],")
        mSQLStr.Append("[net1],")
        mSQLStr.Append("[net2],")
        mSQLStr.Append("[net3],")
        mSQLStr.Append("[net4],")
        mSQLStr.Append("[net5],")
        mSQLStr.Append("[net6],")
        mSQLStr.Append("[net7],")
        mSQLStr.Append("[tnet],")
        mSQLStr.Append("[al_flg],")
        mSQLStr.Append("[az_flg],")
        mSQLStr.Append("[ar_flg],")
        mSQLStr.Append("[ca_flg],")
        mSQLStr.Append("[co_flg],")
        mSQLStr.Append("[ct_flg],")
        mSQLStr.Append("[de_flg],")
        mSQLStr.Append("[dc_flg],")
        mSQLStr.Append("[fl_flg],")
        mSQLStr.Append("[ga_flg],")
        mSQLStr.Append("[id_flg],")
        mSQLStr.Append("[il_flg],")
        mSQLStr.Append("[in_flg],")
        mSQLStr.Append("[ia_flg],")
        mSQLStr.Append("[ks_flg],")
        mSQLStr.Append("[ky_flg],")
        mSQLStr.Append("[la_flg],")
        mSQLStr.Append("[me_flg],")
        mSQLStr.Append("[md_flg],")
        mSQLStr.Append("[ma_flg],")
        mSQLStr.Append("[mi_flg],")
        mSQLStr.Append("[mn_flg],")
        mSQLStr.Append("[ms_flg],")
        mSQLStr.Append("[mo_flg],")
        mSQLStr.Append("[mt_flg],")
        mSQLStr.Append("[ne_flg],")
        mSQLStr.Append("[nv_flg],")
        mSQLStr.Append("[nh_flg],")
        mSQLStr.Append("[nj_flg],")
        mSQLStr.Append("[nm_flg],")
        mSQLStr.Append("[ny_flg],")
        mSQLStr.Append("[nc_flg],")
        mSQLStr.Append("[nd_flg],")
        mSQLStr.Append("[oh_flg],")
        mSQLStr.Append("[ok_flg],")
        mSQLStr.Append("[or_flg],")
        mSQLStr.Append("[pa_flg],")
        mSQLStr.Append("[ri_flg],")
        mSQLStr.Append("[sc_flg],")
        mSQLStr.Append("[sd_flg],")
        mSQLStr.Append("[tn_flg],")
        mSQLStr.Append("[tx_flg],")
        mSQLStr.Append("[ut_flg],")
        mSQLStr.Append("[vt_flg],")
        mSQLStr.Append("[va_flg],")
        mSQLStr.Append("[wa_flg],")
        mSQLStr.Append("[wv_flg],")
        mSQLStr.Append("[wi_flg],")
        mSQLStr.Append("[wy_flg],")
        mSQLStr.Append("[cd_flg],")
        mSQLStr.Append("[mx_flg],")
        mSQLStr.Append("[othr_st_flg],")
        mSQLStr.Append("[int_harm_code],")
        mSQLStr.Append("[indus_class],")
        mSQLStr.Append("[inter_sic],")
        mSQLStr.Append("[dom_canada],")
        mSQLStr.Append("[cs_54],")
        mSQLStr.Append("[o_fs_type],")
        mSQLStr.Append("[t_fs_type],")
        mSQLStr.Append("[o_fs_ratezip],")
        mSQLStr.Append("[t_fs_ratezip],")
        mSQLStr.Append("[o_rate_splc],")
        mSQLStr.Append("[t_rate_splc],")
        mSQLStr.Append("[o_swlimit_splc],")
        mSQLStr.Append("[t_swlimit_splc],")
        mSQLStr.Append("[o_customs_flg],")
        mSQLStr.Append("[t_customs_flg],")
        mSQLStr.Append("[o_grain_flg],")
        mSQLStr.Append("[t_grain_flg],")
        mSQLStr.Append("[o_ramp_code],")
        mSQLStr.Append("[t_ramp_code],")
        mSQLStr.Append("[o_im_flg],")
        mSQLStr.Append("[t_im_flg],")
        mSQLStr.Append("[transborder_flg],")
        mSQLStr.Append("[orr_cntry],")
        mSQLStr.Append("[jrr1_cntry],")
        mSQLStr.Append("[jrr2_cntry],")
        mSQLStr.Append("[jrr3_cntry],")
        mSQLStr.Append("[jrr4_cntry],")
        mSQLStr.Append("[jrr5_cntry],")
        mSQLStr.Append("[jrr6_cntry],")
        mSQLStr.Append("[trr_cntry],")
        mSQLStr.Append("[u_fuel_surchrg],")
        mSQLStr.Append("[o_census_reg],")
        mSQLStr.Append("[t_census_reg],")
        mSQLStr.Append("[exp_factor],")
        mSQLStr.Append("[total_vc],")
        mSQLStr.Append("[rr1_vc],")
        mSQLStr.Append("[rr2_vc],")
        mSQLStr.Append("[rr3_vc],")
        mSQLStr.Append("[rr4_vc],")
        mSQLStr.Append("[rr5_vc],")
        mSQLStr.Append("[rr6_vc],")
        mSQLStr.Append("[rr7_vc],")
        mSQLStr.Append("[rr8_vc],")
        mSQLStr.Append("[Tracking_No]")

        mSQLStr.Append(") VALUES ('")            'do not delete

        ' serial_no
        mSQLStr.Append(Mid(strInline, 1, 6) & "',")
        ' wb_no
        mSQLStr.Append(Mid(strInline, 7, 6) & ",")
        ' wb_date
        mSQLStr.Append("'" & mDate_Str & "', ")
        ' acct_period
        ' This gets the month and year value
        workstr = Trim(Mid(strInline, 21, 6))
        workstr = Replace(workstr, " ", "0")
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' u_cars
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 27, 4)) & ",")
        ' u_car_init
        mSQLStr.Append("'" & Trim(Mid(strInline, 31, 4)) & "', ")
        ' u_car_num
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 35, 6)) & ",")
        ' tofc_serv_code
        mSQLStr.Append("'" & Trim(Mid(strInline, 41, 1)) & "', ")
        ' u_tc_units
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 44, 4)) & ",")
        ' u_tc_init
        mSQLStr.Append("'" & Trim(Mid(strInline, 48, 4)) & "', ")
        ' u_tc_num
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 52, 6)) & ",")
        ' stcc_w49
        mSQLStr.Append("'" & mSTCC_W49 & "', ")
        ' bill_wght
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 65, 9)) & ",")
        ' act_wght
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 74, 9)) & ",")
        ' u_rev
        mSQLStr.Append(CStr(mNew_U_Rev) & ",")
        ' tran_chrg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 92, 9)) & ",")
        ' misc_chrg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 101, 9)) & ",")
        ' intra_state_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 110, 1)) & ",")
        ' transit_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 111, 1)) & ",")
        ' all_rail_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 112, 1)) & ",")
        ' type_move
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 113, 1)) & ",")
        ' move_via_water
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 114, 1)) & ",")
        ' truck_for_rail
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 115, 1)) & ",")
        ' Shortline_Miles
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 116, 4)) & ",")
        ' rebill
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 120, 1)) & ",")
        ' stratum
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 121, 1)) & ",")
        ' subsample
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 122, 1)) & ",")
        ' int_eq_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 123, 1)) & ",")
        ' rate_flg
        mSQLStr.Append(CStr(mRate_Flg) & ",")
        ' wb_id
        mSQLStr.Append("'" & Trim(Mid(strInline, 125, 25)) & "', ")
        ' report_rr
        mSQLStr.Append(CStr(mReport_RR) & ",")
        ' o_fsac
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 153, 5)) & ",")
        ' orr
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 158, 3)) & ",")
        ' jct1
        mSQLStr.Append("'" & Trim(Mid(strInline, 161, 5)) & "', ")
        ' jrr1
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 166, 3)) & ",")
        ' jct2
        mSQLStr.Append("'" & Trim(Mid(strInline, 169, 5)) & "', ")
        ' jrr2
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 174, 3)) & ",")
        ' jct3
        mSQLStr.Append("'" & Trim(Mid(strInline, 177, 5)) & "', ")
        ' jrr3
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 182, 3)) & ",")
        ' jct4
        mSQLStr.Append("'" & Trim(Mid(strInline, 185, 5)) & "', ")
        ' jrr4
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 190, 3)) & ",")
        ' jct5
        mSQLStr.Append("'" & Trim(Mid(strInline, 193, 5)) & "', ")
        ' jrr5
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 198, 3)) & ",")
        ' jct6
        mSQLStr.Append("'" & Trim(Mid(strInline, 201, 5)) & "', ")
        ' jrr6
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 206, 3)) & ",")
        ' jct7
        mSQLStr.Append("'" & Trim(Mid(strInline, 209, 5)) & "', ")
        ' trr
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 214, 3)) & ",")
        ' t_fsac
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 217, 5)) & ",")
        ' pop_cnt
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 222, 8)) & ",")
        ' stratum_cnt
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 230, 6)) & ",")
        ' report_period
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 236, 1)) & ",")
        ' car_own_mark
        mSQLStr.Append("'" & Trim(Mid(strInline, 237, 4)) & "', ")
        ' car_lessee_mark
        mSQLStr.Append("'" & Trim(Mid(strInline, 241, 4)) & "', ")
        ' car_cap
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 245, 5)) & ",")
        ' nom_car_cap
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 250, 3)) & ",")
        ' tare
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 253, 4)) & ",")
        ' outside_l
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 257, 5)) & ",")
        ' outside_w
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 262, 4)) & ",")
        ' outside_h
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 266, 4)) & ",")
        ' ex_outside_h
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 270, 4)) & ",")
        ' type_wheel
        mSQLStr.Append("'" & Trim(Mid(strInline, 274, 1)) & "', ")
        ' no_axles
        mSQLStr.Append("'" & Trim(Mid(strInline, 275, 1)) & "', ")
        ' draft_gear
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 276, 2)) & ",")
        ' art_units
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 278, 1)) & ",")
        ' pool_code
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 279, 7)) & ",")
        ' car_typ
        mSQLStr.Append("'" & Trim(Mid(strInline, 286, 4)) & "', ")
        ' mech
        mSQLStr.Append("'" & Trim(Mid(strInline, 290, 4)) & "', ")
        ' lic_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 294, 2)) & "', ")
        ' mx_wght_rail
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 296, 3)) & ",")
        ' o_splc
        workstr = Trim(Mid(strInline, 299, 6))
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' t_splc
        workstr = Trim(Mid(strInline, 305, 6))
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")

        ' stcc
        workstr = Trim(Mid(strInline, 311, 7))
        Do While Len(workstr) < 7
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' orr_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 318, 4)) & "', ")
        ' jrr1_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 322, 4)) & "', ")
        ' jrr2_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 326, 4)) & "', ")
        ' jrr3_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 330, 4)) & "', ")
        ' jrr4_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 334, 4)) & "', ")
        ' jrr5_slpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 338, 4)) & "', ")
        ' jrr6_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 342, 4)) & "', ")
        ' trr_alpha
        mSQLStr.Append("'" & Trim(Mid(strInline, 346, 4)) & "', ")
        ' jf
        mSQLStr.Append(CStr(mJF) & ",")
        ' exp_factor_th
        mSQLStr.Append(CStr(mExp_Factor_Th) & ",")
        ' error_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 354, 1)) & "', ")
        ' stb_car_typ
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 355, 2)) & ",")
        ' err_code1
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 357, 2)) & ",")
        ' err_code2
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 359, 2)) & ",")
        ' err_code3
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 361, 2)) & ",")
        'car_own
        mSQLStr.Append("'" & Trim(Mid(strInline, 363, 1)) & "', ")
        'tofc_unit_type
        mSQLStr.Append("'" & Mid(strInline, 364, 4) & "', ")
        'build the date string for the deregulation date
        If Val(Mid(strInline, 372, 2)) <> 0 Then
            workstr = Mid(strInline, 372, 2) & "/" & Mid(strInline, 374, 2) & "/" & Mid(strInline, 368, 4)
            mSQLStr.Append("'" & workstr & "', ")
        End If

        'back to the grind
        ' dereg_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 376, 1)) & ",")
        ' service_type
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 377, 1)) & ",")
        ' cars
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 378, 6)) & ",")
        ' bill_wght_tons
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 384, 7)) & ",")
        'tons
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 391, 8)) & ",")
        ' tc_units
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 399, 6)) & ",")
        ' total_rev
        mSQLStr.Append(CStr(mNew_Total_Rev) & ",")
        ' orr_rev
        mSQLStr.Append(CStr(mNew_ORR_Rev) & ",")
        ' jrr1_rev
        mSQLStr.Append(CStr(mNew_JRR1_Rev) & ",")
        ' jrr2_rev
        mSQLStr.Append(CStr(mNew_JRR2_Rev) & ",")
        ' jrr3_rev
        mSQLStr.Append(CStr(mNew_JRR3_Rev) & ",")
        ' jrr4_rev
        mSQLStr.Append(CStr(mNew_JRR4_Rev) & ",")
        ' jrr5_rev
        mSQLStr.Append(CStr(mNew_JRR5_Rev) & ",")
        ' jrr6_rev
        mSQLStr.Append(CStr(mNew_JRR6_Rev) & ",")
        ' trr_rev
        mSQLStr.Append(CStr(mNew_TRR_Rev) & ",")
        ' orr_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 495, 5)) & ",")
        ' jrr1_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 500, 5)) & ",")
        ' jrr2_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 505, 5)) & ",")
        ' jrr3_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 510, 5)) & ",")
        ' jrr4_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 515, 5)) & ",")
        ' jrr5_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 520, 5)) & ",")
        ' jrr6_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 525, 5)) & ",")
        ' trr_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 530, 5)) & ",")
        ' total_dist
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 535, 5)) & ",")
        ' o_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 540, 2)) & "', ")
        ' jct1_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 542, 2)) & "', ")
        ' jct2_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 544, 2)) & "', ")
        ' jct3_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 546, 2)) & "', ")
        ' jct4_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 548, 2)) & "', ")
        ' jct5_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 550, 2)) & "', ")
        ' jct6_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 552, 2)) & "', ")
        ' jct7_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 554, 2)) & "', ")
        ' t_st
        mSQLStr.Append("'" & Trim(Mid(strInline, 556, 2)) & "', ")
        ' o_bea
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 558, 3)) & ",")
        ' t_bea
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 561, 3)) & ",")
        ' o_fips
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 564, 5)) & ",")
        ' t_fips
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 569, 5)) & ",")
        ' o_fa
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 574, 2)) & ",")
        ' t_fa
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 576, 2)) & ",")
        ' o_ft
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 578, 1)) & ",")
        ' t_ft
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 579, 1)) & ",")
        ' o_smsa
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 580, 4)) & ",")
        ' t_smsa
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 584, 4)) & ",")
        ' onet
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 588, 5)) & ",")
        ' net1
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 593, 5)) & ",")
        ' net2
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 598, 5)) & ",")
        ' net3
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 603, 5)) & ",")
        ' net4
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 608, 5)) & ",")
        ' net5
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 613, 5)) & ",")
        ' net6
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 618, 5)) & ",")
        ' net7
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 623, 5)) & ",")
        ' tnet
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 628, 5)) & ",")
        ' al_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 633, 1)) & ",")
        ' az_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 634, 1)) & ",")
        ' ar_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 635, 1)) & ",")
        ' ca_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 636, 1)) & ",")
        ' co_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 637, 1)) & ",")
        ' ct_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 638, 1)) & ",")
        ' de_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 639, 1)) & ",")
        ' dc_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 640, 1)) & ",")
        ' fl_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 641, 1)) & ",")
        ' ga_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 642, 1)) & ",")
        ' id_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 643, 1)) & ",")
        ' il_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 644, 1)) & ",")
        ' in_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 645, 1)) & ",")
        ' ia_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 646, 1)) & ",")
        ' ks_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 647, 1)) & ",")
        ' ky_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 648, 1)) & ",")
        ' la_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 649, 1)) & ",")
        ' me_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 650, 1)) & ",")
        ' md_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 651, 1)) & ",")
        ' ma_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 652, 1)) & ",")
        ' mi_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 653, 1)) & ",")
        ' mn_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 654, 1)) & ",")
        ' ms_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 655, 1)) & ",")
        ' mo_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 656, 1)) & ",")
        ' mt_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 657, 1)) & ",")
        ' ne_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 658, 1)) & ",")
        ' nv_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 659, 1)) & ",")
        ' nh_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 660, 1)) & ",")
        ' nj_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 661, 1)) & ",")
        ' nm_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 662, 1)) & ",")
        ' ny_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 663, 1)) & ",")
        ' nc_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 664, 1)) & ",")
        ' nd_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 665, 1)) & ",")
        ' oh_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 666, 1)) & ",")
        ' ok_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 667, 1)) & ",")
        ' or_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 668, 1)) & ",")
        ' pa_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 669, 1)) & ",")
        ' ri_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 670, 1)) & ",")
        ' sc_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 671, 1)) & ",")
        ' sd_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 672, 1)) & ",")
        ' tn_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 673, 1)) & ",")
        ' tx_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 674, 1)) & ",")
        'ut_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 675, 1)) & ",")
        ' vt_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 676, 1)) & ",")
        ' va_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 677, 1)) & ",")
        ' wa_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 678, 1)) & ",")
        ' wv_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 679, 1)) & ",")
        ' wi_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 680, 1)) & ",")
        ' wy_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 681, 1)) & ",")
        ' cd_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 682, 1)) & ",")
        ' mx_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 683, 1)) & ",")
        ' othr_st_flg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 684, 1)) & ",")
        ' int_harm_code
        mSQLStr.Append("'" & Trim(Mid(strInline, 685, 12)) & "', ")
        ' indus_class
        mSQLStr.Append("'" & Trim(Mid(strInline, 697, 4)) & "', ")
        ' inter_sic
        mSQLStr.Append("'" & Trim(Mid(strInline, 701, 4)) & "', ")
        ' dom_canada
        mSQLStr.Append("'" & Trim(Mid(strInline, 705, 3)) & "', ")
        ' cs_54
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 708, 2)) & ",")
        ' o_fs_type
        mSQLStr.Append("'" & Trim(Mid(strInline, 710, 4)) & "', ")
        ' t_fs_type
        mSQLStr.Append("'" & Trim(Mid(strInline, 714, 4)) & "', ")
        ' o_fs_ratezip
        mSQLStr.Append("'" & Trim(Mid(strInline, 718, 9)) & "', ")
        ' t_fs_ratezip
        mSQLStr.Append("'" & Trim(Mid(strInline, 727, 9)) & "', ")
        ' o_rate_splc
        mSQLStr.Append("'" & Trim(Mid(strInline, 736, 9)) & "', ")
        ' t_rate_splc
        mSQLStr.Append("'" & Trim(Mid(strInline, 745, 9)) & "', ")
        ' o_swlimit_splc
        mSQLStr.Append("'" & Trim(Mid(strInline, 754, 9)) & "', ")
        ' t_swlimit_splc
        mSQLStr.Append("'" & Trim(Mid(strInline, 763, 9)) & "', ")
        ' o_customs_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 772, 1)) & "', ")
        ' t_customs_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 773, 1)) & "', ")
        ' o_grain_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 774, 1)) & "', ")
        ' t_grain_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 775, 1)) & "', ")
        ' o_ramp_code
        mSQLStr.Append("'" & Trim(Mid(strInline, 776, 1)) & "', ")
        ' t_ramp_code
        mSQLStr.Append("'" & Trim(Mid(strInline, 777, 1)) & "', ")
        ' o_im_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 778, 1)) & "', ")
        ' t_im_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 779, 1)) & "', ")
        ' transborder_flg
        mSQLStr.Append("'" & Trim(Mid(strInline, 780, 1)) & "', ")
        ' orr_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 781, 2)) & "', ")
        ' jrr1_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 783, 2)) & "', ")
        ' jrr2_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 785, 2)) & "', ")
        ' jrr3_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 787, 2)) & "', ")
        ' jrr4_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 789, 2)) & "', ")
        ' jrr5_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 791, 2)) & "', ")
        ' jrr6_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 793, 2)) & "', ")
        ' trr_cntry
        mSQLStr.Append("'" & Trim(Mid(strInline, 795, 2)) & "', ")
        ' u_fuel_surchrg
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 797, 9)) & ",")
        ' there are 13 blanks columns at this point
        ' o_census_reg
        mSQLStr.Append("'" & Trim(Mid(strInline, 819, 4)) & "', ")
        ' t_census_reg
        mSQLStr.Append("'" & Trim(Mid(strInline, 823, 4)) & "', ")
        ' exp_factor
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 827, 7)) & ",")
        ' total_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 834, 8)) & ",")
        ' rr1_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 842, 8)) & ",")
        ' rr2_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 850, 8)) & ",")
        ' rr3_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 858, 8)) & ",")
        ' rr4_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 866, 7)) & ",")
        ' rr5_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 873, 7)) & ",")
        ' rr6_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 880, 7)) & ",")
        ' rr7_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 887, 7)) & ",")
        ' rr8_vc
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 894, 7)) & ",")
        'Tracking_No
        mSQLStr.Append(ReturnValidNumber(Mid(strInline, 901, 13)))

        mSQLStr.Append(")")

        OpenSQLConnection(mDatabaseName)

        msqlCommand = New SqlCommand
        msqlCommand.Connection = gbl_SQLConnection
        msqlCommand.CommandType = CommandType.Text
        msqlCommand.CommandText = mSQLStr.ToString
        msqlCommand.ExecuteNonQuery()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Writes a 445 SQL record to the database. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mDatabaseName">    Database name to open connection to. </param>
    ''' <param name="mWBTableName">     Table name to use. </param>
    ''' <param name="mDataLine">        The data line. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub Write445SQL(ByVal mDatabaseName As String,
                           ByVal mWBTableName As String,
                           mDataLine As String)

        Dim mSQLStr As New StringBuilder
        Dim workstr As String
        Dim mSTCC_W49 As String
        Dim mDate As Date
        Dim mDate_Str As String
        Dim mWaybill As Class_445_Waybill
        Dim rand As New Random()
        Dim msqlCommand As New SqlCommand

        mWaybill = New Class_445_Waybill

        'Get the values needed to do the masking routine

        ' STCC_w49
        workstr = Trim(Mid(mDataLine, 54, 7))
        Do While Len(workstr) < 7
            workstr = "0" & workstr
        Loop
        mSTCC_W49 = workstr

        ' Date
        workstr = Mid(mDataLine, 13, 2) & "/" & Mid(mDataLine, 15, 2) & "/" &
                Mid(mDataLine, 17, 2)
        workstr = Replace(workstr, " ", "0")
        mDate_Str = workstr
        If IsDate(workstr) Then
            mDate = CDate(workstr)
        Else
            mDate_Str = ""
        End If

        mSQLStr = New StringBuilder

        mSQLStr.Append("INSERT INTO " & mWBTableName & " (")
        mSQLStr.Append("[serial_no],")
        mSQLStr.Append("[wb_num],")
        mSQLStr.Append("[wb_date],")
        mSQLStr.Append("[acct_period],")
        mSQLStr.Append("[u_cars],")
        mSQLStr.Append("[u_car_init],")
        mSQLStr.Append("[u_car_num],")
        mSQLStr.Append("[tofc_serv_code],")
        mSQLStr.Append("[u_tc_units],")
        mSQLStr.Append("[u_tc_init],")
        mSQLStr.Append("[u_tc_num],")
        mSQLStr.Append("[stcc_w49],")
        mSQLStr.Append("[bill_wght],")
        mSQLStr.Append("[act_wght],")
        mSQLStr.Append("[u_rev],")
        mSQLStr.Append("[tran_chrg],")
        mSQLStr.Append("[misc_chrg],")
        mSQLStr.Append("[intra_state_code],")
        mSQLStr.Append("[transit_code],")
        mSQLStr.Append("[all_rail_code],")
        mSQLStr.Append("[type_move],")
        mSQLStr.Append("[move_via_water],")
        mSQLStr.Append("[truck_for_rail],")
        mSQLStr.Append("[Shortline_Miles],")
        mSQLStr.Append("[rebill],")
        mSQLStr.Append("[stratum],")
        mSQLStr.Append("[subsample],")
        mSQLStr.Append("[transborder_flg],")
        mSQLStr.Append("[rate_flg],")
        mSQLStr.Append("[wb_id],")
        mSQLStr.Append("[report_rr],")
        mSQLStr.Append("[o_fsac],")
        mSQLStr.Append("[orr],")
        mSQLStr.Append("[jct1],")
        mSQLStr.Append("[jrr1],")
        mSQLStr.Append("[jct2],")
        mSQLStr.Append("[jrr2],")
        mSQLStr.Append("[jct3],")
        mSQLStr.Append("[jrr3],")
        mSQLStr.Append("[jct4],")
        mSQLStr.Append("[jrr4],")
        mSQLStr.Append("[jct5],")
        mSQLStr.Append("[jrr5],")
        mSQLStr.Append("[jct6],")
        mSQLStr.Append("[jrr6],")
        mSQLStr.Append("[jct7],")
        mSQLStr.Append("[jrr7],")
        mSQLStr.Append("[jct8],")
        mSQLStr.Append("[jrr8],")
        mSQLStr.Append("[jct9],")
        mSQLStr.Append("[trr],")
        mSQLStr.Append("[t_fsac],")
        mSQLStr.Append("[pop_cnt],")
        mSQLStr.Append("[stratum_cnt],")
        mSQLStr.Append("[report_period],")
        mSQLStr.Append("[car_own_mark],")
        mSQLStr.Append("[car_lessee_mark],")
        mSQLStr.Append("[car_cap],")
        mSQLStr.Append("[nom_car_cap],")
        mSQLStr.Append("[tare],")
        mSQLStr.Append("[outside_l],")
        mSQLStr.Append("[outside_w],")
        mSQLStr.Append("[outside_h],")
        mSQLStr.Append("[ex_outside_h],")
        mSQLStr.Append("[type_wheel],")
        mSQLStr.Append("[no_axles],")
        mSQLStr.Append("[draft_gear],")
        mSQLStr.Append("[art_units],")
        mSQLStr.Append("[pool_code],")
        mSQLStr.Append("[car_typ],")
        mSQLStr.Append("[mech],")
        mSQLStr.Append("[lic_st],")
        mSQLStr.Append("[mx_wght_rail],")
        mSQLStr.Append("[o_splc],")
        mSQLStr.Append("[t_splc],")
        mSQLStr.Append("[u_fuel_surchg],")
        mSQLStr.Append("[err_code1],")
        mSQLStr.Append("[err_code2],")
        mSQLStr.Append("[err_code3],")
        mSQLStr.Append("[err_code4],")
        mSQLStr.Append("[err_code5],")
        mSQLStr.Append("[err_code6],")
        mSQLStr.Append("[err_code7],")
        mSQLStr.Append("[err_code8],")
        mSQLStr.Append("[err_code9],")
        mSQLStr.Append("[err_code10],")
        mSQLStr.Append("[err_code11],")
        mSQLStr.Append("[err_code12],")
        mSQLStr.Append("[err_code13],")
        mSQLStr.Append("[err_code14],")
        mSQLStr.Append("[err_code15],")
        mSQLStr.Append("[err_code16],")
        mSQLStr.Append("[err_code17],")
        mSQLStr.Append("[car_own],")
        mSQLStr.Append("[tofc_unit_type],")
        mSQLStr.Append("[alk_flg],")
        mSQLStr.Append("[Tracking_No]")

        mSQLStr.Append(") VALUES (")            'do not delete

        ' serial_no
        mSQLStr.Append("'" & Mid(mDataLine, 1, 6) & "',")
        ' wb_no
        mSQLStr.Append(Mid(mDataLine, 7, 6) & ",")
        ' wb_date
        mSQLStr.Append("'" & mDate_Str & "', ")
        ' acct_period
        ' This gets the month and year value
        workstr = Trim(Mid(mDataLine, 19, 4))
        workstr = Replace(workstr, " ", "0")
        Do While Len(workstr) < 4
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' u_cars
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 23, 4)) & ",")
        ' u_car_init
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 27, 4)) & "', ")
        ' u_car_num
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 31, 6)) & ",")
        ' tofc_serv_code
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 37, 3)) & "', ")
        ' u_tc_units
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 40, 4)) & ",")
        ' u_tc_init
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 44, 4)) & "', ")
        ' u_tc_num
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 48, 6)) & ",")
        ' stcc_w49
        mSQLStr.Append("'" & mSTCC_W49 & "', ")
        ' bill_wght
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 61, 9)) & ",")
        ' act_wght
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 70, 9)) & ",")
        ' u_rev
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 79, 9)) & ",")
        ' tran_chrg
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 88, 9)) & ",")
        ' misc_chrg
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 97, 9)) & ",")
        ' intra_state_code
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 106, 1)) & ",")
        ' transit_code
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 107, 1)) & ",")
        ' all_rail_code
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 108, 1)) & ",")
        ' type_move
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 109, 1)) & ",")
        ' move_via_water
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 110, 1)) & ",")
        ' truck_for_rail
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 111, 1)) & ",")
        ' Shortline_Miles
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 112, 4)) & ",")
        ' rebill
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 116, 1)) & ",")
        ' stratum
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 117, 1)) & ",")
        ' subsample
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 118, 1)) & ",")
        ' tramsborder_flg
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 119, 1)) & ",")
        ' rate_flg
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 120, 1)) & ",")
        ' wb_id
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 121, 25)) & "', ")
        ' report_rr
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 146, 3)) & ",")
        ' o_fsac
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 149, 5)) & ",")
        ' orr
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 154, 3)) & ",")
        ' jct1
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 157, 5)) & "', ")
        ' jrr1
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 162, 3)) & ",")
        ' jct2
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 165, 5)) & "', ")
        ' jrr2
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 170, 3)) & ",")
        ' jct3
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 173, 5)) & "', ")
        ' jrr3
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 178, 3)) & ",")
        ' jct4
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 181, 5)) & "', ")
        ' jrr4
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 186, 3)) & ",")
        ' jct5
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 189, 5)) & "', ")
        ' jrr5
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 194, 3)) & ",")
        ' jct6
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 197, 5)) & "', ")
        ' jrr6
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 202, 3)) & ",")
        ' jct7
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 205, 5)) & "', ")
        ' jrr7
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 210, 3)) & ",")
        ' jct8
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 213, 5)) & "', ")
        ' jrr8
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 218, 3)) & ",")
        ' jct9
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 221, 5)) & "', ")
        ' trr
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 226, 3)) & ",")
        ' t_fsac
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 229, 5)) & ",")
        ' pop_cnt
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 234, 8)) & ",")
        ' stratum_cnt
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 242, 6)) & ",")
        ' report_period
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 248, 1)) & ",")
        ' car_own_mark
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 249, 4)) & "', ")
        ' car_lessee_mark
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 253, 4)) & "', ")
        ' car_cap
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 257, 5)) & ",")
        ' nom_car_cap
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 262, 3)) & ",")
        ' tare
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 265, 4)) & ",")
        ' outside_l
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 269, 5)) & ",")
        ' outside_w
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 274, 4)) & ",")
        ' outside_h
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 278, 4)) & ",")
        ' ex_outside_h
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 282, 4)) & ",")
        ' type_wheel
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 286, 1)) & "', ")
        ' no_axles
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 287, 1)) & "', ")
        ' draft_gear
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 288, 2)) & ",")
        ' art_units
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 290, 1)) & ",")
        ' pool_code
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 291, 7)) & ",")
        ' car_typ
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 298, 4)) & "', ")
        ' mech
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 302, 4)) & "', ")
        ' lic_st
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 306, 2)) & "', ")
        ' mx_wght_rail
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 308, 3)) & ",")
        ' o_splc
        workstr = Trim(Mid(mDataLine, 311, 6))
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' t_splc
        workstr = Trim(Mid(mDataLine, 317, 6))
        Do While Len(workstr) < 6
            workstr = "0" & workstr
        Loop
        mSQLStr.Append("'" & workstr & "', ")
        ' u_fuel_surchg
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 323, 9)) & ",")
        ' err_code1
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 383, 2)) & ",")
        ' err_code2
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 385, 2)) & ",")
        ' err_code3
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 387, 2)) & ",")
        ' err_code4
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 389, 2)) & ",")
        ' err_code5
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 391, 2)) & ",")
        ' err_code6
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 393, 2)) & ",")
        ' err_code7
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 395, 2)) & ",")
        ' err_code8
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 397, 2)) & ",")
        ' err_code9
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 399, 2)) & ",")
        ' err_code10
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 401, 2)) & ",")
        ' err_code11
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 403, 2)) & ",")
        ' err_code12
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 405, 2)) & ",")
        ' err_code13
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 407, 2)) & ",")
        ' err_code14
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 409, 2)) & ",")
        ' err_code15
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 411, 2)) & ",")
        ' err_code16
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 413, 2)) & ",")
        ' err_code17
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 415, 2)) & ",")
        'car_own
        mSQLStr.Append("'" & Trim(Mid(mDataLine, 417, 1)) & "', ")
        'tofc_unit_type
        mSQLStr.Append("'" & Mid(mDataLine, 419, 4) & "', ")
        'alk_flg
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 432, 1)) & ",")
        'Tracking_No
        mSQLStr.Append(ReturnValidNumber(Mid(mDataLine, 433, 13)))

        mSQLStr.Append(")")

        OpenSQLConnection(mDatabaseName)

        msqlCommand = New SqlCommand
        msqlCommand.Connection = gbl_SQLConnection
        msqlCommand.CommandType = CommandType.Text
        msqlCommand.CommandText = mSQLStr.ToString
        msqlCommand.ExecuteNonQuery()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Updates the alpha field. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mDatabaseName">            Database name to open connection to. </param>
    ''' <param name="mTableName">               The table to use. </param>
    ''' <param name="mSerial_No_Field_Name">    Name of the serial no field. </param>
    ''' <param name="mSerial_No">               The serial no. </param>
    ''' <param name="mFieldName">               Name of the field. </param>
    ''' <param name="mValue">                   The value to update in the trans table. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub Update_Alpha_Field(ByVal mDatabaseName As String,
                            ByVal mTableName As String,
                            ByVal mSerial_No_Field_Name As String,
                            ByVal mSerial_No As String,
                            ByVal mFieldName As String,
                            ByVal mValue As String)

        Dim mSQLStr As New StringBuilder
        Dim mSQLCommand As SqlCommand

        OpenSQLConnection(mDatabaseName)

        ' Build the sql statement
        mSQLStr.Append("UPDATE " & mTableName & " ")
        mSQLStr.Append("SET " & mFieldName & " = '" & mValue.ToString & "' WHERE " & mSerial_No_Field_Name & " = '" & mSerial_No & "'")

        mSQLCommand = New SqlCommand
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mSQLStr.ToString
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.ExecuteNonQuery()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Updates the numeric field. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mDatabaseName">            Database name to open connection to. </param>
    ''' <param name="mTableName">               The table to use. </param>
    ''' <param name="mSerial_No_Field_Name">    Name of the serial no field. </param>
    ''' <param name="mSerial_No">               The serial no. </param>
    ''' <param name="mFieldName">               Name of the field. </param>
    ''' <param name="mValue">                   The value to update in the trans table. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub Update_Numeric_Field(ByVal mDatabaseName As String,
                            ByVal mTableName As String,
                            ByVal mSerial_No_Field_Name As String,
                            ByVal mSerial_No As String,
                            ByVal mFieldName As String,
                            ByVal mValue As String)

        Dim mSQLStr As New StringBuilder
        Dim mSQLCommand As SqlCommand

        OpenSQLConnection(mDatabaseName)

        ' Build the sql statement
        mSQLStr.Append("UPDATE " & mTableName & " ")
        mSQLStr.Append("SET " & mFieldName & " = " & mValue.ToString & " WHERE " & mSerial_No_Field_Name & " = '" & mSerial_No & "'")

        mSQLCommand = New SqlCommand
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mSQLStr.ToString
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.ExecuteNonQuery()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Gets a fields value. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mDatabaseName">        Database name to open connection to. </param>
    ''' <param name="mTableName">           Name of the table to count records. </param>
    ''' <param name="mFieldName">           Name of the field. </param>
    ''' <param name="mFieldToSearchName">   Name of the field to search. </param>
    ''' <param name="mValueToSearchFor">    The value to search for. </param>
    '''
    ''' <returns>   The field value. </returns>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Function Get_Field_Value(ByVal mDatabaseName As String,
                             ByVal mTableName As String,
                             ByVal mFieldName As String,
                             ByVal mFieldToSearchName As String,
                             ByVal mValueToSearchFor As String) As String

        Dim mDataTable As New DataTable
        Dim mStrSQL As String

        OpenSQLConnection(mDatabaseName)

        mStrSQL = "SELECT " & mFieldName & " FROM " & mTableName & " WHERE " & mFieldToSearchName & " = '" & mValueToSearchFor & "'"

        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        If mDataTable.Rows.Count > 0 Then
            Get_Field_Value = mDataTable.Rows(0)(0)
        Else
            Get_Field_Value = ""
        End If

        mDataTable = Nothing

    End Function

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Inserts all miled record. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mDatabase">            The name of the database to connect to. </param>
    ''' <param name="mTableName">           The table to use. </param>
    ''' <param name="mSerial_No">           The serial no. </param>
    ''' <param name="mSeg_No">              The segment no. </param>
    ''' <param name="mTotal_Segs">          The total segments. </param>
    ''' <param name="mSeg_Type">            Type of the segment. </param>
    ''' <param name="mRoute_Formula">       The route formula. </param>
    ''' <param name="mRR_Num">              The rr number. </param>
    ''' <param name="mRR_Alpha">            The rr alpha. </param>
    ''' <param name="mFrom_GeoCodeType">    Type of from geo code. </param>
    ''' <param name="mFrom_GeoCodeValue">   from geo code value. </param>
    ''' <param name="mTo_GeoCodeType">      Type of to geo code. </param>
    ''' <param name="mTo_GeoCodeValue">     to geo code value. </param>
    ''' <param name="mRR_Dist">             The rr distance. </param>
    ''' <param name="mErrorMsg">            Message describing the error. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub Insert_All_Miled_Record(mDatabase As String,
                                       mTableName As String,
                                       mSerial_No As String,
                                       mSeg_No As String,
                                       mTotal_Segs As String,
                                       mSeg_Type As String,
                                       mRoute_Formula As String,
                                       mRR_Num As String,
                                       mRR_Alpha As String,
                                       mFrom_GeoCodeType As String,
                                       mFrom_GeoCodeValue As String,
                                       mTo_GeoCodeType As String,
                                       mTo_GeoCodeValue As String,
                                       mRR_Dist As String,
                                       mErrorMsg As String)

        Dim mSQLCommand As SqlCommand
        Dim mSQLString As New StringBuilder

        mSQLString.Append("INSERT INTO " & mTableName & " (")
        mSQLString.Append("Serial_No,")
        mSQLString.Append("Seg_No,")
        mSQLString.Append("Total_Segs,")
        mSQLString.Append("Seg_Type,")
        mSQLString.Append("Route_Formula,")
        mSQLString.Append("RR_Num,")
        mSQLString.Append("RR_Alpha,")
        mSQLString.Append("From_GeoCodeType,")
        mSQLString.Append("From_GeoCodeValue,")
        mSQLString.Append("To_GeoCodeType,")
        mSQLString.Append("To_GeoCodeValue,")
        mSQLString.Append("RR_Dist,")
        mSQLString.Append("ErrorMsg) VALUES (")
        mSQLString.Append("'" & mSerial_No & "',")
        mSQLString.Append(mSeg_No & ",")
        mSQLString.Append(mTotal_Segs & ",")
        mSQLString.Append("'" & mSeg_Type & "',")
        mSQLString.Append(mRoute_Formula & ",")
        mSQLString.Append(mRR_Num & ",")
        mSQLString.Append("'" & mRR_Alpha & "',")
        mSQLString.Append(mFrom_GeoCodeType & ",")
        mSQLString.Append("'" & mFrom_GeoCodeValue & "',")
        mSQLString.Append(mTo_GeoCodeType & ",")
        mSQLString.Append("'" & mTo_GeoCodeValue & "',")
        mSQLString.Append(mRR_Dist & ",'")
        mSQLString.Append(mErrorMsg & "')")

        mSQLCommand = New SqlCommand
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mSQLString.ToString

        ' Open the SQL connection
        OpenSQLConnection(mDatabase)

        ' execute the command
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.ExecuteNonQuery()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Updates the segment field. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mDatabaseName">    Database name to open connection to. </param>
    ''' <param name="mTableName">       The table to use. </param>
    ''' <param name="mSerial_No">       The serial no. </param>
    ''' <param name="mSegment_No">      The segment no. </param>
    ''' <param name="mFieldName">       Name of the field. </param>
    ''' <param name="mValue">           The value to update in the trans table. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub Update_Segment_Field(ByVal mDatabaseName As String,
                            ByVal mTableName As String,
                            ByVal mSerial_No As String,
                            ByVal mSegment_No As String,
                            ByVal mFieldName As String,
                            ByVal mValue As String)

        Dim mSQLStr As New StringBuilder
        Dim mSQLCommand As SqlCommand

        OpenSQLConnection(mDatabaseName)

        ' Build the sql statement
        mSQLStr.Append("UPDATE " & mTableName & " ")
        mSQLStr.Append("SET " & mFieldName & " = " & mValue.ToString & " WHERE Serial_No = '" & mSerial_No & "' And Seg_No = " & mSegment_No)

        mSQLCommand = New SqlCommand
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mSQLStr.ToString
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.ExecuteNonQuery()

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Column exist. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mDatabaseName">    Database name to open connection to. </param>
    ''' <param name="mTableName">       The table to use. </param>
    ''' <param name="mColumnName">      Name of the column. </param>
    '''
    ''' <returns>   True if it succeeds, false if it fails. </returns>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Function Column_Exist(ByVal mDatabaseName As String, ByVal mTableName As String, mColumnName As String) As Boolean
        Dim mTable As New DataTable
        Dim mSQLStr As StringBuilder

        Column_Exist = False
        mSQLStr = New StringBuilder

        OpenSQLConnection(mDatabaseName)

        ' Build the sql statement
        mSQLStr.Append("SELECT * FROM Information_Schema.Columns ")
        mSQLStr.Append("WHERE Table_Name = '" & mTableName & "' ")
        mSQLStr.Append("AND Column_Name = '" & mColumnName & "'")

        ' Fill thedatatable from SQL
        Using daAdapter As New SqlDataAdapter(mSQLStr.ToString, gbl_SQLConnection)
            daAdapter.Fill(mTable)
        End Using

        If mTable.Rows.Count > 0 Then
            Column_Exist = True
        Else
            Column_Exist = False
        End If

        mTable = Nothing

    End Function

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Column add. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/7/2020. </remarks>
    '''
    ''' <param name="mDatabaseName">    Database name to open connection to. </param>
    ''' <param name="mTableName">       The table to use. </param>
    ''' <param name="mColumnName">      Name of the column. </param>
    ''' <param name="mType">            The type. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub Column_Add(ByVal mDatabaseName As String, ByVal mTableName As String, mColumnName As String, mType As String)
        Dim mSQLStr As New StringBuilder
        Dim mSQLCommand As SqlCommand

        ' Build the sql statement
        mSQLStr.Append("ALTER TABLE " & mTableName & " ")
        mSQLStr.Append("ADD " & mColumnName & " " & mType)

        OpenSQLConnection(mDatabaseName)

        mSQLCommand = New SqlCommand
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.CommandText = mSQLStr.ToString
        mSQLCommand.Connection = gbl_SQLConnection
        mSQLCommand.ExecuteNonQuery()

    End Sub

End Module
