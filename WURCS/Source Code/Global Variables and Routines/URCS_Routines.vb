Imports System.Data.SqlClient
Imports System.Text
Module URCS_Routines

    Function GetExcelConnection(ByVal Path As String,
    Optional ByVal Headers As Boolean = True) As ADODB.Connection
        Dim strConn As String
        Dim objConn As ADODB.Connection
        objConn = New ADODB.Connection
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" &
                  "Data Source=" & Path & ";" &
                  "Extended Properties=""Excel 12.0 XML;HDR=" &
                  IIf(Headers, "Yes", "No") & """"
        objConn.Open(strConn)
        GetExcelConnection = objConn
    End Function

    Function Array1DFindFirst(
         ByVal avValues As Object,
         ByVal vFindItem As Object,
         Optional ByVal lDefault As Long = 0) As Long

        'Purpose   :    Finds the position of the first matching item in an 1d array.
        'Inputs    :    avValues                    The array to evaluate.
        '               vFindItem                   The value to look for in the array.
        '               [lDefault]                  The value to return if vFindItem is not found
        'Outputs   :    The position of the item with the array
        '               OR 0/lDefault if the item was not found.
        'Notes     :    Make the module Option Compare Text for case insensative searches


        Dim lThisRow As Long, bFound As Boolean

        If IsArray(avValues) Then
            On Error Resume Next
            For lThisRow = LBound(avValues) To UBound(avValues)
                If vFindItem = avValues(lThisRow) Then
                    bFound = True
                    Array1DFindFirst = lThisRow
                    Exit For
                End If
            Next
        End If

        If bFound = False Then
            Array1DFindFirst = lDefault
        End If
        On Error GoTo 0

    End Function

    Function Right_Justify(
        ByVal rstField As Object,
        ByVal FieldLength As Integer) As String

        Dim mString As String
        Dim mPad As Long
        Dim y As Integer
        Dim mformatstr As String

        mformatstr = ""

        If String.IsNullOrEmpty(rstField) Then
            Right_Justify = Space(FieldLength)
        Else
            If IsNumeric(rstField) Then
                mString = ""
                mformatstr = Trim(CStr(rstField))
                For y = 1 To FieldLength - Len(mformatstr)
                    mString = mString + " "
                Next
                Right_Justify = mString & mformatstr
            Else
                mString = rstField
                mPad = Len(mString)
                Right_Justify = Space(FieldLength - mPad) & CStr(mString)
            End If
        End If

    End Function

    Function STB_Car_Type(ByVal mValue As Integer) As Integer

        Select Case mValue
            Case 36
                STB_Car_Type = 1    'unequipped box car
            Case 37
                STB_Car_Type = 2    '50 ft box car
            Case 38
                STB_Car_Type = 3    'equipped box car
            Case 39
                STB_Car_Type = 4    'unequipped general service gondola
            Case 40
                STB_Car_Type = 5    'equipped general service gondola
            Case 41
                STB_Car_Type = 6    'covered hopper
            Case 42
                STB_Car_Type = 7    'general service covered hopper
            Case 43
                STB_Car_Type = 8    'open, special service hopper
            Case 44
                STB_Car_Type = 9    'Mechanical Refrigerator
            Case 45
                STB_Car_Type = 10   'non-mechanical refrigerator
            Case 46
                STB_Car_Type = 11   'TOFC flat
            Case 47
                STB_Car_Type = 12   'multi level flat
            Case 48
                STB_Car_Type = 13   'general service flat
            Case 49
                STB_Car_Type = 14   'other flat
            Case 50
                STB_Car_Type = 15   'tank, less than 22000 gallons
            Case 51
                STB_Car_Type = 16   'tank, more than 22000 gallons
            Case 52
                STB_Car_Type = 17   'all other freight cars
            Case 54
                STB_Car_Type = 17   'average car - used to be cabooses
            Case Else
                STB_Car_Type = 18   'default type of car if not found
        End Select

    End Function

    Function Railroad_Owned_Car(ByVal mValue As String) As Integer

        Select Case mValue
            Case "ABOX", "RBOX", "CSX ", "CSXT", "GONX"
                Railroad_Owned_Car = 0
            Case Else
                If Right(mValue, 1) = "X" Then
                    Railroad_Owned_Car = 1
                Else
                    Railroad_Owned_Car = 0
                End If
        End Select

    End Function

    Function Shipment_Type(
        ByVal mJunction As Integer,
        ByVal mJF As Integer) _
        As Integer

        Shipment_Type = 0

        If mJunction = 1 Then
            If mJF = 1 Then
                Shipment_Type = 1
            Else
                Shipment_Type = 2
            End If
        Else
            If mJunction > 1 Then
                If mJF = mJunction Then
                    Shipment_Type = 3
                Else
                    Shipment_Type = 4
                End If
            End If
        End If

    End Function

    Function TableExist(ByVal mDatabase As String,
                        ByVal mTableName As String) As Boolean

        Dim mDataTable As DataTable
        Dim mSQLStr As String

        gbl_SQLConnection = New SqlConnection
        gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(mDatabase)
        gbl_SQLConnection.Open()

        mSQLStr = "SELECT * FROM dbo.sysobjects WHERE name = '" & mTableName & "'"

        mDataTable = New DataTable

        Using daAdapter As New SqlDataAdapter(mSQLStr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        'If the table exists, rst.recordcount will be greater than 0
        If mDataTable.Rows.Count > 0 Then
            TableExist = True
        Else
            TableExist = False
        End If

        mDataTable = Nothing

    End Function

    Function FunctionExist(ByVal mYear As String) As Boolean

        Dim mDataTable As DataTable
        Dim mSQLStr As String

        gbl_Database_Name = "URCS" & mYear

        gbl_SQLConnection = New SqlConnection
        gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(gbl_Database_Name)
        gbl_SQLConnection.Open()

        mSQLStr = "SELECT * FROM dbo.sysobjects WHERE name = 'ufn_EValues'"

        mDataTable = New DataTable

        Using daAdapter As New SqlDataAdapter(mSQLStr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        'If the function exists, rst.recordcount will be greater than 0
        If mDataTable.Rows.Count > 0 Then
            FunctionExist = True
        Else
            FunctionExist = False
        End If

        mDataTable = Nothing

    End Function

    Function ProcedureExist(ByVal mYear As String,
                            ByVal mProcName As String) As Boolean

        Dim mDataTable As DataTable
        Dim mSQLStr As String

        gbl_Database_Name = "URCS" & mYear

        gbl_SQLConnection = New SqlConnection
        gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(gbl_Database_Name)
        gbl_SQLConnection.Open()

        mSQLStr = "SELECT * FROM dbo.sysobjects WHERE name = '" & mProcName & "'"

        mDataTable = New DataTable

        Using daAdapter As New SqlDataAdapter(mSQLStr, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        'If the function exists, rst.recordcount will be greater than 0
        If mDataTable.Rows.Count > 0 Then
            ProcedureExist = True
        Else
            ProcedureExist = False
        End If

        mDataTable = Nothing

    End Function

    Function Set_Commodity_Schedule(ByVal mSTCC As Object) As Integer

        Dim mWorkVal As Integer

        Set_Commodity_Schedule = 0

        If mSTCC IsNot DBNull.Value Then
            If Len(Trim(mSTCC)) = 2 Then
                mWorkVal = CInt(mSTCC)
                Select Case mWorkVal
                    Case 1 To 47
                        Set_Commodity_Schedule = mWorkVal + 100
                    Case 48
                        Set_Commodity_Schedule = 107
                    Case 98
                        Set_Commodity_Schedule = 146
                    Case 99
                        Set_Commodity_Schedule = 147
                    Case Else
                        Set_Commodity_Schedule = 0
                End Select
            End If
        End If

    End Function

    Function Set_URCS_Code(ByVal RRICC As Decimal) As Integer

        Select Case RRICC
            Case 113300
                Set_URCS_Code = 5    'CR
            Case 114900
                Set_URCS_Code = 8    'GTW
            Case 117000
                Set_URCS_Code = 10   'NS
            Case 124100
                Set_URCS_Code = 18   'IC
            Case 125600
                Set_URCS_Code = 20   'CSX
            Case 130500
                Set_URCS_Code = 22   'BN
            Case 134500
                Set_URCS_Code = 30   'KCS
            Case 137700
                Set_URCS_Code = 35   'SOO
            Case 138100
                Set_URCS_Code = 36   'SP
            Case 139300
                Set_URCS_Code = 37   'UP
            Case 900004
                Set_URCS_Code = 44   'Region 4 - East
            Case 900007
                Set_URCS_Code = 47   'Region 7 - West
            Case 900099
                Set_URCS_Code = 49   'National
            Case Else
                Set_URCS_Code = 0
        End Select
    End Function

    Sub LoadArrayData()
        '**********************************************************************
        ' Title:        Load Array Data
        ' Author:       Michael Sanders
        ' Purpose:      This Subroutine loads the Unmasking and Railroad information to global arrays
        ' Revisions:    Conversion from Access database/VBA - 14 Mar 2013
        '               Conversion to Compartmentalized Databases - 24 Jul 2015
        ' 
        ' This program is US Government Property - For Official Use Only
        '**********************************************************************
        ' Database & Tables List
        '
        ' Databases used
        ' ----------------------------------------------------------------------------
        ' Global_Variables.Gbl_Controls_Database_Name
        '
        ' Tables used (Assigned via Table Locator table)
        ' ----------------------------------------------------------------------------
        ' Global_Variables.Gbl_Unmasking_BNSF_TableName
        ' Global_Variables.Gbl_Unmasking_CSX1990_TableName
        ' Global_Variables.Gbl_Unmasking_CSX1991_TableName
        ' Global_Variables.Gbl_Unmasking_CSX2020_Tablename
        ' Global_Variables.Gbl_Unmasking_CSXWB_TableName
        ' Global_Variables.Gbl_Unmasking_UP1993_TableName
        ' Global_Variables.Gbl_Unmasking_UP2001_TableName
        ' Global_Variables.Gbl_Unmasking_UP_TableName
        ' Global_Variables.Gbl_Unmasking_Conrail_TableName
        ' Global_Variables.Gbl_Unmasking_CNW_TableName
        ' Global_Variables.Gbl_Unmasking_Generic_TableName
        ' Global_Variables.Gbl_Railroads_TableName
        ' WAYRRR

        Dim rst As ADODB.Recordset
        Dim mStrSQL As String
        Dim mIndex As Integer

        ' Open the connection to the Controls database
        OpenADOConnection(Gbl_Controls_Database_Name)

        'Load BNSF Array
        Gbl_Unmasking_BNSF_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_BNSF")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_BNSF_TableName
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            BNSFunmaskArray(rst.Fields("bnsfindex1").Value, rst.Fields("bnsfindex2").Value) =
                rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        'Load CSXT Arrays
        Gbl_Unmasking_CSX1990_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_CSX1990")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CSX1990_TableName & " where RecordType = 'STCC'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CSX1990STCC(rst.Fields("csxindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        Gbl_Unmasking_CSX1990_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_CSX1990")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CSX1990_TableName & " where RecordType = 'RATE'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CSX1990Rate(rst.Fields("csxindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        Gbl_Unmasking_CSX1991_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_CSX1991")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CSX1991_TableName & " where RecordType = 'STCC'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CSX1991STCC(rst.Fields("csxindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        Gbl_Unmasking_CSX1991_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_CSX1991")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CSX1991_TableName & " where RecordType = 'RATE'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CSX1991Rate(rst.Fields("csxindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        Gbl_Unmasking_CSXWB_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_CSXWB")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CSXWB_TableName & " where csxindex1 = 0"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CSXwb00(rst.Fields("csxindex2").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CSXWB_TableName & " where csxindex1 = 20"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CSXwb20(rst.Fields("csxindex2").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CSXWB_TableName & " where csxindex1 = 40"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CSXwb40(rst.Fields("csxindex2").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CSXWB_TableName & " where csxindex1 = 60"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CSXwb60(rst.Fields("csxindex2").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CSXWB_TableName & " where csxindex1 = 80"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CSXwb80(rst.Fields("csxindex2").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        Gbl_Unmasking_CSX2020WB_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_CSX2020")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CSX2020WB_TableName
        rst.Open(mStrSQL, gbl_ADOConnection)

        mIndex = 1

        rst.MoveFirst()
        Do While Not rst.EOF

            CSX2020_STCC(mIndex) = rst.Fields("STCC").Value
            CSX2020_00(mIndex) = rst.Fields("Last_Two_WB_Num_00-19").Value
            CSX2020_20(mIndex) = rst.Fields("Last_Two_WB_Num_20-39").Value
            CSX2020_40(mIndex) = rst.Fields("Last_Two_WB_Num_40-59").Value
            CSX2020_60(mIndex) = rst.Fields("Last_Two_WB_Num_60-79").Value
            CSX2020_80(mIndex) = rst.Fields("Last_Two_WB_Num_80-99").Value

            mIndex = mIndex + 1
            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        'Load UP Arrays
        Gbl_Unmasking_UP1993_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_UP1993")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_UP1993_TableName & " where RecType = 'LOW'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            UP1993STCC_Low(rst.Fields("UPindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_UP1993_TableName & " where RecType = 'HIGH'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            UP1993STCC_High(rst.Fields("UPindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_UP1993_TableName & " where RecType = 'ROW'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            UP1993STCC_Row(rst.Fields("UPindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        Gbl_Unmasking_UP2001_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_UP2001")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_UP2001_TableName & " where RecType = 'LOW'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            UP2001STCC_Low(rst.Fields("UPindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_UP2001_TableName & " where RecType = 'HIGH'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            UP2001STCC_High(rst.Fields("UPindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_UP2001_TableName & " where RecType = 'ROW'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            UP2001STCC_Row(rst.Fields("UPindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        Gbl_Unmasking_UP_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_UP")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_UP_TableName & " where RecType = 'COL'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            UPWbNum_Col(rst.Fields("UPindex1").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_UP_TableName & " where RecType = 'GRP'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            UPfgrp(rst.Fields("UPindex1").Value, rst.Fields("UPIndex2").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        'Load Conrail Arrays

        Gbl_Unmasking_Conrail_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_CONRAIL")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_Conrail_TableName & " where RecType = 'LOCALRATE'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            ConrailLocalRate(rst.Fields("CRindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_Conrail_TableName & " where RecType = 'LOCALSTCC'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            ConrailLocalRate(rst.Fields("CRindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_Conrail_TableName & " where RecType = 'INTERRATE'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            ConrailLocalRate(rst.Fields("CRindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_Conrail_TableName & " where RecType = 'INTERSTCC'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            ConrailLocalRate(rst.Fields("CRindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        ' Load CNW Arrays

        Gbl_Unmasking_CNW_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_CNW")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CNW_TableName & " where RecType = 'STATE'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CNWState(rst.Fields("CNWindex").Value) = rst.Fields("CNWText").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CNW_TableName & " where RecType = 'MULT'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CNWMult(rst.Fields("CNWindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_CNW_TableName & " where RecType = 'UNIT'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            CNWUnit(rst.Fields("CNWindex").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        ' Load Generic indexes

        Gbl_Unmasking_Generic_TableName = Get_Table_Name_From_SQL("1", "R_UNMASKING_GENERIC")
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_Generic_TableName & " where RecType = 'ODD'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            RR_Odd_Factor(rst.Fields("GenIndex1").Value, rst.Fields("GenIndex2").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Unmasking_Generic_TableName & " where RecType = 'EVEN'"
        rst.Open(mStrSQL, gbl_ADOConnection)

        rst.MoveFirst()
        Do While Not rst.EOF

            RR_Even_Factor(rst.Fields("GenIndex1").Value, rst.Fields("GenIndex2").Value) = rst.Fields("factor").Value

            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

        'Load the railroad info from the lookup table to the arrays
        ' This needs to be converted to WAYRRR if possible

        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Get_Table_Name_From_SQL("1", "WAYRRR") & " WHERE URCSCODE <> ''"
        rst.Open(mStrSQL, gbl_ADOConnection)
        mIndex = 1

        Do While Not rst.EOF
            Class1Railroads(mIndex) = rst.Fields("aarid").Value
            Class1RailroadName(mIndex) = rst.Fields("rr_name").Value
            Class1Abbv(mIndex) = Trim(rst.Fields("rr_alpha").Value)
            Class1RRICC(mIndex) = ReturnDecimal(rst.Fields("rricc_id").Value)
            URCSCodes(mIndex) = ReturnInteger(rst.Fields("urcscode").Value)
            Class1Regions(mIndex) = ReturnInteger(rst.Fields("region_id").Value)
            rst.MoveNext()
            mIndex = mIndex + 1
        Loop

        rst.Close()
        rst = Nothing

    End Sub

    Function Divide(ByVal numerator As Double, ByVal denominator As Double) As Double
        ' Using this function prevents a divide by zero situation

        If denominator <> 0 Then
            Divide = numerator / denominator
        Else
            Divide = 0
        End If

    End Function

    Function ReturnValidNumber(ByVal mItem As Object) As Decimal
        If IsNothing(mItem) Then
            ReturnValidNumber = 0
        ElseIf IsNumeric(mItem.ToString) Then
            ReturnValidNumber = mItem
        Else
            ReturnValidNumber = 0
        End If
    End Function

    Function FindDataYear(ByVal mYear As Integer) As String

        Dim rst As New ADODB.Recordset
        Dim strSQL As String
        Dim mTableName As String
        Dim mDatabaseName As String
        Dim mGotIt As Boolean
        Dim mDataYears(50) As Integer, mLooper As Integer, mPos As Integer
        Dim mMaskedTableNames(50) As String
        Dim mMaskedOnly(50) As Boolean

        FindDataYear = 0
        Get_Table_Name_From_SQL(1, "WB_Years")
        mDatabaseName = Global_Variables.gbl_Database_Name
        mTableName = Global_Variables.gbl_Table_Name

        ' Open the SQL connection
        OpenADOConnection(mDatabaseName)

        rst = SetRST()
        mGotIt = False
        For mLooper = 1 To 50
            mMaskedOnly(mLooper) = False
        Next mLooper

        'We'll load the years from the tlk_WB_Years table
        strSQL = "SELECT * from " & mTableName
        rst.Open(strSQL, Global_Variables.gbl_ADOConnection)
        rst.MoveFirst()
        mLooper = 1
        Do While Not rst.EOF
            mDataYears(mLooper) = rst.Fields("wb_year").Value
            mMaskedTableNames(mLooper) = rst.Fields("wb_maskedtablename").Value
            mMaskedOnly(mLooper) = rst.Fields("wb_maskedonlyflag").Value
            rst.MoveNext()
            mLooper = mLooper + 1
        Loop

        rst.Close()
        rst = Nothing

        'Now we find the year in mDataYears to pass back to the calling sub.
        mPos = Array1DFindFirst(mDataYears, mYear, 0)
        FindDataYear = mMaskedTableNames(mPos)

    End Function

    Function ReturnString(ByVal mValue As Object, Optional ByVal mlen As Integer = 0) As String
        If mValue.ToString IsNot "" Then
            ReturnString = mValue
        Else
            ReturnString = Space(mlen)
        End If
    End Function

    Function ReturnByte(ByVal mValue As Object) As Byte
        If IsNumeric(mValue) Then
            ReturnByte = CByte(mValue)
        Else
            ReturnByte = 0
        End If
    End Function

    Function ReturnDate(ByVal mValue As Object) As Date
        If IsDate(mValue) Then
            If IsDate(mValue) Then
                ReturnDate = CDate(mValue)
            Else
                ReturnDate = ""
            End If
        Else
            ReturnDate = ""
        End If
    End Function

    Function ReturnDecimal(ByVal mValue As Object) As Decimal
        If IsNumeric(mValue) Then
            ReturnDecimal = CDec(mValue)
        Else
            ReturnDecimal = 0
        End If
    End Function

    Function ReturnInteger(ByVal mValue As Object) As Integer
        If IsNumeric(mValue) Then
            ReturnInteger = CInt(mValue)
        Else
            ReturnInteger = 0
        End If
    End Function

    Function ReturnLong(ByVal mValue As Object) As Long
        If IsNumeric(mValue) Then
            ReturnLong = CLng(mValue)
        Else
            ReturnLong = 0
        End If
    End Function

    Function ArrayFind1993UPRange(
         ByVal vFindItem As Object,
         Optional ByVal lDefault As Long = 0) As Long

        'Purpose   :    Finds the position of the Range matching item in UP 1993 arrays.
        'Inputs    :    vFindItem                   The value to look for in the arrays.
        '               [lDefault]                  The value to return if vFindItem is not found
        'Outputs   :    The position of the item with the array
        '               OR 0/lDefault if the item was not found.
        'Notes     :    Make the module Option Compare Text for case insensative searches


        Dim lThisRow As Long, bFound As Boolean

        For lThisRow = LBound(UP1993STCC_Low) To UBound(UP1993STCC_Low)
            If (UP1993STCC_Low(lThisRow) <= vFindItem) And (vFindItem <= UP1993STCC_High(lThisRow)) Then
                bFound = True
                ArrayFind1993UPRange = lThisRow
                Exit For
            End If
        Next

        If bFound = False Then
            ArrayFind1993UPRange = lDefault
        End If
        On Error GoTo 0

    End Function

    Function ArrayFind2001UPRange(
         ByVal vFindItem As Object,
         Optional ByVal lDefault As Long = 0) As Long

        'Purpose   :    Finds the position of the Range matching item in UP arrays.
        'Inputs    :    vFindItem                   The value to look for in the arrays.
        '               [lDefault]                  The value to return if vFindItem is not found
        'Outputs   :    The position of the item with the array
        '               OR 0/lDefault if the item was not found.
        'Notes     :    Make the module Option Compare Text for case insensative searches


        Dim lThisRow As Long, bFound As Boolean

        For lThisRow = LBound(UP2001STCC_Low) To UBound(UP2001STCC_Low)
            If (UP2001STCC_Low(lThisRow) <= vFindItem) And (vFindItem <= UP2001STCC_High(lThisRow)) Then
                bFound = True
                ArrayFind2001UPRange = lThisRow
                Exit For
            End If
        Next

        If bFound = False Then
            ArrayFind2001UPRange = lDefault
        End If
        On Error GoTo 0

    End Function

    Function Field_Right(
         ByVal rstField As Object,
         ByVal FieldLength As Integer) As String

        Dim mDec As Decimal
        Dim mString As String
        Dim mPad As Long
        Dim y As Integer
        Dim mformatstr As String

        mformatstr = ""

        If IsDBNull(rstField) Then
            Field_Right = Space(FieldLength)
        Else
            If IsNumeric(rstField) Then
                mDec = Str(Math.Round(CDec(rstField), MidpointRounding.AwayFromZero))
                If mDec >= 0 Then
                    For y = 1 To FieldLength
                        mformatstr = mformatstr & "0"
                    Next
                Else            ' Accomodate the negative sign
                    For y = 1 To (FieldLength - 1)
                        mformatstr = mformatstr & "0"
                    Next
                End If
                If mformatstr = "" Then
                    mformatstr = "0"
                End If
                Field_Right = Format(mDec, mformatstr)
            Else
                mString = rstField
                mPad = Len(mString)
                Field_Right = Space(FieldLength - mPad) & CStr(mString)
            End If
        End If

    End Function

    Function Field_Left(
         ByVal rstField As Object,
         ByVal FieldLength As Integer,
         Optional ByVal PadZero As Boolean = False) As String

        Dim mString As String
        Dim mPad As Long

        Field_Left = ""

        If IsDBNull(rstField) Then
            Field_Left = Space(FieldLength)
        Else
            mString = rstField
            mPad = Len(mString)
            If FieldLength = mPad Then
                Field_Left = mString
            Else
                If PadZero = False Then
                    Field_Left = mString & Space(FieldLength - mPad)
                Else
                    Field_Left = mString
                    Do While Len(Field_Left) < FieldLength
                        Field_Left = Field_Left & "0"
                    Loop
                End If
            End If
        End If

    End Function

    Function Zip_File(ByVal FileNamePath As String) As String
        Dim myProcess As New Process()
        Dim FileNameZip As String
        Dim ShellStr As String, strDate As String
        Dim fs

        fs = CreateObject("Scripting.FileSystemObject")

        strDate = Format(Now, " dd-mm-yy h-mm-ss")

        FileNameZip = Left(FileNamePath, Len(FileNamePath) - 4) & ".zip"
        If fs.FileExists(FileNameZip) Then
            Kill(FileNameZip)
        End If

        ShellStr = "-add" _
                 & " " & FileNameZip _
                 & " " & FileNamePath

        Try
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.FileName = "pkzipc.exe"
            myProcess.StartInfo.Arguments = ShellStr
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.Start()
            myProcess.WaitForExit()
        Catch e As System.Exception
            MsgBox(e.Message, MsgBoxStyle.OkOnly, "Error!")
        End Try

        Zip_File = Left(FileNamePath, Len(FileNamePath) - 4) & ".zip"

    End Function

    Public Function Get_AARID(ByRef mYear As Integer, ByRef mRailroad As String) As Integer

        Dim mStrSQL As String

        Dim rst As New ADODB.Recordset

        ' Get the database and table information for the WAYRRR table
        OpenADOConnection(Global_Variables.Gbl_Controls_Database_Name)

        'Find out what the AARID number is for the railroad name selected on the form
        rst = SetRST()

        mStrSQL = "SELECT AARID FROM " & Get_Table_Name_From_SQL("1", "WAYRRR") & " WHERE RR_NAME = '" & mRailroad & "'"
        rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)

        Get_AARID = rst.Fields(0).Value

        rst.Close()
        rst = Nothing

    End Function

    Function SetRegion(ByVal mRegionIndex As Integer) As Integer

        Select Case mRegionIndex
            Case 1
                SetRegion = 4
            Case 2
                SetRegion = 7
            Case 3
                SetRegion = 9
            Case Else
                SetRegion = 0   'Error!
        End Select

    End Function

    Function RoundD(ByVal dec As Decimal)
        Dim d As Decimal = dec
        Dim r As Decimal = Math.Ceiling(d * 100D) / 100D
        Return r
    End Function

    Function RemoveSpaces(ByVal mString As String) As String
        Dim mLooper As Integer

        RemoveSpaces = mString
        For mLooper = 1 To 50
            RemoveSpaces = Replace(RemoveSpaces, "  ", " ")
        Next

    End Function

    Function CheckNull(ByVal mstring As String) As String
        If IsDBNull(mstring) Then
            CheckNull = "Null"
        Else
            CheckNull = mstring
        End If
    End Function

    Function Calculate_Variance(LastYearValue As Decimal, ThisYearValue As Decimal) As Decimal

        Calculate_Variance = 0

        If LastYearValue > 0 Then
            Calculate_Variance = ((LastYearValue - ThisYearValue) / LastYearValue) * 100
        End If

        Calculate_Variance = Calculate_Variance * -1

    End Function

    Public Function ColText(ByVal mCol As String) As String

        Dim mThisCol As String = ""

        Select Case mCol
            Case "C1"
                mThisCol = " Col B"
            Case "C2"
                mThisCol = " Col C"
            Case "C3"
                mThisCol = " Col D"
            Case "C4"
                mThisCol = " Col E"
            Case "C5"
                mThisCol = " Col F"
            Case "C6"
                mThisCol = " Col G"
            Case "C7"
                mThisCol = " Col H"
            Case "C8"
                mThisCol = " Col I"
            Case "C9"
                mThisCol = " Col J"
            Case "C10"
                mThisCol = " Col K"
            Case "C11"
                mThisCol = " Col L"
            Case "C12"
                mThisCol = " Col M"
            Case "C13"
                mThisCol = " Col N"
            Case "C14"
                mThisCol = " Col O"
            Case "C15"
                mThisCol = " Col P"
        End Select

        ColText = mThisCol

    End Function

    Public Function ColLetter(ByVal mCol As String) As String

        Dim mThisChar As String = "BCDEFGHIJKLMNOP"
        Dim mCharArray() As Char = mThisChar.ToCharArray

        ColLetter = mCharArray(mCol - 1)

    End Function

    Public Function Convert_Zero_Based_Column_Number_To_Text(ByVal mColumnNumber As Integer) As String

        Dim mThisChar As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim mCharArray() As Char = mThisChar.ToCharArray

        Convert_Zero_Based_Column_Number_To_Text = mCharArray(mColumnNumber)

    End Function

    Public Function TotalLines(filePath As String) As Integer
        Using r As New StreamReader(filePath)
            Dim i As Integer = 0
            While r.ReadLine() IsNot Nothing
                i += 1
            End While
            Return i
        End Using
    End Function

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Convert aar car type to urcs car type. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/14/2020. </remarks>
    '''
    ''' <param name="AAR_Car_Type"> Type of the aar car. </param>
    '''
    ''' <returns>   The aar converted car type to urcs car type. </returns>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Function Convert_AAR_Car_Type_To_URCS_Car_Type(ByVal AAR_Car_Type As String) As Integer

        Dim mEquipmentType As String = Mid(AAR_Car_Type, 1, 1)
        Dim mFirstNumeric As Integer = Val(Mid(AAR_Car_Type, 2, 1))
        Dim mSecondNumeric As Integer = Val(Mid(AAR_Car_Type, 3, 1))
        Dim mThirdNumeric As Integer = Val(Mid(AAR_Car_Type, 4, 1))
        Dim mFirstTwo As String = Mid(AAR_Car_Type, 1, 2)
        Dim mFirstThree As String = Mid(AAR_Car_Type, 1, 3)

        'Set default URCS_Car_Typ
        Convert_AAR_Car_Type_To_URCS_Car_Type = 17

        Select Case mEquipmentType
            Case "A"
                Select Case mSecondNumeric
                    Case 0 To 4, 6 To 9
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 3
                    Case 5
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 17
                End Select
            Case "B"
                Select Case mFirstNumeric
                    Case 1, 2
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 1
                    Case 3 To 8
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 2
                End Select
            Case "C"
                Select Case mThirdNumeric
                    Case 1 To 4
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 6
                End Select
            Case "E"
                Convert_AAR_Car_Type_To_URCS_Car_Type = 5
            Case "F"
                Select Case mFirstThree
                    Case "F10", "F20", "F30"
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 13
                    Case "F40"
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 14
                End Select
                Select Case mSecondNumeric
                    Case 1 To 6, 8
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 14
                    Case 7
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 17
                End Select
            Case "G"
                Convert_AAR_Car_Type_To_URCS_Car_Type = 4
            Case "H"
                Convert_AAR_Car_Type_To_URCS_Car_Type = 7
            Case "J"
                Select Case mThirdNumeric
                    Case 0
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 8
                    Case 1 To 4
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 4
                End Select
            Case "K"
                Convert_AAR_Car_Type_To_URCS_Car_Type = 8
            Case "L"
                Convert_AAR_Car_Type_To_URCS_Car_Type = 17
            Case "M"
                If AAR_Car_Type = "M930" Then
                    Convert_AAR_Car_Type_To_URCS_Car_Type = 17
                End If
            Case "P"
                Convert_AAR_Car_Type_To_URCS_Car_Type = 11
            Case "Q"
                If mFirstNumeric = 8 Then
                    Convert_AAR_Car_Type_To_URCS_Car_Type = 17
                Else
                    Convert_AAR_Car_Type_To_URCS_Car_Type = 11
                End If
            Case "R"
                Select Case mSecondNumeric
                    Case 0 To 2
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 10
                    Case 5 To 9
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 9
                End Select
            Case "S"
                Convert_AAR_Car_Type_To_URCS_Car_Type = 11
            Case "T"
                Select Case mThirdNumeric
                    Case 0 To 5
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 15
                    Case 6 To 9
                        Convert_AAR_Car_Type_To_URCS_Car_Type = 16
                End Select
            Case "V"
                Convert_AAR_Car_Type_To_URCS_Car_Type = 12
        End Select

    End Function

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Convert urcs car typ to sch 710 stb car typ. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/14/2020. </remarks>
    '''
    ''' <param name="URCS_Car_Typ"> The urcs car typ. </param>
    '''
    ''' <returns>   The urcs converted car typ to sch 710 stb car typ. </returns>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Function Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ(ByVal URCS_Car_Typ As Integer) As Integer

        ' Note: Although it would be easier to calculate the SCH710 Car Type value by adding
        ' 35 to the URCS_Car_Typ, it would not convey the meaning of each.  I wrote this long
        ' version for clarity.

        ' Set the default
        Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 52   'default type of car if not found

        Select Case URCS_Car_Typ
            Case 1
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 36    'unequipped box car
            Case 2
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 37    '50 ft box car
            Case 3
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 38    'equipped box car
            Case 4
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 39    'unequipped general service gondola
            Case 5
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 40    'equipped general service gondola
            Case 6
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 41    'covered hopper
            Case 7
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 42    'general service covered hopper
            Case 8
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 43    'open, special service hopper
            Case 9
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 44    'Mechanical Refrigerator
            Case 10
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 45   'non-mechanical refrigerator
            Case 11
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 46   'TOFC flat
            Case 12
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 47   'multi level flat
            Case 13
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 48   'general service flat
            Case 14
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 49   'other flat
            Case 15
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 50   'tank, less than 22000 gallons
            Case 16
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 51   'tank, more than 22000 gallons
            Case 17
                Convert_URCS_Car_Typ_To_Sch710_STB_Car_Typ = 52   'all other freight cars
        End Select

    End Function

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Zero fill numeric. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/14/2020. </remarks>
    '''
    ''' <param name="mField">       The field. </param>
    ''' <param name="mFieldLength"> Length of the field. </param>
    '''
    ''' <returns>   A String. </returns>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Function Zero_Fill_Numeric(
        ByVal mField As Object,
        ByVal mFieldLength As Integer) As String

        Dim mString As String
        Dim y As Integer
        Dim mformatstr As String

        Zero_Fill_Numeric = mField.ToString
        mformatstr = ""

        If String.IsNullOrEmpty(mField) Then
            For y = 1 To mFieldLength
                Zero_Fill_Numeric = Zero_Fill_Numeric & "0"
            Next
        Else
            If IsNumeric(mField) Then
                mString = mField
                mformatstr = Trim(CStr(mField))
                For y = 1 To mFieldLength - Len(mformatstr)
                    mString = "0" & mString
                Next
                Zero_Fill_Numeric = mString
            End If
        End If

    End Function

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Copy files from one directory to another. </summary>
    '''
    ''' <remarks>   Michael Sanders, 2/22/2022. </remarks>
    '''
    ''' <param name="sourcePath">       The folder that the files are copied from. </param>
    ''' <param name="destinationPath">  The folder that the files are copied to. </param>
    '''
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CopyFiles(ByVal sourcePath As String, ByVal DestinationPath As String, ByVal DeleteSource As Boolean)
        If (Directory.Exists(sourcePath)) Then
            For Each fName As String In Directory.GetFiles(sourcePath)
                If File.Exists(fName) Then
                    Dim dFile As String = String.Empty

                    dFile = Path.GetFileName(fName)
                    Dim dFilePath As String = String.Empty

                    dFilePath = DestinationPath + dFile
                    File.Copy(fName, dFilePath, True)

                    If DeleteSource = True Then
                        File.Delete(fName)
                    End If
                End If
            Next
        End If
    End Sub

    Function Return_Elapsed_Time(mTimeSpan_Start As DateTime, mTimeSpan_Stop As DateTime) As String

        Dim mWorkStr As New StringBuilder
        Dim mTS As TimeSpan

        mWorkStr.Append("Elapsed Time: ")
        mTS = mTimeSpan_Stop - mTimeSpan_Start

        If mTS.Days > 0 Then
            If mTS.Days < 2 Then
                mWorkStr.Append(mTS.Days.ToString & " Day,")
            Else
                mWorkStr.Append(mTS.Days.ToString & " Days,")
            End If
        End If

        If mTS.Hours > 0 Then
            If mTS.Hours < 2 Then
                mWorkStr.Append(mTS.Hours.ToString & " Hour, ")
            Else
                mWorkStr.Append(mTS.Hours.ToString & " Hours, ")
            End If
        End If

        If mTS.Minutes > 0 Then
            If mTS.Minutes < 2 Then
                mWorkStr.Append(mTS.Minutes.ToString & " Minute, ")
            Else
                mWorkStr.Append(mTS.Minutes.ToString & " Minutes, ")
            End If
        End If

        If mTS.Seconds < 2 Then
            mWorkStr.Append(mTS.Seconds.ToString & " Second")
        Else
            mWorkStr.Append(mTS.Seconds.ToString & " Seconds")
        End If

        Return_Elapsed_Time = mWorkStr.ToString

    End Function
End Module
