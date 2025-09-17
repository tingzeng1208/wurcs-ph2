Imports System.Data.SqlClient
Module Global_Variables

    Public Const ForReading = 1, ForWriting = 2, ForAppending = 3

    ' This are the store for the SQL Server status.  
    Public gbl_Server_Name As String            'Store for the server that the user ihas actively chosen
    Public gbl_Production_Server_OK As Boolean
    Public gbl_Test_Server_OK As Boolean
    Public gbl_Local_Server_OK As Boolean


    ' This is the store for the database name.  It will change during execution
    Public gbl_Database_Name As String
    ' This is the store for a table name.  It will change during execution
    Public gbl_Table_Name As String

    ' These global variables store/maintain the SQL connection strings
    Public gbl_SQLConnString As String
    Public gbl_ADOConnString As String
    Public gbl_Lookup_ADOConnString As String

    ' These global objects are the SQL connections used by the system
    Public gbl_SQLConnection As SqlConnection
    Public gbl_ADOConnection As ADODB.Connection
    Public gbl_Lookup_ADOConnection As ADODB.Connection

    ' This will hold a copy of the Table_Locator table from SQL so that we don't have to keep querying it.
    Public gbl_Table_Locator_Table As DataTable

    ' Set by the Menu_Main.vb load event, this boolean is used to control displaying of
    ' modules that only a SQL writer should see
    Public gbl_IsWriter As Boolean

    ' Set by the SplashScreen.vb load event, this is used to control displaying of
    ' the location of the database that was found
    Public gbl_DatabaseLocation As String

    ' These are global elemnts used by multiple forms for file and directory operations
    Public FolderBrowser As FolderBrowserDialog
    Public FileBrowser As FileDialog
    Public Gbl_sOutputFolder As String
    Public Gbl_sLogFolder As String

    ' Used for storing the names of the databases
    Public Gbl_Waybill_Database_Name As String
    Public Gbl_Controls_Database_Name As String
    Public Gbl_URCSProdYear_Database_Name As String
    Public Gbl_Interim_Waybills_Database_Name As String

    ' Used for Control tables - these will always be in the Gbl_Controls_Database_Name database
    Public Gbl_Table_Locator_TableName As String
    Public Gbl_Trans_DatabaseName As String
    Public Gbl_Trans_TableName As String
    Public Gbl_Railroads_TableName As String
    Public Gbl_Price_Index_TableName As String
    Public Gbl_Price_Index_DatabaseName As String
    Public Gbl_URCS_Index_TableName As String
    Public Gbl_URCS_Years_TableName As String
    Public Gbl_URCS_AARIndex_TableName As String
    Public Gbl_URCS_FCS_TableName As String
    Public Gbl_URCS_WAYRRR_TableName As String
    Public Gbl_URCS_Writers_TableName As String
    Public Gbl_WB_Years_TableName As String
    Public Gbl_Class1RailList_TableName As String
    Public Gbl_Class1RailList_DatabaseName As String
    Public Gbl_State_Codes_TableName As String
    Public Gbl_STCC_Codes_TableName As String
    Public Gbl_CSM_TableName As String
    Public Gbl_Marks_Tablename As String
    Public Gbl_Productivity_TableName As String
    Public Gbl_URCS_Schedules_TableName As String
    Public Gbl_Unmasking_BNSF_TableName As String
    Public Gbl_Unmasking_CNW_TableName As String
    Public Gbl_Unmasking_Conrail_TableName As String
    Public Gbl_Unmasking_CSX1990_TableName As String
    Public Gbl_Unmasking_CSX1991_TableName As String
    Public Gbl_Unmasking_CSX2020WB_TableName As String
    Public Gbl_Unmasking_CSXWB_TableName As String
    Public Gbl_Unmasking_Generic_TableName As String
    Public Gbl_Unmasking_UP_TableName As String
    Public Gbl_Unmasking_UP1993_TableName As String
    Public Gbl_Unmasking_UP2001_TableName As String
    Public Gbl_AuditTrailLog_Tablename As String
    Public Gbl_Ordnance_STCCs As String
    Public Gbl_Interim_Masked As String
    Public Gbl_Annual_Interim_TableName As String
    Public Gbl_Interim_Raw As String
    Public Gbl_Interim_Unmasked_Rev As String
    Public Gbl_Interim_Segments As String
    Public Gbl_Interim_Unmasked_Segments As String
    Public Gbl_Interim_BatchPro_All_Miled As String
    Public Gbl_interim_PUWS As String
    Public Gbl_Interim_PUWS_Masked_Rev_TableName As String

    ' Used for Waybill tables - these will always be in the Gbl_Waybill_Database_Name database
    Public Gbl_Masked_TableName As String
    Public Gbl_RailInc_455_TableName As String
    Public Gbl_Unmasked_Rev_TableName As String
    Public Gbl_PUWS_Masked_Tablename As String
    Public Gbl_PUWS_Masking_Factors_Tablename As String
    Public Gbl_PUWS_ReMasked_Tablename As String
    Public Gbl_PUWS_ReMasking_Factors_Tablename As String
    Public Gbl_Segments_TableName As String
    Public Gbl_Unmasked_Segments_TableName As String
    Public Gbl_EIA_STCCs_TableName As String
    Public Gbl_STCC_W49_Translation_TableName As String

    ' Used for URCS Production for a single year tables
    ' These will always be in the Gbl_URCSProdYear_Database_Name database - URCSyyyy
    Public Gbl_AValues_TableName As String
    Public Gbl_AValues_DatabaseName As String
    Public Gbl_ACode_TableName As String
    Public Gbl_ACode_DatabaseName As String
    Public Gbl_Region_TableName As String
    Public Gbl_Op_Stats_By_Car_Type_TableName As String
    Public Gbl_Op_Stats_By_Car_Type2_TableName As String
    Public Gbl_Op_Stats_By_Car_Type3_TableName As String
    Public Gbl_Line_Source_Text_TableName As String
    Public Gbl_Data_Dictionary_TableName As String
    Public Gbl_ECode_TableName As String
    Public Gbl_Errors_TableName As String
    Public Gbl_Substitutions_TableName As String
    Public Gbl_EValues_TableName As String
    Public Gbl_URCS_Defaults_TableName As String
    Public Gbl_URCS_Codes_TableName As String
    Public Gbl_URCS_Tare_TableName As String
    Public Gbl_Makewhole_Factors_Tablename As String

    'Public Global Variable Declarations
    '-------------------------------

    'BNSF Array
    Public BNSFunmaskArray(3, 7) As Decimal

    'CSX Arrays
    Public CSX1990STCC(44) As Integer
    Public CSX1990Rate(44) As Decimal
    Public CSX1991STCC(48) As Integer
    Public CSX1991Rate(48) As Decimal
    Public CSX2020STCC() As Integer
    Public CSX2020Rate() As Decimal
    Public CSXwb00(48) As Decimal
    Public CSXwb20(48) As Decimal
    Public CSXwb40(48) As Decimal
    Public CSXwb60(48) As Decimal
    Public CSXwb80(48) As Decimal
    Public CSX2020_STCC(170) As Integer
    Public CSX2020_00(170) As Decimal
    Public CSX2020_20(170) As Decimal
    Public CSX2020_40(170) As Decimal
    Public CSX2020_60(170) As Decimal
    Public CSX2020_80(170) As Decimal

    'UP arrays
    Public UP1993STCC_Low(115) As Decimal
    Public UP1993STCC_High(115) As Decimal
    Public UP1993STCC_Row(115) As Integer
    Public UP1993WbNum_Col(100) As Integer
    Public UP2001STCC_Low(86) As Decimal
    Public UP2001STCC_High(86) As Decimal
    Public UP2001STCC_Row(86) As Integer
    Public UPWbNum_Col(100) As Integer
    Public UPfgrp(6, 5) As Decimal

    'Conrail Arrays
    Public ConrailLocalRate(29) As Decimal
    Public ConrailLocalSTCC(29) As Integer
    Public ConrailInterRate(28) As Decimal
    Public ConrailInterSTCC(28) As Integer

    'CNW Arrays

    Public CNWState(9) As String
    Public CNWMult(9) As Single
    Public CNWUnit(9) As Single

    ' Generic Railroad Arrays
    Public RR_Odd_Factor(10, 10) As Single
    Public RR_Even_Factor(10, 10) As Single

    'Class 1 Railroad Array
    Public Class1Railroads(9) As Integer

    'Class 1 Railroad Names
    Public Class1RailroadName(9) As String

    'Class 1 Railroad Abbreviations List
    Public Class1Abbv(9) As String

    'Class 1 Railroad RRICC List
    Public Class1RRICC(9) As Decimal

    'URCS Codes Array
    Public URCSCodes(9) As Integer

    'Region values for each road
    Public Class1Regions(9) As Integer

    'Waybill Years Array
    Public WBYears(50) As String

    'Waybill Table Names Array
    Public WBTables(50) As String

    'Waybill Masked Only Flag Array
    Public WBMaskedOnly(50) As Boolean

End Module
