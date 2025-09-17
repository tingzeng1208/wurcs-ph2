Imports System.Data.SqlClient

' Converted from ADODB to SQLClient 12/7/2021
' By Michael Sanders
' Changed to use new format of input file 12/8/2021
' By Michael Sanders

Public Class TOFC_Load_Factors

    Private Const ForWriting = 8

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close this Form
        Me.Close()
    End Sub

    Private Sub TOFC_Load_Factors_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        mDataTable = Get_URCS_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            cmb_URCSYear.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
        Next

        mDataTable = Nothing
    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click
        Dim mSQLCommand As SqlCommand
        Dim xclcnn As New ADODB.Connection

        Dim xclrst As ADODB.Recordset

        Dim xcl As System.Data.OleDb.OleDbConnection

        Dim mWorkVal As Decimal, mWorkStr As String
        Dim mStrSQL As String

        ' Perform Error checking for form controls
        If Me.cmb_URCSYear.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo EndIt
        End If

        If Me.txt_Input_FilePath.Text = "" Then
            MsgBox("You must select an input file.", vbOKOnly)
            GoTo EndIt
        End If

        'Open the connection to the Controls database
        OpenSQLConnection(My.Settings.Controls_DB)

        'Set the tablename
        Gbl_Trans_TableName = Get_Table_Name_From_SQL(1, "Trans")

        'Set up the connection to the Excal file
        xclcnn = GetExcelConnection(Me.txt_Input_FilePath.Text)

        ' Establish the connection to the spreadsheet file
        xcl = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source= " &
                Me.txt_Input_FilePath.Text & "; Extended Properties=Excel 12.0")
        xcl.Open()

        'Delete SCH 42 line 581 for all roads and regions
        mSQLCommand = New SqlCommand
        mSQLCommand.CommandType = CommandType.Text
        mSQLCommand.Connection = gbl_SQLConnection

        mSQLCommand.CommandText = Build_Delete_Trans_SQL_Statement(cmb_URCSYear.Text, 42, 581)
        mSQLCommand.ExecuteNonQuery()

        ' Get the values from this sheet
        mStrSQL = "Select * FROM [Weighted Average Load Factor$D5:M5]" 'Input file format changed
        xclrst = New ADODB.Recordset
        xclrst.Open(mStrSQL, xclcnn)

        'BNSF
        mWorkStr = Replace(xclrst.Fields(0).Name, "#", ".")
        mWorkVal = CDec(mWorkStr) * 1000000
        mSQLCommand.CommandText = Build_Insert_Trans_SQL_Field_Statement(Me.cmb_URCSYear.Text, 130500, 42, 581, 1, mWorkVal).ToString
        mSQLCommand.ExecuteNonQuery()

        'UP
        mWorkStr = Replace(xclrst.Fields(1).Name, "#", ".")
        mWorkVal = CDec(mWorkStr) * 1000000
        mSQLCommand.CommandText = Build_Insert_Trans_SQL_Field_Statement(Me.cmb_URCSYear.Text, 139300, 42, 581, 1, mWorkVal).ToString
        mSQLCommand.ExecuteNonQuery()

        'CP
        mWorkStr = Replace(xclrst.Fields(2).Name, "#", ".")
        mWorkVal = CDec(mWorkStr) * 1000000
        mSQLCommand.CommandText = Build_Insert_Trans_SQL_Field_Statement(Me.cmb_URCSYear.Text, 137700, 42, 581, 1, mWorkVal).ToString
        mSQLCommand.ExecuteNonQuery()

        'KCS
        mWorkStr = Replace(xclrst.Fields(3).Name, "#", ".")
        mWorkVal = CDec(mWorkStr) * 1000000
        mSQLCommand.CommandText = Build_Insert_Trans_SQL_Field_Statement(Me.cmb_URCSYear.Text, 134500, 42, 581, 1, mWorkVal).ToString
        mSQLCommand.ExecuteNonQuery()

        'CSXT
        mWorkStr = Replace(xclrst.Fields(5).Name, "#", ".")
        mWorkVal = CDec(mWorkStr) * 1000000
        mSQLCommand.CommandText = Build_Insert_Trans_SQL_Field_Statement(Me.cmb_URCSYear.Text, 125600, 42, 581, 1, mWorkVal).ToString
        mSQLCommand.ExecuteNonQuery()

        'NS
        mWorkStr = Replace(xclrst.Fields(6).Name, "#", ".")
        mWorkVal = CDec(mWorkStr) * 1000000
        mSQLCommand.CommandText = Build_Insert_Trans_SQL_Field_Statement(Me.cmb_URCSYear.Text, 117000, 42, 581, 1, mWorkVal).ToString
        mSQLCommand.ExecuteNonQuery()

        'CN
        mWorkStr = Replace(xclrst.Fields(7).Name, "#", ".")
        mWorkVal = CDec(mWorkStr) * 1000000
        mSQLCommand.CommandText = Build_Insert_Trans_SQL_Field_Statement(Me.cmb_URCSYear.Text, 114900, 42, 581, 1, mWorkVal).ToString
        mSQLCommand.ExecuteNonQuery()

        xclrst.Close()
        mStrSQL = "SELECT * FROM [Weighted Average Load Factor$H8:M8]" 'Format changed from 2017
        xclrst = New ADODB.Recordset
        xclrst.Open(mStrSQL, xclcnn)

        'Western Region
        mWorkStr = Replace(xclrst.Fields(0).Name, "#", ".")
        mWorkVal = System.Math.Round(CDec(mWorkStr) * 1000000)
        mSQLCommand.CommandText = Build_Insert_Trans_SQL_Field_Statement(Me.cmb_URCSYear.Text, 900007, 42, 581, 1, mWorkVal).ToString
        mSQLCommand.ExecuteNonQuery()

        'Eastern Region
        mWorkStr = Replace(xclrst.Fields(4).Name, "#", ".")
        mWorkVal = System.Math.Round(CDec(mWorkStr) * 1000000)
        mSQLCommand.CommandText = Build_Insert_Trans_SQL_Field_Statement(Me.cmb_URCSYear.Text, 900004, 42, 581, 1, mWorkVal).ToString
        mSQLCommand.ExecuteNonQuery()

        'National
        mWorkStr = Replace(xclrst.Fields(5).Name, "#", ".")
        mWorkVal = System.Math.Round(CDec(mWorkStr) * 1000000)
        mSQLCommand.CommandText = Build_Insert_Trans_SQL_Field_Statement(Me.cmb_URCSYear.Text, 900099, 42, 581, 1, mWorkVal).ToString
        mSQLCommand.ExecuteNonQuery()

        ' We're done
        xclrst.Close()
        xclrst = Nothing

        Me.txt_StatusBox.Text = "Done!"

EndIt:
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub btn_Input_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_FilePath.Text = fd.FileName
        End If
    End Sub

End Class