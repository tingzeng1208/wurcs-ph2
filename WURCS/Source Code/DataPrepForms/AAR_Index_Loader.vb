'Imports Excel = Microsoft.Office.Interop.Excel - No longer used.  10/26/2020 M.Sanders
Imports SpreadsheetGear
Imports System.Data.SqlClient
Public Class AAR_Index_Loader

    '**********************************************************************
    ' Title:        AAR Index Loader Form
    ' Author:       Michael Sanders
    ' Purpose:      This form handles the entry and updating of AAR Index information for URCS.
    ' Revisions:    Conversion from Access database/VBA - 14 Mar 2013
    ' 
    ' This program is US Government Property - For Official Use Only
    '**********************************************************************

    Private mRailroads(50, 3) As Decimal
    Private mRegionCount(3) As Integer
    Private mCurrentIndex(23) As Single
    Private mSingleVal As Single
    Private mBase(3) As Class_URCS_AARIndex
    Private mIndexes(5, 3, 23) As Single
    Dim charactersAllowed As String = "1234567890."

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box fuel us text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Fuel_US_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Fuel_US.TextChanged
        Dim theText As String = TextBox_Fuel_US.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_Fuel_US.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_Fuel_US.Text.Length - 1
            Letter = TextBox_Fuel_US.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_Fuel_US.Text = theText
        TextBox_Fuel_US.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box fuel east text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Fuel_East_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Fuel_East.TextChanged
        Dim theText As String = TextBox_Fuel_East.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_Fuel_East.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_Fuel_East.Text.Length - 1
            Letter = TextBox_Fuel_East.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_Fuel_East.Text = theText
        TextBox_Fuel_East.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box fuel west text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Fuel_West_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Fuel_West.TextChanged
        Dim theText As String = TextBox_Fuel_West.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_Fuel_West.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_Fuel_West.Text.Length - 1
            Letter = TextBox_Fuel_West.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_Fuel_West.Text = theText
        TextBox_Fuel_West.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box milliseconds us text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MS_US_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MS_US.TextChanged
        Dim theText As String = TextBox_MS_US.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_MS_US.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_MS_US.Text.Length - 1
            Letter = TextBox_MS_US.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_MS_US.Text = theText
        TextBox_MS_US.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box milliseconds east text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MS_East_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MS_East.TextChanged
        Dim theText As String = TextBox_MS_East.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_MS_East.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_MS_East.Text.Length - 1
            Letter = TextBox_MS_East.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_MS_East.Text = theText
        TextBox_MS_East.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box milliseconds west text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MS_West_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MS_West.TextChanged
        Dim theText As String = TextBox_MS_West.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_MS_West.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_MS_West.Text.Length - 1
            Letter = TextBox_MS_West.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_MS_West.Text = theText
        TextBox_MS_West.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box ps us text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_PS_US_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_PS_US.TextChanged
        Dim theText As String = TextBox_PS_US.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_PS_US.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_PS_US.Text.Length - 1
            Letter = TextBox_PS_US.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_PS_US.Text = theText
        TextBox_PS_US.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box ps east text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_PS_East_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_PS_East.TextChanged
        Dim theText As String = TextBox_PS_East.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_PS_East.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_PS_East.Text.Length - 1
            Letter = TextBox_PS_East.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_PS_East.Text = theText
        TextBox_PS_East.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box ps west text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_PS_West_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_PS_West.TextChanged
        Dim theText As String = TextBox_PS_West.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_PS_West.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_PS_West.Text.Length - 1
            Letter = TextBox_PS_West.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_PS_West.Text = theText
        TextBox_PS_West.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box wage us text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Wage_US_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Wage_US.TextChanged
        Dim theText As String = TextBox_Wage_US.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_Wage_US.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_Wage_US.Text.Length - 1
            Letter = TextBox_Wage_US.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_Wage_US.Text = theText
        TextBox_Wage_US.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box wage east text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Wage_East_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Wage_East.TextChanged
        Dim theText As String = TextBox_Wage_East.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_Wage_East.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_Wage_East.Text.Length - 1
            Letter = TextBox_Wage_East.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_Wage_East.Text = theText
        TextBox_Wage_East.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box wage west text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Wage_West_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Wage_West.TextChanged
        Dim theText As String = TextBox_Wage_West.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_Wage_West.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_Wage_West.Text.Length - 1
            Letter = TextBox_Wage_West.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_Wage_West.Text = theText
        TextBox_Wage_West.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box mp us text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MP_US_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MP_US.TextChanged
        Dim theText As String = TextBox_MP_US.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_MP_US.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_MP_US.Text.Length - 1
            Letter = TextBox_MP_US.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_MP_US.Text = theText
        TextBox_MP_US.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box mp east text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MP_East_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MP_East.TextChanged
        Dim theText As String = TextBox_MP_East.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_MP_East.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_MP_East.Text.Length - 1
            Letter = TextBox_MP_East.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_MP_East.Text = theText
        TextBox_MP_East.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box mp west text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MP_West_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MP_West.TextChanged
        Dim theText As String = TextBox_MP_West.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox_MP_West.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To TextBox_MP_West.Text.Length - 1
            Letter = TextBox_MP_West.Text.Substring(x, 1)
            If charactersAllowed.Contains(Letter) = False Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox_MP_West.Text = theText
        TextBox_MP_West.Select(SelectionIndex - Change, 0)
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button return to data prep menu click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
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
    ''' <summary>   Aar index loader load. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub AAR_Index_Loader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        CenterToScreen()

        ' Set the default value of 0.0 to the fields
        TextBox_Fuel_US.Text = "0.0"
        TextBox_MS_US.Text = "0.0"
        TextBox_PS_US.Text = "0.0"
        TextBox_Wage_US.Text = "0.0"
        TextBox_MP_US.Text = "0.0"
        TextBox_Fuel_East.Text = "0.0"
        TextBox_MS_East.Text = "0.0"
        TextBox_PS_East.Text = "0.0"
        TextBox_Wage_East.Text = "0.0"
        TextBox_MP_East.Text = "0.0"
        TextBox_Fuel_West.Text = "0.0"
        TextBox_MS_West.Text = "0.0"
        TextBox_PS_West.Text = "0.0"
        TextBox_Wage_West.Text = "0.0"
        TextBox_MP_West.Text = "0.0"

        ' Load the Year combobox from the SQL database
        mDataTable = Get_URCS_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            URCS_Year_Combobox.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
        Next

        mDataTable = Nothing

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Urcs year combobox text changed. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub URCS_Year_Combobox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles URCS_Year_Combobox.TextChanged
        Dim mDataTable As DataTable

        ' Get the table from SQL
        mDataTable = Get_AAR_Indexes_Table(URCS_Year_Combobox.Text)

        If mDataTable.Rows.Count > 0 Then
            For mLooper = 0 To mDataTable.Rows.Count - 1
                Select Case mDataTable.Rows(mLooper)("region")
                    Case 9 'US
                        TextBox_Fuel_US.Text = mDataTable.Rows(mLooper)("Fuel")
                        TextBox_MS_US.Text = mDataTable.Rows(mLooper)("MS")
                        TextBox_PS_US.Text = mDataTable.Rows(mLooper)("PS")
                        TextBox_Wage_US.Text = mDataTable.Rows(mLooper)("Wage")
                        TextBox_MP_US.Text = mDataTable.Rows(mLooper)("MP")
                    Case 4 'East
                        TextBox_Fuel_East.Text = mDataTable.Rows(mLooper)("Fuel")
                        TextBox_MS_East.Text = mDataTable.Rows(mLooper)("MS")
                        TextBox_PS_East.Text = mDataTable.Rows(mLooper)("PS")
                        TextBox_Wage_East.Text = mDataTable.Rows(mLooper)("Wage")
                        TextBox_MP_East.Text = mDataTable.Rows(mLooper)("MP")
                    Case 7 'West
                        TextBox_Fuel_West.Text = mDataTable.Rows(mLooper)("Fuel")
                        TextBox_MS_West.Text = mDataTable.Rows(mLooper)("MS")
                        TextBox_PS_West.Text = mDataTable.Rows(mLooper)("PS")
                        TextBox_Wage_West.Text = mDataTable.Rows(mLooper)("Wage")
                        TextBox_MP_West.Text = mDataTable.Rows(mLooper)("MP")
                End Select
            Next
        Else
            ' The values for the selected year don't exist - load zeros into the form fields.
            TextBox_Fuel_US.Text = "0.0"
            TextBox_MS_US.Text = "0.0"
            TextBox_PS_US.Text = "0.0"
            TextBox_Wage_US.Text = "0.0"
            TextBox_MP_US.Text = "0.0"
            TextBox_Fuel_East.Text = "0.0"
            TextBox_MS_East.Text = "0.0"
            TextBox_PS_East.Text = "0.0"
            TextBox_Wage_East.Text = "0.0"
            TextBox_MP_East.Text = "0.0"
            TextBox_Fuel_West.Text = "0.0"
            TextBox_MS_West.Text = "0.0"
            TextBox_PS_West.Text = "0.0"
            TextBox_Wage_West.Text = "0.0"
            TextBox_MP_West.Text = "0.0"
        End If

        mDataTable = Nothing

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button execute click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        Const Eastern = 4
        Const Western = 7
        Const National = 9

        Dim mDataTable As DataTable
        Dim mSQLCommand As SqlCommand

        ' Variables for writing to Excel - commented out 10/26/20 by M. Sanders
        'Dim oXL As Excel.Application = Nothing
        'Dim oWBs As Excel.Workbooks = Nothing
        'Dim oWB As Excel.Workbook = Nothing
        'Dim oSheet As Excel.Worksheet = Nothing
        'Dim oCells As Excel.Range = Nothing

        ' Variables for SpreadsheetGear
        Dim mWorkbook As IWorkbook
        Dim mExcelSheet As IWorksheet

        Dim mStrSQL As String, mField As String
        Dim mDataYear As Integer
        Dim mIndex As Integer
        Dim mHistorical_Year As Integer
        Dim mRegion As Integer
        Dim mWriteIt As Integer
        Dim mThisRegion As Integer
        Dim mIndexYear As Integer
        Dim mLooper As Integer, mLine As Integer

        Dim ReportArray(23, 5) As Single

        mWriteIt = vbNo
        mIndex = 0

        'Check to make sure that we have values in the form fields
        If Len(Trim(URCS_Year_Combobox.Text)) = 0 Then
            MsgBox("Error - No Year selected.", MsgBoxStyle.OkOnly, "Error!")
            GoTo EndIt
        End If

        If txt_Report_FilePath.TextLength = 0 Then
            MsgBox("You must select an Report File value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        ' If the output file exists verify that the user wants to overwrite it.  If not, end this.
        If My.Computer.FileSystem.FileExists(txt_Report_FilePath.Text) Then
            If MsgBox("File already exists!  Overwrite?", vbYesNo, "Warning!") = vbYes Then
                My.Computer.FileSystem.DeleteFile(txt_Report_FilePath.Text)
            Else
                GoTo EndIt
            End If
        End If

        'This entry should never need changing
        mHistorical_Year = 1978

        ' Get the location of the Index table
        Gbl_Price_Index_TableName = Get_Table_Name_From_SQL("1", "INDEX")
        Gbl_Price_Index_DatabaseName = Get_Database_Name_From_SQL("1", "INDEX")

        ' Determine if we need to update existing records or insert new ones into the table
        mDataTable = Get_AAR_Indexes_Table(URCS_Year_Combobox.Text)

        If mDataTable.Rows.Count > 0 Then
            ' The year already exists.  Prompt the user to make sure they want to overwrite.
            mWriteIt = MsgBox("This year's data already exists in the database!  Overwrite it?", MsgBoxStyle.YesNo, "WARNING - WARNING - WARNING")
            If mWriteIt = vbYes Then

                ' Overwrite the values in the database for US
                mStrSQL = Build_Update_AARIndex_SQL_Statement(
                    Get_Table_Name_From_SQL("1", "URCS_AAR_INDEXES"),
                    National,
                    URCS_Year_Combobox.Text,
                    TextBox_Fuel_US.Text,
                    TextBox_MS_US.Text,
                    TextBox_PS_US.Text,
                    TextBox_Wage_US.Text,
                    TextBox_MP_US.Text)
                mSQLCommand = New SqlCommand
                mSQLCommand.CommandType = CommandType.Text
                mSQLCommand.CommandText = mStrSQL
                mSQLCommand.Connection = gbl_SQLConnection
                mSQLCommand.ExecuteNonQuery()

                ' Overwrite the values in the database for East
                mStrSQL = Build_Update_AARIndex_SQL_Statement(
                    Get_Table_Name_From_SQL("1", "URCS_AAR_INDEXES"),
                    Eastern,
                    URCS_Year_Combobox.Text,
                    TextBox_Fuel_East.Text,
                    TextBox_MS_East.Text,
                    TextBox_PS_East.Text,
                    TextBox_Wage_East.Text,
                    TextBox_MP_East.Text)
                mSQLCommand = New SqlCommand
                mSQLCommand.CommandType = CommandType.Text
                mSQLCommand.CommandText = mStrSQL
                mSQLCommand.Connection = gbl_SQLConnection
                mSQLCommand.ExecuteNonQuery()

                ' Overwrite the values in the database for West
                mStrSQL = Build_Update_AARIndex_SQL_Statement(
                    Get_Table_Name_From_SQL("1", "URCS_AAR_INDEXES"),
                    Western,
                    URCS_Year_Combobox.Text,
                    TextBox_Fuel_West.Text,
                    TextBox_MS_West.Text,
                    TextBox_PS_West.Text,
                    TextBox_Wage_West.Text,
                    TextBox_MP_West.Text)
                mSQLCommand = New SqlCommand
                mSQLCommand.CommandType = CommandType.Text
                mSQLCommand.CommandText = mStrSQL
                mSQLCommand.Connection = gbl_SQLConnection
                mSQLCommand.ExecuteNonQuery()

                mDataTable = Nothing
            Else
                ' The user opted out.  Exit the routine.
                mDataTable = Nothing
                GoTo EndIt
            End If
        Else
            ' Insert the new year's values to the table

            ' Overwrite the values in the database for US
            mStrSQL = Build_Insert_AARIndex_SQL_statement(
                Get_Table_Name_From_SQL("1", "URCS_AAR_INDEXES"),
                National,
                URCS_Year_Combobox.Text,
                TextBox_Fuel_US.Text,
                TextBox_MS_US.Text,
                TextBox_PS_US.Text,
                TextBox_Wage_US.Text,
                TextBox_MP_US.Text).ToString
            mSQLCommand = New SqlCommand
            mSQLCommand.CommandType = CommandType.Text
            mSQLCommand.CommandText = mStrSQL
            mSQLCommand.Connection = gbl_SQLConnection
            mSQLCommand.ExecuteNonQuery()

            ' Overwrite the values in the database for East
            mStrSQL = Build_Insert_AARIndex_SQL_statement(
                Get_Table_Name_From_SQL("1", "URCS_AAR_INDEXES"),
                Eastern,
                URCS_Year_Combobox.Text,
                TextBox_Fuel_East.Text,
                TextBox_MS_East.Text,
                TextBox_PS_East.Text,
                TextBox_Wage_East.Text,
                TextBox_MP_East.Text).ToString
            mSQLCommand = New SqlCommand
            mSQLCommand.CommandType = CommandType.Text
            mSQLCommand.CommandText = mStrSQL
            mSQLCommand.Connection = gbl_SQLConnection
            mSQLCommand.ExecuteNonQuery()

            ' Overwrite the values in the database for West
            mStrSQL = Build_Insert_AARIndex_SQL_statement(
                Get_Table_Name_From_SQL("1", "URCS_AAR_INDEXES"),
                Western,
                URCS_Year_Combobox.Text,
                TextBox_Fuel_West.Text,
                TextBox_MS_West.Text,
                TextBox_PS_West.Text,
                TextBox_Wage_West.Text,
                TextBox_MP_West.Text).ToString
            mSQLCommand = New SqlCommand
            mSQLCommand.CommandType = CommandType.Text
            mSQLCommand.CommandText = mStrSQL
            mSQLCommand.Connection = gbl_SQLConnection
            mSQLCommand.ExecuteNonQuery()

            mDataTable = Nothing
        End If

        ' Now that we have the data stored/updated, compute the values for the Trans table
        For mDataYear = CInt(URCS_Year_Combobox.Text) To (CInt(URCS_Year_Combobox.Text) - 4) Step -1 ' Calulates Current_Year to Current_Year_Minus 4 value

            ' Let the user know what is going on
            txt_StatusBox.Text = "Processing " & CStr(mDataYear) & "...  Please wait"
            Refresh()

            mIndexYear = URCS_Year_Combobox.Text - mDataYear

            ' Load the variables from the Trans table for the year
            Call OpenTransFile(mDataYear)

            'Process the 2 Regions and National records for the year
            For mRegion = 1 To 3

                ' Let the user know what is going on
                txt_StatusBox.Text = "Processing " & CStr(mDataYear) & "...  Please wait"
                Refresh()

                'Initialize the array
                For mLooper = 1 To 23
                    mIndexes(mIndexYear, mRegion, mLooper) = 1
                Next

                'Calculate index 1 through 5
                Call CalculateIndex1_5(mDataYear, mIndexYear, mRegion)

                'Calculate index 6 - Relative wieghted average for depreciation
                Call CalculateIndex6(mDataYear, mIndexYear, mRegion)

                'Calculate index 7 - Relative weighted average for lease rentals
                Call CalculateIndex7(mDataYear, mIndexYear, mRegion)

                'Calculate index 8 - Relative weighted averages for other expenses
                Call CalculateIndex8(mDataYear, mIndexYear, mRegion)

                'Calculate index 9 - Relative weighted average for other rents
                Call CalculateIndex9(mDataYear, mIndexYear, mRegion)

                'calculate index 10 - Relative weighted average for locomotive
                'repair and maintenance
                Call CalculateIndex10(mDataYear, mIndexYear, mRegion)

                'Calculate index 11 - Relative weighted average for freight car
                'repair and maintenance
                Call CalculateIndex11(mDataYear, mIndexYear, mRegion)

                'Calculate index 12 - Relative weighted average for other equipment
                'repair and maintenance
                Call CalculateIndex12(mDataYear, mIndexYear, mRegion)

                'Calculate index 13 - Relative weighted average for locomotive
                'depreciation
                Call CalculateIndex13(mDataYear, mIndexYear, mRegion)

                'Calculate index 14 - Relative weighted average for freight car
                'depreciation
                Call CalculateIndex14(mDataYear, mIndexYear, mRegion)

                'Calculate index 15 - Relative weighted average for other equipment
                'depreciation
                Call CalculateIndex15(mDataYear, mIndexYear, mRegion)

                'Calculate index 16 - Relative weighted average for locomotive
                'joint facility debt
                Call CalculateIndex16(mDataYear, mIndexYear, mRegion)

                'Calculate index 17 - Relative Weighted average for other equipment
                'lease rentals
                Call CalculateIndex17(mDataYear, mIndexYear, mRegion)

                'Calculate index 18 - Relative weighted average for total other
                'equipment
                Call CalculateIndex18(mDataYear, mIndexYear, mRegion)

                'Indexes 19-21 are all set to the value of Idx3
                mIndexes(mIndexYear, mRegion, 19) = mIndexes(mIndexYear, mRegion, 3)
                mIndexes(mIndexYear, mRegion, 20) = mIndexes(mIndexYear, mRegion, 3)
                mIndexes(mIndexYear, mRegion, 21) = mIndexes(mIndexYear, mRegion, 3)

                'Calculate index 22 - Relative weighted average for specialized
                'service operations
                Call CalculateIndex22(mDataYear, mIndexYear, mRegion)

                'Set idx23 to 1 for all years
                mIndexes(mIndexYear, mRegion, 23) = 1

                For mLooper = 1 To 23

                    ' Determine which Region we're working on
                    Select Case mRegion
                        Case 1
                            mThisRegion = Eastern
                        Case 2
                            mThisRegion = Western
                        Case 3
                            mThisRegion = National
                    End Select

                    ' Create the record does not exist in the SQL table
                    If Count_Price_Index_Records(URCS_Year_Combobox.Text, mLooper, mThisRegion.ToString) = 0 Then
                        Insert_Price_Index_Record(URCS_Year_Combobox.Text, mLooper, mThisRegion, 0, 0, 0, 0, 0)
                    End If

                    mField = ""

                    ' Update the record in the SQL table with the value for this index depending on what year's value we're processing
                    Select Case mIndexYear
                        Case 0
                            mField = "Current_Year"
                        Case 1
                            mField = "Current_Year_Minus_1"
                        Case 2
                            mField = "Current_Year_Minus_2"
                        Case 3
                            mField = "Current_Year_Minus_3"
                        Case 4
                            mField = "Current_Year_Minus_4"
                    End Select

                    Update_Price_Index_Record(URCS_Year_Combobox.Text,
                                              mLooper,
                                              mThisRegion,
                                              mField,
                                              mIndexes(mIndexYear, mRegion, mLooper))
                Next mLooper
            Next mRegion
        Next mDataYear

        ' Now we need to create the processing report

        ' Open the Workbook
        mWorkbook = Factory.GetWorkbook(Globalization.CultureInfo.CurrentCulture)

        'Generate the report
        ' Set the active sheet's name
        If mWorkbook.ActiveSheet.Name <> "Sheet1" Then
            mWorkbook.Worksheets.Add()
        End If

        mWorkbook.ActiveSheet.Name = "Input Data"

        ' Set the data to the Summary page
        mExcelSheet = mWorkbook.ActiveSheet
        mExcelSheet.Cells().Font.Size = 12 ' Sets default font size

        mExcelSheet.Cells("A1").Value = "URCS Indexes Load Process Report - " & URCS_Year_Combobox.Text.ToString
        mExcelSheet.Cells("A1").Font.Bold = True
        mExcelSheet.Cells("A1").Font.Size = 14

        mExcelSheet.Cells("A2").Value = "Input Data"
        mExcelSheet.Cells("A3").Value = "Date/Time: " & CStr(Now).ToString

        mExcelSheet.Cells("A6").Value = "Fuel"
        mExcelSheet.Cells("A7").Value = "Materials & Supply"
        mExcelSheet.Cells("A8").Value = "Purchased Services"
        mExcelSheet.Cells("A9").Value = "Wage Rates & Supplements"
        mExcelSheet.Cells("A10").Value = "Matl Prices & Wage Rates" & vbCrLf & "Combined (incl. Fuel)"

        mExcelSheet.Cells("B5").Value = "National"
        mExcelSheet.Cells("C5").Value = "East"
        mExcelSheet.Cells("D5").Value = "West"

        mExcelSheet.Cells("A1").ColumnWidth = 30
        mExcelSheet.Cells("A10").RowHeight = 30

        mExcelSheet.Cells("B5:D5").HorizontalAlignment = SpreadsheetGear.HAlign.Right
        mExcelSheet.Cells("B5:D5").Font.Bold = True
        mExcelSheet.Cells("B5:D5").Font.Underline = True

        mExcelSheet.Cells("B6").Value = TextBox_Fuel_US.Text
        mExcelSheet.Cells("C6").Value = TextBox_Fuel_East.Text
        mExcelSheet.Cells("D6").Value = TextBox_Fuel_West.Text
        mExcelSheet.Cells("B7").Value = TextBox_MS_US.Text
        mExcelSheet.Cells("C7").Value = TextBox_MS_East.Text
        mExcelSheet.Cells("D7").Value = TextBox_MS_West.Text
        mExcelSheet.Cells("B8").Value = TextBox_PS_US.Text
        mExcelSheet.Cells("C8").Value = TextBox_PS_East.Text
        mExcelSheet.Cells("D8").Value = TextBox_PS_West.Text
        mExcelSheet.Cells("B9").Value = TextBox_Wage_US.Text
        mExcelSheet.Cells("C9").Value = TextBox_Wage_East.Text
        mExcelSheet.Cells("D9").Value = TextBox_Wage_West.Text
        mExcelSheet.Cells("B10").Value = TextBox_MP_US.Text
        mExcelSheet.Cells("C10").Value = TextBox_MP_East.Text
        mExcelSheet.Cells("D10").Value = TextBox_MP_West.Text

        'Add and select the second sheet
        mWorkbook.Worksheets.Add()
        mWorkbook.Worksheets("Sheet2").Select()
        mWorkbook.ActiveSheet.Name = "Results"

        mExcelSheet = mWorkbook.ActiveSheet
        mExcelSheet.Cells().Font.Size = 12 ' Sets default font size

        mExcelSheet.Cells("A1").Value = "URCS Indexes Load Process Report - " & URCS_Year_Combobox.Text.ToString
        mExcelSheet.Cells("A1").Font.Bold = True
        mExcelSheet.Cells("A1").Font.Size = 14

        mExcelSheet.Cells("A2").Value = "Computation Results"
        mExcelSheet.Cells("A3").Value = "Date/Time: " & CStr(Now).ToString

        mExcelSheet.Cells("A5").Value = "Region"
        mExcelSheet.Cells("B5").Value = "Index"
        mExcelSheet.Cells("C5").Value = "Current" & vbCrLf & "Year"
        mExcelSheet.Cells("D5").Value = Val(URCS_Year_Combobox.Text) - 1
        mExcelSheet.Cells("E5").Value = Val(URCS_Year_Combobox.Text) - 2
        mExcelSheet.Cells("F5").Value = Val(URCS_Year_Combobox.Text) - 3
        mExcelSheet.Cells("G5").Value = Val(URCS_Year_Combobox.Text) - 4
        mExcelSheet.Cells("H5").Value = "Use"

        mExcelSheet.Cells("A5:H5").HorizontalAlignment = SpreadsheetGear.HAlign.Left
        mExcelSheet.Cells("A5:H5").Font.Bold = True
        mExcelSheet.Cells("A5:H5").Font.Underline = True

        mExcelSheet.Cells("A6:E60").ColumnWidth = 10
        mExcelSheet.Cells("A6:H60").WrapText = False

        mDataYear = 0
        mLine = 5
        For mRegion = 1 To 3

            ' Determine which Region we're working on
            Select Case mRegion
                Case 1
                    mLine = mLine + 1
                    mExcelSheet.Cells("A" & mLine.ToString).Value = "Eastern"
                Case 2
                    mLine = mLine + 1
                    mExcelSheet.Cells("A" & mLine.ToString).Value = "Western"
                Case 3
                    mLine = mLine + 1
                    mExcelSheet.Cells("A" & mLine.ToString).Value = "National"
            End Select

            For mLooper = 1 To 23

                mExcelSheet.Cells("B" & mLine.ToString).Value = mLooper
                mExcelSheet.Cells("C" & mLine.ToString).Value = mIndexes(mDataYear, mRegion, mLooper)
                mExcelSheet.Cells("D" & mLine.ToString).Value = mIndexes(mDataYear + 1, mRegion, mLooper)
                mExcelSheet.Cells("E" & mLine.ToString).Value = mIndexes(mDataYear + 2, mRegion, mLooper)
                mExcelSheet.Cells("F" & mLine.ToString).Value = mIndexes(mDataYear + 3, mRegion, mLooper)
                mExcelSheet.Cells("G" & mLine.ToString).Value = mIndexes(mDataYear + 4, mRegion, mLooper)
                mExcelSheet.Cells("H" & mLine.ToString).Value = Get_Misc_Report_Line("AAR_Index", mLooper.ToString)
                mExcelSheet.Cells("H" & mLine.ToString).WrapText = False
                mLine = mLine + 1
            Next
        Next

        ' Set the Input Data sheet as default when the file opens in Excel for the first time
        mWorkbook.Worksheets(0).Select()
        mWorkbook.SaveAs(txt_Report_FilePath.Text, FileFormat.OpenXMLWorkbook)

        txt_StatusBox.Text = "Done!"
        Refresh()

EndIt:


    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Opens transaction file. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="Year"> The year. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub OpenTransFile(ByVal Year As Integer)

        Dim mDataTable As New DataTable

        Dim mStrSQL As String
        Dim mRRCount As Integer
        Dim mIndex As Integer
        Dim mRegion As Integer
        Dim mLooper As Integer

        ' Build the connection string
        gbl_Table_Name = Get_Table_Name_From_SQL("1", "URCS_RAILROADS")
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "URCS_RAILROADS")
        OpenSQLConnection(Gbl_Controls_Database_Name)

        mStrSQL = "SELECT RRICC FROM " & gbl_Table_Name & " WHERE RRICC < 200000 AND region = 4"

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        mRRCount = 0
        mRegionCount(1) = mDataTable.Rows.Count

        For mLooper = 0 To mDataTable.Rows.Count - 1
            mRRCount = mRRCount + 1
            mRailroads(mRRCount, 1) = mDataTable.Rows(mLooper)("rricc")
        Next

        'Do it again for the West Railroads
        mDataTable = New DataTable

        mStrSQL = "SELECT RRICC FROM " & gbl_Table_Name & " WHERE RRICC < 200000 AND region = 7"

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        mRRCount = 0
        mRegionCount(2) = mDataTable.Rows.Count

        For mLooper = 0 To mDataTable.Rows.Count - 1
            mRRCount = mRRCount + 1
            mRailroads(mRRCount, 2) = mDataTable.Rows(mLooper)("rricc")
        Next

        'Now get all the info together for the National Railroads
        mRegionCount(3) = mRegionCount(1) + mRegionCount(2)
        mRRCount = 0
        For mRegion = 1 To 2
            For mIndex = 1 To mRegionCount(mRegion)
                mRRCount = mRRCount + 1
                mRailroads(mRRCount, 3) = mRailroads(mIndex, mRegion)
            Next mIndex
        Next mRegion

        mDataTable = Nothing

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 1 5. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub CalculateIndex1_5(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable

        Dim mData As Class_URCS_AARIndex
        Dim mStrSQL As String
        Dim mThisRegion As Integer

        'Determine which region index was passed and convert to standard
        'value used in the database
        mThisRegion = SetRegion(mRegion)

        ' Build the connection string
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "URCS_AAR_INDEXES")
        OpenSQLConnection(gbl_Database_Name)

        mStrSQL = Build_Select_AARIndex_By_Region_SQL_Statement(mThisRegion, mDataYear)

        ' Fill the datatable from SQL
        Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
            daAdapter.Fill(mDataTable)
        End Using

        If mDataYear = URCS_Year_Combobox.Text Then
            'Set the base level indicies
            mBase(mRegion) = New Class_URCS_AARIndex
            With mBase(mRegion)
                .Fuel = mDataTable.Rows(0)("fuel")
                .MP = mDataTable.Rows(0)("mp")
                .MS = mDataTable.Rows(0)("ms")
                .PS = mDataTable.Rows(0)("ps")
                .Wage = mDataTable.Rows(0)("wage")
                .Region = mDataTable.Rows(0)("region")
                .Year = mDataTable.Rows(0)("year")
            End With
        End If

        'Get indicies for this year
        mData = New Class_URCS_AARIndex
        With mData
            .Fuel = mDataTable.Rows(0)("fuel")
            .MP = mDataTable.Rows(0)("mp")
            .MS = mDataTable.Rows(0)("ms")
            .PS = mDataTable.Rows(0)("ps")
            .Wage = mDataTable.Rows(0)("wage")
            .Region = mDataTable.Rows(0)("region")
            .Year = mDataTable.Rows(0)("year")

            'Calculate deflated current year index level relative to base year
            .Fuel = Divide(mBase(mRegion).Fuel, .Fuel)
            .MS = Divide(mBase(mRegion).MS, .MS)
            .PS = Divide(mBase(mRegion).PS, .PS)
            .Wage = Divide(mBase(mRegion).Wage, .Wage)
            .MP = Divide(mBase(mRegion).MP, .MP)
        End With

        'Put the first five indicies to the Current Index class
        mIndexes(mIndexYear, mRegion, 1) = mData.Wage
        mIndexes(mIndexYear, mRegion, 2) = mData.MS
        mIndexes(mIndexYear, mRegion, 3) = mData.MP
        mIndexes(mIndexYear, mRegion, 4) = mData.PS
        mIndexes(mIndexYear, mRegion, 5) = mData.Fuel

        'Clean up
        mDataTable = Nothing

    End Sub ' CalculateIndex1_5

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 6. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex6(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable

        Dim accumulate As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer
        Dim mLooper As Integer

        accumulate = New Class_URCS_AARIndex

        ' Build the connection string
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenSQLConnection(gbl_Database_Name)

        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "136", "138")

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    accumulate.Wage = accumulate.Wage + mDataTable.Rows(mLooper)("c1")
                    accumulate.MS = accumulate.MS + mDataTable.Rows(mLooper)("c2")
                    accumulate.PS = accumulate.PS + mDataTable.Rows(mLooper)("c3")
                    accumulate.Gen = accumulate.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next

        'Clean up
        mDataTable = Nothing

        mUnweighted = accumulate.Wage +
            accumulate.MS +
            accumulate.PS +
            accumulate.Gen

        mWeighted = (accumulate.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (accumulate.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (accumulate.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (accumulate.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 6) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 6) = 1
        End If

    End Sub ' CalculateIndex6

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 7. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex7(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable

        Dim trans As Class_URCS_AARIndex
        Dim LeaseRental_C As Class_URCS_AARIndex
        Dim LeaseRental_D As Class_URCS_AARIndex
        Dim OtherRents_C As Class_URCS_AARIndex
        Dim OtherRents_D As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer
        Dim mLooper As Integer

        trans = New Class_URCS_AARIndex
        LeaseRental_C = New Class_URCS_AARIndex
        LeaseRental_D = New Class_URCS_AARIndex
        OtherRents_C = New Class_URCS_AARIndex
        OtherRents_D = New Class_URCS_AARIndex

        ' Open the SQL connection using the global variable holding the connection string
        'Get the database_name and table_name for the URCS_Years value from the database

        ' Build the connection string
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenSQLConnection(gbl_Database_Name)

        'Get the leaserental_C data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "121", "123")

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    LeaseRental_C.Wage = LeaseRental_C.Wage + mDataTable.Rows(mLooper)("c1")
                    LeaseRental_C.MS = LeaseRental_C.MS + mDataTable.Rows(mLooper)("c2")
                    LeaseRental_C.PS = LeaseRental_C.PS + mDataTable.Rows(mLooper)("c3")
                    LeaseRental_C.Gen = LeaseRental_C.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        'Get the leaserental_D data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "118", "120")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    LeaseRental_D.Wage = LeaseRental_D.Wage + mDataTable.Rows(mLooper)("c1")
                    LeaseRental_D.MS = LeaseRental_D.MS + mDataTable.Rows(mLooper)("c2")
                    LeaseRental_D.PS = LeaseRental_D.PS + mDataTable.Rows(mLooper)("c3")
                    LeaseRental_D.Gen = LeaseRental_D.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        'Get the otherrents_C data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "133", "135")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    OtherRents_C.Wage = OtherRents_C.Wage + mDataTable.Rows(mLooper)("c1")
                    OtherRents_C.MS = OtherRents_C.MS + mDataTable.Rows(mLooper)("c2")
                    OtherRents_C.PS = OtherRents_C.PS + mDataTable.Rows(mLooper)("c3")
                    OtherRents_C.Gen = OtherRents_C.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        'Get the otherrents_D data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "130", "132")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    OtherRents_D.Wage = OtherRents_D.Wage + mDataTable.Rows(mLooper)("c1")
                    OtherRents_D.MS = OtherRents_D.MS + mDataTable.Rows(mLooper)("c2")
                    OtherRents_D.PS = OtherRents_D.PS + mDataTable.Rows(mLooper)("c3")
                    OtherRents_D.Gen = OtherRents_D.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        'combine the results
        trans.Wage = LeaseRental_C.Wage +
            LeaseRental_D.Wage +
            OtherRents_C.Wage +
            OtherRents_D.Wage

        trans.MS = LeaseRental_C.MS +
            LeaseRental_D.MS +
            OtherRents_C.MS +
            OtherRents_D.MS

        trans.Gen = LeaseRental_C.Gen +
            LeaseRental_D.Gen +
            OtherRents_C.Gen +
            OtherRents_D.Gen

        trans.PS = LeaseRental_C.PS +
            LeaseRental_D.PS +
            OtherRents_C.PS +
            OtherRents_D.PS

        mUnweighted = trans.Wage +
            trans.MS +
            trans.PS +
            trans.Gen

        mWeighted = (trans.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (trans.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (trans.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (trans.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 7) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 7) = 1
        End If

    End Sub ' CalculateIndex7

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 8. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex8(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable

        Dim accumulate As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer

        accumulate = New Class_URCS_AARIndex

        ' Build the connection string
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenSQLConnection(gbl_Database_Name)

        'Get the data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "148", "150")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    accumulate.Wage = accumulate.Wage + mDataTable.Rows(mLooper)("c1")
                    accumulate.MS = accumulate.MS + mDataTable.Rows(mLooper)("c2")
                    accumulate.PS = accumulate.PS + mDataTable.Rows(mLooper)("c3")
                    accumulate.Gen = accumulate.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        mUnweighted = accumulate.Wage +
            accumulate.MS +
            accumulate.PS +
            accumulate.Gen

        mWeighted = (accumulate.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (accumulate.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (accumulate.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (accumulate.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 8) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 8) = 1
        End If

    End Sub ' CalulateIndex8

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 9. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex9(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable
        Dim accumulate As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer

        accumulate = New Class_URCS_AARIndex

        ' Open the SQL connection
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenSQLConnection(gbl_Database_Name)

        'Get the data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "231")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    accumulate.Wage = accumulate.Wage + mDataTable.Rows(mLooper)("c1")
                    accumulate.MS = accumulate.MS + mDataTable.Rows(mLooper)("c2")
                    accumulate.PS = accumulate.PS + mDataTable.Rows(mLooper)("c3")
                    accumulate.Gen = accumulate.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        mUnweighted = accumulate.Wage +
            accumulate.MS +
            accumulate.PS +
            accumulate.Gen

        mWeighted = (accumulate.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (accumulate.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (accumulate.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (accumulate.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 9) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 9) = 1
        End If

    End Sub ' CalculateIndex9

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 10. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex10(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable
        Dim repair As Class_URCS_AARIndex
        Dim billedrepair As Class_URCS_AARIndex
        Dim netrepair As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer

        repair = New Class_URCS_AARIndex
        billedrepair = New Class_URCS_AARIndex
        netrepair = New Class_URCS_AARIndex

        ' Open the SQL connection
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenSQLConnection(gbl_Database_Name)

        'Get the repair data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "202", "203")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    repair.Wage = repair.Wage + mDataTable.Rows(mLooper)("c1")
                    repair.MS = repair.MS + mDataTable.Rows(mLooper)("c2")
                    repair.PS = repair.PS + mDataTable.Rows(mLooper)("c3")
                    repair.Gen = repair.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        'Get the billed repair data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "216")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    billedrepair.Wage = billedrepair.Wage + mDataTable.Rows(mLooper)("c1")
                    billedrepair.MS = billedrepair.MS + mDataTable.Rows(mLooper)("c2")
                    billedrepair.PS = billedrepair.PS + mDataTable.Rows(mLooper)("c3")
                    billedrepair.Gen = billedrepair.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        netrepair.Wage = repair.Wage - billedrepair.Wage
        netrepair.MS = repair.MS - billedrepair.MS
        netrepair.PS = repair.PS - billedrepair.PS
        netrepair.Gen = repair.Gen - billedrepair.Gen

        mUnweighted = netrepair.Wage +
            netrepair.MS +
            netrepair.PS +
            netrepair.Gen

        mWeighted = (netrepair.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (netrepair.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (netrepair.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (netrepair.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 10) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 10) = 1
        End If

    End Sub ' CalculateIndex10

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 11. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex11(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable
        Dim repair As Class_URCS_AARIndex
        Dim billedrepair As Class_URCS_AARIndex
        Dim netrepair As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer

        repair = New Class_URCS_AARIndex
        billedrepair = New Class_URCS_AARIndex
        netrepair = New Class_URCS_AARIndex

        ' Open the SQL connection
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenSQLConnection(gbl_Database_Name)

        'Get the repair data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "221", "222")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    repair.Wage = repair.Wage + mDataTable.Rows(mLooper)("c1")
                    repair.MS = repair.MS + mDataTable.Rows(mLooper)("c2")
                    repair.PS = repair.PS + mDataTable.Rows(mLooper)("c3")
                    repair.Gen = repair.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        'Get the billed repair data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "235")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    billedrepair.Wage = billedrepair.Wage + mDataTable.Rows(mLooper)("c1")
                    billedrepair.MS = billedrepair.MS + mDataTable.Rows(mLooper)("c2")
                    billedrepair.PS = billedrepair.PS + mDataTable.Rows(mLooper)("c3")
                    billedrepair.Gen = billedrepair.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        netrepair.Wage = repair.Wage - billedrepair.Wage
        netrepair.MS = repair.MS - billedrepair.MS
        netrepair.PS = repair.PS - billedrepair.PS
        netrepair.Gen = repair.Gen - billedrepair.Gen

        mUnweighted = netrepair.Wage +
            netrepair.MS +
            netrepair.PS +
            netrepair.Gen

        mWeighted = (netrepair.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (netrepair.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (netrepair.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (netrepair.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 11) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 11) = 1
        End If

    End Sub ' CalculateIndex11

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 12. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex12(
        ByVal mDataYear As Integer,
        ByVal mIndexYear As Integer,
        ByVal mRegion As Integer)

        Dim mDataTable As New DataTable
        Dim repair As Class_URCS_AARIndex
        Dim billedrepair As Class_URCS_AARIndex
        Dim netrepair As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer

        repair = New Class_URCS_AARIndex
        billedrepair = New Class_URCS_AARIndex
        netrepair = New Class_URCS_AARIndex

        ' Open the SQL connection
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenSQLConnection(gbl_Database_Name)

        'Get the repair data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "302", "307")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    repair.Wage = repair.Wage + mDataTable.Rows(mLooper)("c1")
                    repair.MS = repair.MS + mDataTable.Rows(mLooper)("c2")
                    repair.PS = repair.PS + mDataTable.Rows(mLooper)("c3")
                    repair.Gen = repair.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        'Get the billed repair data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "320")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    billedrepair.Wage = billedrepair.Wage + mDataTable.Rows(mLooper)("c1")
                    billedrepair.MS = billedrepair.MS + mDataTable.Rows(mLooper)("c2")
                    billedrepair.PS = billedrepair.PS + mDataTable.Rows(mLooper)("c3")
                    billedrepair.Gen = billedrepair.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        netrepair.Wage = repair.Wage - billedrepair.Wage
        netrepair.MS = repair.MS - billedrepair.MS
        netrepair.PS = repair.PS - billedrepair.PS
        netrepair.Gen = repair.Gen - billedrepair.Gen

        mUnweighted = netrepair.Wage +
            netrepair.MS +
            netrepair.PS +
            netrepair.Gen

        mWeighted = (netrepair.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (netrepair.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (netrepair.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (netrepair.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 12) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 12) = 1
        End If

    End Sub ' CalculateIndex12

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 13. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex13(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable
        Dim accumulate As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer

        accumulate = New Class_URCS_AARIndex

        ' Open the SQL connection
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenADOConnection(gbl_Database_Name)

        'Get the data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "213")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    accumulate.Wage = accumulate.Wage + mDataTable.Rows(mLooper)("c1")
                    accumulate.MS = accumulate.MS + mDataTable.Rows(mLooper)("c2")
                    accumulate.PS = accumulate.PS + mDataTable.Rows(mLooper)("c3")
                    accumulate.Gen = accumulate.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        mUnweighted = accumulate.Wage +
            accumulate.MS +
            accumulate.PS +
            accumulate.Gen

        mWeighted = (accumulate.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (accumulate.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (accumulate.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (accumulate.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 13) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 13) = 1
        End If

    End Sub ' CalculateIndex13

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 14. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex14(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable
        Dim accumulate As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer

        accumulate = New Class_URCS_AARIndex

        ' Open the SQL connection
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenSQLConnection(gbl_Database_Name)

        'Get the data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "232")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    accumulate.Wage = accumulate.Wage + mDataTable.Rows(mLooper)("c1")
                    accumulate.MS = accumulate.MS + mDataTable.Rows(mLooper)("c2")
                    accumulate.PS = accumulate.PS + mDataTable.Rows(mLooper)("c3")
                    accumulate.Gen = accumulate.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        mUnweighted = accumulate.Wage +
            accumulate.MS +
            accumulate.PS +
            accumulate.Gen

        mWeighted = (accumulate.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (accumulate.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (accumulate.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (accumulate.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 14) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 14) = 1
        End If

    End Sub ' CalculateIndex14

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 15. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex15(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable
        Dim accumulate As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer

        accumulate = New Class_URCS_AARIndex

        ' Open the SQL connection
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenSQLConnection(gbl_Database_Name)

        'Get the data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "317")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    accumulate.Wage = accumulate.Wage + mDataTable.Rows(mLooper)("c1")
                    accumulate.MS = accumulate.MS + mDataTable.Rows(mLooper)("c2")
                    accumulate.PS = accumulate.PS + mDataTable.Rows(mLooper)("c3")
                    accumulate.Gen = accumulate.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        mUnweighted = accumulate.Wage +
            accumulate.MS +
            accumulate.PS +
            accumulate.Gen

        mWeighted = (accumulate.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (accumulate.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (accumulate.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (accumulate.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 15) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 15) = 1
        End If

    End Sub ' CalculateIndex15

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 16. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex16(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable
        Dim accumulate As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer

        accumulate = New Class_URCS_AARIndex

        ' Open the SQL connection
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenADOConnection(gbl_Database_Name)

        'Get the data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "218")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    accumulate.Wage = accumulate.Wage + mDataTable.Rows(mLooper)("c1")
                    accumulate.MS = accumulate.MS + mDataTable.Rows(mLooper)("c2")
                    accumulate.PS = accumulate.PS + mDataTable.Rows(mLooper)("c3")
                    accumulate.Gen = accumulate.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        mUnweighted = accumulate.Wage +
            accumulate.MS +
            accumulate.PS +
            accumulate.Gen

        mWeighted = (accumulate.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (accumulate.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (accumulate.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (accumulate.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 16) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 16) = 1
        End If

    End Sub ' CalculateIndex16

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 17. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex17(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable
        Dim accumulate As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long

        accumulate = New Class_URCS_AARIndex

        ' Open the SQL connection
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenADOConnection(gbl_Database_Name)

        'Get the data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "237")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    accumulate.Wage = accumulate.Wage + mDataTable.Rows(mLooper)("c1")
                    accumulate.MS = accumulate.MS + mDataTable.Rows(mLooper)("c2")
                    accumulate.PS = accumulate.PS + mDataTable.Rows(mLooper)("c3")
                    accumulate.Gen = accumulate.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        mUnweighted = accumulate.Wage +
            accumulate.MS +
            accumulate.PS +
            accumulate.Gen

        mWeighted = (accumulate.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (accumulate.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (accumulate.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (accumulate.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 17) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 17) = 1
        End If

    End Sub ' CalculateIndex17

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 18. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex18(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable
        Dim accumulate As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer

        accumulate = New Class_URCS_AARIndex

        ' Open the SQL connection
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenADOConnection(gbl_Database_Name)

        'Get the data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "322")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    accumulate.Wage = accumulate.Wage + mDataTable.Rows(mLooper)("c1")
                    accumulate.MS = accumulate.MS + mDataTable.Rows(mLooper)("c2")
                    accumulate.PS = accumulate.PS + mDataTable.Rows(mLooper)("c3")
                    accumulate.Gen = accumulate.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        mUnweighted = accumulate.Wage +
            accumulate.MS +
            accumulate.PS +
            accumulate.Gen

        mWeighted = (accumulate.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (accumulate.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (accumulate.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (accumulate.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 18) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 18) = 1
        End If

    End Sub ' CalculateIndex18

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Calculates the index 22. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="mDataYear">    The data year. </param>
    ''' <param name="mIndexYear">   The index year. </param>
    ''' <param name="mRegion">      The region. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Sub CalculateIndex22(
    ByVal mDataYear As Integer,
    ByVal mIndexYear As Integer,
    ByVal mRegion As Integer)

        Dim mDataTable As New DataTable
        Dim spcservice As Class_URCS_AARIndex
        Dim jctfacilitycr As Class_URCS_AARIndex
        Dim other As Class_URCS_AARIndex
        Dim total As Class_URCS_AARIndex
        Dim mUnweighted As Double
        Dim mWeighted As Double
        Dim mStrSQL As String
        Dim mRailroad As Long
        Dim mCounter As Integer

        spcservice = New Class_URCS_AARIndex
        jctfacilitycr = New Class_URCS_AARIndex
        other = New Class_URCS_AARIndex
        total = New Class_URCS_AARIndex

        ' Open the SQL connection 
        gbl_Database_Name = Get_Database_Name_From_SQL("1", "TRANS")
        OpenADOConnection(gbl_Database_Name)

        'Get the repair data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "507", "514")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    spcservice.Wage = spcservice.Wage + mDataTable.Rows(mLooper)("c1")
                    spcservice.MS = spcservice.MS + mDataTable.Rows(mLooper)("c2")
                    spcservice.PS = spcservice.PS + mDataTable.Rows(mLooper)("c3")
                    spcservice.Gen = spcservice.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        'Get the jctfacilitycr data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "515")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    jctfacilitycr.Wage = jctfacilitycr.Wage + mDataTable.Rows(mLooper)("c1")
                    jctfacilitycr.MS = jctfacilitycr.MS + mDataTable.Rows(mLooper)("c2")
                    jctfacilitycr.PS = jctfacilitycr.PS + mDataTable.Rows(mLooper)("c3")
                    jctfacilitycr.Gen = jctfacilitycr.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        'Get the other data for each railroad in the region
        For mCounter = 1 To mRegionCount(mRegion)
            mRailroad = mRailroads(mCounter, mRegion)
            mStrSQL = Build_Select_Trans_SQL_Statement(CStr(mDataYear), CStr(mRailroad), "412", "516")
            mDataTable = New DataTable

            ' Fill the datatable from SQL
            Using daAdapter As New SqlDataAdapter(mStrSQL, gbl_SQLConnection)
                daAdapter.Fill(mDataTable)
            End Using

            If mDataTable.Rows.Count > 0 Then
                For mLooper = 0 To mDataTable.Rows.Count - 1
                    other.Wage = other.Wage + mDataTable.Rows(mLooper)("c1")
                    other.MS = other.MS + mDataTable.Rows(mLooper)("c2")
                    other.PS = other.PS + mDataTable.Rows(mLooper)("c3")
                    other.Gen = other.Gen + mDataTable.Rows(mLooper)("c4")
                Next
            End If
        Next mCounter

        mDataTable = Nothing

        total.Wage = spcservice.Wage + other.Wage
        total.MS = spcservice.MS + other.MS
        total.PS = spcservice.PS + other.PS - jctfacilitycr.PS
        total.Gen = spcservice.Gen + other.Gen

        mUnweighted = total.Wage +
            total.MS +
            total.PS +
            total.Gen

        mWeighted = (total.Wage * mIndexes(mIndexYear, mRegion, 1)) +
            (total.MS * mIndexes(mIndexYear, mRegion, 2)) +
            (total.PS * mIndexes(mIndexYear, mRegion, 3)) +
            (total.Gen * mIndexes(mIndexYear, mRegion, 4))

        If Divide(mWeighted, mUnweighted) <> 0 Then
            mIndexes(mIndexYear, mRegion, 22) = Divide(mWeighted, mUnweighted)
        Else
            mIndexes(mIndexYear, mRegion, 22) = 1
        End If

    End Sub ' CalculateIndex22

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box fuel us leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Fuel_US_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Fuel_US.Leave
        If TextBox_Fuel_US.Text = "" Then
            TextBox_Fuel_US.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box fuel east leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Fuel_East_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Fuel_East.Leave
        If TextBox_Fuel_East.Text = "" Then
            TextBox_Fuel_East.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box fuel west leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Fuel_West_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Fuel_West.Leave
        If TextBox_Fuel_West.Text = "" Then
            TextBox_Fuel_West.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box milliseconds us leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MS_US_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MS_US.Leave
        If TextBox_MS_US.Text = "" Then
            TextBox_MS_US.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box milliseconds east leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MS_East_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MS_East.Leave
        If TextBox_MS_East.Text = "" Then
            TextBox_MS_East.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box milliseconds west leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MS_West_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MS_West.Leave
        If TextBox_MS_West.Text = "" Then
            TextBox_MS_West.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box ps us leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_PS_US_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_PS_US.Leave
        If TextBox_PS_US.Text = "" Then
            TextBox_PS_US.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box ps east leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_PS_East_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_PS_East.Leave
        If TextBox_PS_East.Text = "" Then
            TextBox_PS_East.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box ps west leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_PS_West_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_PS_West.Leave
        If TextBox_PS_West.Text = "" Then
            TextBox_PS_West.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box wage us leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Wage_US_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Wage_US.Leave
        If TextBox_Wage_US.Text = "" Then
            TextBox_Wage_US.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box wage east leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Wage_East_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Wage_East.Leave
        If TextBox_Wage_East.Text = "" Then
            TextBox_Wage_East.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box wage west leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_Wage_West_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Wage_West.Leave
        If TextBox_Wage_West.Text = "" Then
            TextBox_Wage_West.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box mp us leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MP_US_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MP_US.Leave
        If TextBox_MP_US.Text = "" Then
            TextBox_MP_US.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box mp east leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MP_East_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MP_East.Leave
        If TextBox_MP_East.Text = "" Then
            TextBox_MP_East.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Text box mp west leave. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub TextBox_MP_West_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_MP_West.Leave
        If TextBox_MP_West.Text = "" Then
            TextBox_MP_West.Text = "0.0"
        End If
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button report file entry click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/26/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Report_File_Entry_Click(sender As System.Object, e As System.EventArgs) Handles btn_Report_File_Entry.Click
        Dim fd As New FolderBrowserDialog

        If URCS_Year_Combobox.Text = "" Then
            MsgBox("You must select a year first.", vbOKOnly, "Error!")
        Else
            fd.Description = "Select the location in which you want the output report placed."

            If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txt_Report_FilePath.Text = fd.SelectedPath.ToString & "\WB" & URCS_Year_Combobox.Text & " AAR Index Load Report.xlsx"
            End If
        End If


    End Sub

End Class