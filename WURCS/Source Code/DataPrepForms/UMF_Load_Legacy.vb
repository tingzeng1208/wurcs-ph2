Imports System.Data.SqlClient
Imports System.Text
Public Class UMF_Load_Legacy

    Private Sub btn_Output_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New FolderBrowserDialog

        fd.Description = "Select the location in which you want the output file placed."

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.txt_Output_FilePath.Text = fd.SelectedPath.ToString & "\URCUMF." & Me.cmb_URCS_Year.Text
        End If
    End Sub

    Private Sub UMF_Load_Legacy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        'Set the default values for the checkboxes
        Me.chk_Created_Blank_Owner_Records.Checked = False
        Me.chk_Created_System_Records.Checked = False
        Me.chk_Trans_Table_Update.Checked = False

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        mDataTable = Get_URCS_Years_Table()

        For mLooper = 1 To mDataTable.Rows.Count
            cmb_URCS_Year.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
        Next

        mDataTable = Nothing

    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click
        'Check to ensure that the user has selected a directory for the UMF File
        If Me.txt_Output_FilePath.Text = "" Then

            'Ignore the click
        Else

            'Create the UMF File
            CreateUMFFile()

            Me.txt_StatusBox.Text = "Done!"
            Me.Refresh()
        End If
    End Sub

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close this Form
        Me.Close()
    End Sub

    Sub CreateUMFFile()

        Dim rst As ADODB.Recordset
        Dim mString As StringBuilder

        Dim mStrSQL As String, mOutLine As String, EFMT As String
        Dim mRRICC(20) As Decimal, mRegion(20) As Integer, mRRName(20) As String
        Dim mTotalRecords As Integer, mThisRailroad As Integer, RecordsModded As Integer
        Dim mOwner As Integer, mLooper As Integer, mLooper2 As Integer
        Dim mFirstYear As Integer, mWorkRecord As Integer, mWorkYear As Integer
        Dim TrnDat(2650, 5) As Double, mValue As Double, mPrintCheck As Boolean
        Dim fs, outfile

        RecordsModded = 0

        ' Lock the exit button.
        Me.btn_Execute.Visible = False
        Me.Refresh()

        ' Open the SQL connection using the global variable holding the connection string
        txt_StatusBox.Text = "Connecting to database..."
        OpenADOConnection(Get_Database_Name_From_SQL("1", "TRANS"))

        'Process Trans Table Updates
        RecordsModded = TransMod(Me.cmb_URCS_Year.Text)

        If RecordsModded = 0 Then
            MsgBox("Warning - Possible Error: No Trans records Modified", vbOKOnly, "WARNING")
        End If

        Me.chk_Created_Blank_Owner_Records.Checked = True
        Me.chk_Created_Blank_Owner_Records.Refresh()
        System.Windows.Forms.Application.DoEvents()

        'open the output file for writing
        fs = CreateObject("Scripting.FileSystemObject")
        ' If the file exists, we'll overwrite it.
        outfile = fs.createtextfile(Me.txt_Output_FilePath.Text, True)

        'Count the number of railroads to process
        mStrSQL = Build_Count_Railroads_SQL()
        rst = SetRST()
        rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)
        mTotalRecords = rst.Fields("MyCount").Value
        rst.Close()
        rst = Nothing

        mStrSQL = ""
        txt_StatusBox.Text = "Processing..."
        System.Windows.Forms.Application.DoEvents()

        'Get the railroad information and save it to the arrays for use
        'and write the output file
        mStrSQL = Build_Select_Railroads_By_URCSCode_SQL()
        rst = SetRST()
        rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)

        mOutLine = Me.cmb_URCS_Year.Text & " Surface Transportation Board Data " &
        CStr(mTotalRecords) & " Railroads to Process  " & Date.Today
        outfile.writeline(mOutLine)

        mLooper = 0

        Do While Not rst.EOF
            mLooper = mLooper + 1
            mString = New StringBuilder
            mRegion(mLooper) = rst.Fields("region").Value
            mRRICC(mLooper) = rst.Fields("rricc").Value
            mRRName(mLooper) = rst.Fields("name").Value
            mString.Append(Format(rst.Fields("region").Value, "00 "))
            mString.Append(rst.Fields("name").Value)
            outfile.writeline(mString.ToString)
            rst.MoveNext()
        Loop

        'clean up
        rst.Close()
        rst = Nothing

        'now we can create the Owner0 Records from the Dictionary
        mOwner = 0

        'Open a recordset for the Data Dictionary
        rst = SetRST()

        mStrSQL = Build_Select_Dictionary_SQL()

        rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)

        rst.MoveFirst()
        Do Until rst.EOF
            mString = New StringBuilder
            mString.Append(Format(mOwner, "00"))
            mString.Append(Field_Left(rst.Fields("urcsid").Value, 8))
            mString.Append(Field_Left(rst.Fields("loadcode").Value, 1))
            mString.Append(Field_Left(rst.Fields("acct").Value, 8))
            mString.Append(Field_Left(rst.Fields("wtall").Value, 12))
            mString.Append(rst.Fields("name").Value)
            outfile.writeline(mString.ToString)
            rst.MoveNext()
        Loop

        mOutLine = "99       00                    DUMMY CONTROL RECORD"
        outfile.writeline(mOutLine)

        Me.chk_Created_System_Records.Checked = True
        Me.chk_Created_System_Records.Refresh()
        rst.Close()

        'Now we create the owner1 records using the Dictionary
        mStrSQL = Build_Select_Dictionary_SQL("loadcode < 6")
        rst = SetRST()
        rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)
        mOwner = 1
        mValue = 0
        EFMT = " .00000000E+00"

        rst.MoveFirst()
        Do Until rst.EOF
            mString = New StringBuilder
            'Record owner
            mString.Append(Format(mOwner, "00"))
            'Numeric worktable account code
            mString.Append(Field_Left(rst.Fields("urcsid").Value, 8))
            'Print dummy value
            For mLooper = 1 To 5
                mString.Append(Format(mValue, EFMT))
            Next mLooper
            'URCS Index ID Code
            mString.Append(Field_Right(rst.Fields("index").Value, 3))
            'URCS Account Annualization Period
            mString.Append(Field_Left(rst.Fields("annperiod").Value, 1, True))
            'URCS Account assumed sign
            mString.Append(Field_Left(rst.Fields("sign").Value, 1, True))
            'URCS Account Compilation Code
            mString.Append(Field_Left(rst.Fields("comp").Value, 1, True))
            'URCS Account load code
            mString.Append(Field_Left(rst.Fields("loadcode").Value, 1, True))
            'R1 Account Group Code
            mString.Append(Field_Left(rst.Fields("acct").Value, 8))
            'Worktable Account Code
            mString.Append(Field_Left(rst.Fields("wtall").Value, 12))
            'URCS Accumulation Code
            mString.Append(Field_Left(rst.Fields("accumcode").Value, 1))
            outfile.writeline(mString.ToString)
            rst.MoveNext()
        Loop

        Me.chk_Created_Blank_Owner_Records.Checked = True
        Me.Refresh()

        rst.Close()
        rst = Nothing

        'Now we can process the Railroad records
        mFirstYear = Val(Me.cmb_URCS_Year.Text) - 4

        'Cycle through for each railroad, region and national record
        For mThisRailroad = 1 To mTotalRecords

            'Clear the data array
            For mWorkRecord = 1 To 2650
                For mWorkYear = 1 To 5
                    TrnDat(mWorkRecord, mWorkYear) = 0
                Next mWorkYear
            Next mWorkRecord

            mWorkYear = 0
            rst = SetRST()

            'Process each year's trans table records
            For mLooper = Val(Me.cmb_URCS_Year.Text) To mFirstYear Step -1

                Me.txt_StatusBox.Text = mRRName(mThisRailroad)
                Me.txt_Processing_Year.Text = Str(mLooper)
                Me.Refresh()

                mWorkYear = mWorkYear + 1

                ' Build the SQL statement for the year to process

                mStrSQL = Build_Select_Dictionary_SQL(mRRICC(mThisRailroad))
                rst = SetRST()
                rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)

                mWorkRecord = 1
                Do Until rst.EOF
                    TrnDat(mWorkRecord, mWorkYear) =
                        Select_Scaled_Trans_Value(
                            mLooper,
                            mRRICC(mThisRailroad),
                            rst.Fields("sch").Value,
                            rst.Fields("line").Value,
                            rst.Fields("column").Value,
                            rst.Fields("scaler").Value)
                    mWorkRecord = mWorkRecord + 1
                    rst.MoveNext()
                Loop

                rst.Close()
            Next

            'We now have data for 5 years - write it to output table
            mStrSQL = Build_Select_Dictionary_SQL(mRRICC(mThisRailroad))
            rst = SetRST()
            rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)

            mWorkRecord = 1
            Do Until rst.EOF
                mOutLine = ""
                mPrintCheck = False
                For mLooper2 = 1 To 5
                    If Math.Abs(TrnDat(mWorkRecord, mLooper2)) > 0 Then
                        mPrintCheck = True
                    End If
                Next mLooper2
                If mPrintCheck = True Then
                    mString = New StringBuilder
                    'Record Owner
                    mString.Append(Field_Left(Set_URCS_Code(mRRICC(mThisRailroad)), 2))
                    'worktable account code
                    mString.Append(Field_Left(rst.Fields("urcsid").Value, 8))
                    'print the values from the array that we loaded from trans table
                    For mLooper2 = 1 To 5
                        mString.Append(Format(TrnDat(mWorkRecord, mLooper2), " .00000000E+00;-.00000000E+00"))
                    Next mLooper2
                    'URCS Index ID Code
                    mString.Append(Format(rst.Fields("index").Value, "000"))
                    'Annualization Period
                    mString.Append(Format(rst.Fields("annperiod").Value, "0"))
                    'Assumed sign
                    mString.Append(Format(rst.Fields("sign").Value, "0"))
                    'Compulation Code
                    mString.Append(Format(rst.Fields("comp").Value, "0"))
                    'Load Code
                    mString.Append(Format(rst.Fields("loadcode").Value, "0"))
                    'R1 Group code
                    mString.Append(Field_Left(rst.Fields("acct").Value, 8))
                    'worktable accountcode
                    mString.Append(Field_Left(rst.Fields("wtall").Value, 12))
                    'accumulation code
                    mString.Append(Format(rst.Fields("accumcode").Value, "0"))
                    outfile.writeline(mString.ToString)
                End If
                mWorkRecord = mWorkRecord + 1
                rst.MoveNext()
            Loop
            rst.Close()
        Next

        rst = Nothing

        ' Unlock the Exit button
        Me.btn_Execute.Visible = True
        Me.Refresh()

    End Sub

    Function TransMod(ByVal mYear As Integer) As Integer

        Dim rst As ADODB.Recordset
        Dim mStrSQL As String
        Dim mFirstYear As Integer, mLooper As Integer
        Dim mC8 As Decimal, mC9 As Decimal, mC10 As Decimal
        Dim mC12 As Decimal, mC13 As Decimal, mC14 As Decimal
        Dim mIsDirty As Boolean

        TransMod = 0

        Me.txt_StatusBox.Text = "Processing Trans Table Updates..."
        Me.Refresh()

        ' Open the SQL connection using the global variable holding the connection string
        OpenADOConnection(Get_Database_Name_From_SQL("1", "Trans"))

        mFirstYear = mYear - 4
        'Process each year's trans table records
        For mLooper = mFirstYear To mYear

            Me.txt_Processing_Year.Text = Str(mLooper)
            Me.Refresh()

            rst = SetRST()

            'Transform QCS data into URCS Worktable format
            mStrSQL = "SELECT * FROM " & Get_Table_Name_From_SQL("1", "TRANS") & " " &
                "WHERE year = " & Str(mLooper) & " AND " &
                "sch > 100 AND sch < 148"

            rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)

            rst.MoveFirst()
            Do Until rst.EOF
                mIsDirty = False

                'Compute Total Carloads Originated and Terminated (CLOT)
                mC12 = ((rst.Fields("c1").Value) * 2) + rst.Fields("c3").Value + rst.Fields("c5").Value
                If rst.Fields("C12").Value <> mC12 Then
                    mIsDirty = True
                End If

                'Compute Total Carloads Handled (CLOR)
                mC13 = rst.Fields("c1").Value + rst.Fields("c3").Value + rst.Fields("c5").Value + rst.Fields("c7").Value
                If rst.Fields("c13").Value <> mC13 Then
                    mIsDirty = True
                End If

                'Compute Total Carloads Interchanged (CLRF)
                mC14 = rst.Fields("c3").Value + rst.Fields("c5").Value + (rst.Fields("c7").Value * 2)
                If rst.Fields("c14").Value <> mC14 Then
                    mIsDirty = True
                End If

                'Update the record if any of them are not correct
                If mIsDirty = True Then
                    mStrSQL = Build_Update_Trans_SQL_Record_Statement(mLooper,
                        rst.Fields("rricc").Value,
                        rst.Fields("sch").Value,
                        rst.Fields("line").Value,
                        rst.Fields("c1").Value,
                        rst.Fields("c2").Value,
                        rst.Fields("c3").Value,
                        rst.Fields("c4").Value,
                        rst.Fields("c5").Value,
                        rst.Fields("c6").Value,
                        rst.Fields("c7").Value,
                        rst.Fields("c8").Value,
                        rst.Fields("c9").Value,
                        rst.Fields("c10").Value,
                        rst.Fields("c11").Value,
                        mC12,
                        mC13,
                        mC14).ToString
                    Global_Variables.gbl_ADOConnection.Execute(mStrSQL)
                    TransMod = TransMod + 1
                End If
                rst.MoveNext()
            Loop
            rst.Close()
            rst = Nothing

            'Produce Miles of Road Calculations
            rst = SetRST()
            mStrSQL = "SELECT * FROM " & Get_Table_Name_From_SQL("1", "TRANS") & " " &
                "WHERE year = " & Str(mLooper) & " AND " &
                "sch = 33 AND line = 57"

            rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)

            rst.MoveFirst()
            Do Until rst.EOF

                mIsDirty = False

                mC8 = rst.Fields("c1").Value + rst.Fields("c2").Value + rst.Fields("c3").Value + rst.Fields("c4").Value
                If rst.Fields("c8").Value <> mC8 Then
                    mIsDirty = True
                End If

                mC9 = rst.Fields("c5").Value + rst.Fields("c6").Value
                If rst.Fields("c9").Value <> mC9 Then
                    mIsDirty = True
                End If

                mC10 = rst.Fields("c5").Value + rst.Fields("c8").Value
                If rst.Fields("c10").Value <> mC10 Then
                    mIsDirty = True
                End If

                If mIsDirty = True Then
                    mStrSQL = Build_Update_Trans_SQL_Record_Statement(mLooper,
                        rst.Fields("rricc").Value,
                        rst.Fields("sch").Value,
                        rst.Fields("line").Value,
                        rst.Fields("c1").Value,
                        rst.Fields("c2").Value,
                        rst.Fields("c3").Value,
                        rst.Fields("c4").Value,
                        rst.Fields("c5").Value,
                        rst.Fields("c6").Value,
                        rst.Fields("c7").Value,
                        mC8,
                        mC9,
                        mC10,
                        rst.Fields("c11").Value,
                        rst.Fields("c12").Value,
                        rst.Fields("c13").Value,
                        rst.Fields("c14").Value).ToString
                    Global_Variables.gbl_ADOConnection.Execute(mStrSQL)
                    TransMod = TransMod + 1
                End If
                rst.MoveNext()
            Loop
            rst.Close()
            rst = Nothing

            'Produce Equipment Supporting Schedule
            rst = SetRST()

            mStrSQL = "SELECT * FROM " & Get_Table_Name_From_SQL("1", "TRANS") & " " &
                "WHERE year = " & Str(mLooper) & " AND " &
                "sch = 420"

            rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)

            rst.MoveFirst()
            Do Until rst.EOF

                mIsDirty = False

                mC12 = rst.Fields("c2").Value + rst.Fields("c3").Value
                If Not IsDBNull(rst.Fields("C12").Value) Then
                    If rst.Fields("c12").Value <> mC12 Then
                        mIsDirty = True
                    End If
                End If

                If mIsDirty = True Then
                    mStrSQL = Build_Update_Trans_SQL_Record_Statement(mLooper,
                        rst.Fields("rricc").Value,
                        rst.Fields("sch").Value,
                        rst.Fields("line").Value,
                        rst.Fields("c1").Value,
                        rst.Fields("c2").Value,
                        rst.Fields("c3").Value,
                        rst.Fields("c4").Value,
                        rst.Fields("c5").Value,
                        rst.Fields("c6").Value,
                        rst.Fields("c7").Value,
                        rst.Fields("c8").Value,
                        rst.Fields("c9").Value,
                        rst.Fields("c10").Value,
                        rst.Fields("c11").Value,
                        mC12,
                        rst.Fields("c13").Value,
                        rst.Fields("c14").Value).ToString
                    Global_Variables.gbl_ADOConnection.Execute(mStrSQL)
                    TransMod = TransMod + 1
                End If
                rst.MoveNext()
            Loop

            Me.chk_Trans_Table_Update.Checked = True
            Me.txt_StatusBox.Text = ""
            Me.Refresh()

            rst.Close()
            rst = Nothing

        Next


    End Function


    Private Sub chk_Trans_Table_Update_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Trans_Table_Update.CheckedChanged
        chk_Trans_Table_Update.Checked = True
    End Sub

    Private Sub chk_Created_System_Records_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Created_System_Records.CheckedChanged
        chk_Created_System_Records.Checked = True
    End Sub

    Private Sub chk_Created_Blank_Owner_Records_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Created_Blank_Owner_Records.CheckedChanged
        chk_Created_Blank_Owner_Records.Checked = True
    End Sub
End Class