Imports SpreadsheetGear
Public Class frm_Trans_Comparison

Private Sub btn_Return_To_Menu_Click(sender As System.Object, e As System.EventArgs) Handles btn_Return_To_Menu.Click
    Dim frmNew As New frm_PreProcessing_Checks
    frmNew.Show()
    Me.Close()
End Sub

Private Sub frm_Trans_Comparison_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    Dim rst As ADODB.Recordset
    Dim mStrSQL As String

    'Set the form so it centers on the user's screen
    Me.CenterToScreen()

    ' Load the Year comboboxesz from the SQL database
    OpenADOConnection(Get_Database_Name_From_SQL("1", "TRANS"))

    rst = SetRST()
    mStrSQL = "SELECT DISTINCT year FROM " & Get_Table_Name_From_SQL("1", "TRANS")

    rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)
    rst.MoveFirst()

    ' Load the Year values into the combobox
    Do While Not rst.EOF
            cmb_PreviousYear.Items.Add(rst.Fields("year").Value)
            cmb_CurrentYear.Items.Add(rst.Fields("year").Value)
            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

    End Sub

    Private Sub btn_Output_File_Click(sender As System.Object, e As System.EventArgs) Handles btn_Output_File.Click
        Dim fd As New FolderBrowserDialog

        fd.Description = "Select the location in which you want the output file placed."

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Output_FilePath.Text = fd.SelectedPath.ToString & "\URCS_Trans_Comparison_" & cmb_PreviousYear.Text &
                "_v_" & cmb_CurrentYear.Text & ".xls"
        End If
    End Sub

    Private Sub btn_Execute_Click(sender As System.Object, e As System.EventArgs) Handles btn_Execute.Click

        Dim mTable As New DataTable
        Dim mView As New DataView
        Dim mAdded As Integer = 0
        Dim mUpdated As Integer = 0

        Dim mFields(6) As VariantType
        Dim mValues(6) As VariantType

        Dim rst As ADODB.Recordset
        Dim mWork_rst As New ADODB.Recordset

        Dim mStrSQL As String
        Dim mCellLine As Integer
        Dim mLooper As Integer
        Dim mColumns As Integer
        Dim mRRICC As Integer
        Dim mCellsAdded As Integer
        Dim mCellsInserted As Integer

        ' Variable for SpreadsheetGear
        Dim mWorkbook As IWorkbook
        Dim mExcelSheet As IWorksheet

        ' Check to see that the user has selected all the files
        If IsNothing(Me.cmb_PreviousYear.Text) Then
            MsgBox("You must select a Base Year.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        If IsNothing(Me.cmb_CurrentYear.Text) Then
            MsgBox("You must select a Comparison Year.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        If Me.txt_Output_FilePath.Text = "" Then
            MsgBox("You must select an output file destination.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        ' Make sure that they haven't selected the same file to check against itself
        If Me.cmb_PreviousYear.Text = Me.cmb_CurrentYear.Text Then
            MsgBox("The Base year and the Comparison year cannot be the same.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        ' If the output file exists verify that the user wants to overwrite it.  If not, end this.
        If My.Computer.FileSystem.FileExists(txt_Output_FilePath.Text) Then
            If MsgBox("File already exists!  Overwrite?", vbYesNo, "Warning!") = vbYes Then
                My.Computer.FileSystem.DeleteFile(Me.txt_Output_FilePath.Text)
            Else
                GoTo EndIt
            End If
        End If

        'Open the connection and locate the Trans table
        OpenADOConnection(Get_Database_Name_From_SQL("1", "TRANS"))
        Gbl_Trans_TableName = Get_Table_Name_From_SQL("1", "TRANS")

        mCellsAdded = 0
        mCellsInserted = 0

        ' create the output file
        txt_StatusBox.Text = "Loading Previous Year's SQL Data."
        Refresh()

        'Create the work recordset in memory
        With mWork_rst.Fields
            .Append("RRICC", ADODB.DataTypeEnum.adInteger)
            .Append("Sch", ADODB.DataTypeEnum.adChar, 7)
            .Append("Line", ADODB.DataTypeEnum.adInteger)
            .Append("Col", ADODB.DataTypeEnum.adInteger)
            .Append("Previous_Val", ADODB.DataTypeEnum.adDecimal)
            .Append("Current_Val", ADODB.DataTypeEnum.adDecimal)
        End With
        mWork_rst.Open()

        'get the Road's previous year's information from the Trans table
        rst = New ADODB.Recordset
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Trans_TableName & " WHERE year = " & cmb_PreviousYear.Text &
                " ORDER BY rricc, sch, line"
        rst.Open(mStrSQL, gbl_ADOConnection)

        'load the data for the previous year into the work recordset
        Do While Not rst.EOF
            With mWork_rst
                If rst.Fields("Sch").Value > 100 And rst.Fields("sch").Value < 148 Then
                    mColumns = 14 ' Allows Loss and Damage records to be loaded
                Else
                    mColumns = Get_URCS_Columns(rst.Fields("Sch").Value)
                End If
                For mLooper = 1 To mColumns
                    .AddNew()
                    .Fields("RRICC").Value = rst.Fields("RRICC").Value
                    .Fields("Sch").Value = rst.Fields("Sch").Value
                    .Fields("Line").Value = rst.Fields("Line").Value
                    .Fields("Col").Value = mLooper
                    .Fields("Previous_Val").Value = rst.Fields("C" & mLooper).Value
                    .Fields("Current_Val").Value = 0
                    .Update()
                    mCellsAdded = mCellsAdded + 1
                    If mCellsAdded Mod 1000 = 0 Then
                        txt_StatusBox.Text = "Loading Previous Year's SQL Data. Cells Added: " & mCellsAdded.ToString
                        Text = "Loading " & cmb_PreviousYear.Text & " Data"
                        Refresh()
                    End If
                Next
            End With
            rst.MoveNext()
        Loop

        txt_StatusBox.Text = "Loading Current Year's SQL Data..."
        Refresh()

        rst.Close()
        rst = New ADODB.Recordset
        rst = SetRST()
        mStrSQL = "SELECT * FROM " & Gbl_Trans_TableName & " WHERE year = " & cmb_CurrentYear.Text &
                " ORDER BY rricc, sch, line"
        rst.Open(mStrSQL, gbl_ADOConnection)

        'load the data for the current year
        rst.MoveFirst()
        Do While Not rst.EOF
            With mWork_rst
                If rst.Fields("Sch").Value > 100 And rst.Fields("sch").Value < 148 Then
                    mColumns = 14 ' Allows Loss and Damage records to be loaded
                Else
                    mColumns = Get_URCS_Columns(rst.Fields("Sch").Value)
                End If
                For mLooper = 1 To mColumns
                    .Filter = "RRICC = " & rst.Fields("rricc").Value.ToString &
                        " AND Sch = '" & rst.Fields("sch").Value.ToString &
                        "' AND Line = " & rst.Fields("Line").Value.ToString &
                        " AND Col = " & mLooper.ToString

                    If .RecordCount = 0 Then
                        .AddNew()
                        .Fields("RRICC").Value = rst.Fields("RRICC").Value
                        .Fields("Sch").Value = rst.Fields("Sch").Value
                        .Fields("Line").Value = rst.Fields("Line").Value
                        .Fields("Col").Value = mLooper
                        .Fields("Previous_Val").Value = 0
                    End If
                    .Fields("Current_Val").Value = rst.Fields("C" & mLooper).Value
                    .Update()
                    mCellsInserted = mCellsInserted + 1
                    If mCellsInserted Mod 100 = 0 Then
                        txt_StatusBox.Text = "Loading Current Year's SQL Data. Cells Inserted: " &
                            mCellsInserted.ToString
                        Text = "Loading " & cmb_CurrentYear.Text & " Data"
                        Refresh()
                    End If
                    .Filter = ""
                    .MoveFirst()
                Next
            End With
            rst.MoveNext()
        Loop

        mRRICC = 0

        ' Open the Excel file
        txt_StatusBox.Text = "Creating & Loading Excel Output File..."
        Refresh()

        mWorkbook = Factory.GetWorkbook(Globalization.CultureInfo.CurrentCulture)
        mExcelSheet = mWorkbook.ActiveSheet
        mWorkbook.ActiveSheet.Name = "Summary"

        With mExcelSheet
            ' Set the data to the Summary page and format it
            .Cells("A1").Value = "URCS Trans/R-1 Comparison Summary - " & cmb_PreviousYear.Text.ToString & " v. " &
                cmb_CurrentYear.Text
            .Cells("A1").HorizontalAlignment = HAlign.Left
            .Cells("A1").Font.Bold = True
            .Cells("A1").Font.Size = 18
            .Cells("A1").ColumnWidth = 15
            .Cells("A2").Value = "Previous Year:"
            .Cells("B2").Value = cmb_PreviousYear.Text
            .Cells("A3").Value = "Current Year:"
            .Cells("B3").Value = cmb_CurrentYear.Text
            .Cells("B2:B3").NumberFormat = "@"
            .Cells("A4").Value = "Date/Time:"
            .Cells("B4").Value = FormatDateTime(Now, DateFormat.ShortDate)
            .Cells("B4").NumberFormat = "MM/DD/YYYY"
            .Cells("A5").Value = "Railroads in Previous Year:"
            .Cells("A6").Value = "Railroads in Current Year:"

            rst = SetRST()
            mStrSQL = "SELECT DISTINCT rricc FROM " & Gbl_Trans_TableName & " WHERE year = " & cmb_PreviousYear.Text &
                " AND rricc < 900000"
            rst.Open(mStrSQL, gbl_ADOConnection)
            .Cells("A5").Value = .Cells("A5").Value & " " & rst.RecordCount.ToString

            rst = SetRST()
            mStrSQL = "SELECT DISTINCT rricc FROM " & Gbl_Trans_TableName & " WHERE year = " & cmb_CurrentYear.Text &
                " AND rricc < 900000"
            rst.Open(mStrSQL, gbl_ADOConnection)
            .Cells("A6").Value = .Cells("A6").Value & " " & rst.RecordCount.ToString

            .Cells("A2:B6").Columns.AutoFit()
            .Cells("A2:B6").HorizontalAlignment = HAlign.Left
        End With

        With mWork_rst
            .Sort = "RRICC, Sch, Line, Col"
            .MoveFirst()
            Do While Not .EOF
                If mRRICC = 0 Or mRRICC <> .Fields("rricc").Value Then
                    mRRICC = .Fields("rricc").Value
                    ' Add a new sheet and populate headers
                    mWorkbook.Worksheets.Add()
                    mWorkbook.ActiveSheet.Name = Get_Short_Name_By_RRICC(mRRICC)
                    mExcelSheet = mWorkbook.ActiveSheet

                    txt_StatusBox.Text = "Creating Excel File - Creating Sheet for " & mExcelSheet.Name
                    Text = "Creating " & mExcelSheet.Name & " Sheet..."
                    Refresh()
                    With mExcelSheet
                        .Cells("A1").Value = "URCS Trans/R-1 Comparison - " & cmb_PreviousYear.Text.ToString & " v. " &
                    cmb_CurrentYear.Text & " - " & .Name
                        .Cells("A1").HorizontalAlignment = HAlign.Left
                        .Cells("A1").Font.Bold = True
                        .Cells("A1").Font.Size = 18
                        .Cells("A1").ColumnWidth = 15
                        .Cells("A2").Value = "Previous Year:"
                        .Cells("B2").Value = cmb_PreviousYear.Text
                        .Cells("A3").Value = "Current Year:"
                        .Cells("B3").Value = cmb_CurrentYear.Text
                        .Cells("B2:B3").HorizontalAlignment = HAlign.Left
                        .Cells("B2:B3").NumberFormat = "@"
                        .Cells("A4").Value = "Date/Time:"
                        .Cells("B4").Value = Now.ToString("MM/dd/yyyy")
                        .Cells("B4").NumberFormat = "MM/dd/yyyy"
                        .Cells("A2:B4").Columns.AutoFit()

                        .Cells("A6").Value = "Sch"
                        .Cells("A6").HorizontalAlignment = HAlign.Right
                        .Cells("B6").Value = "Line"
                        .Cells("C6").Value = "Col"
                        .Cells("D6").Value = cmb_PreviousYear.Text & vbCrLf & " Value"
                        .Cells("E6").Value = cmb_CurrentYear.Text & vbCrLf & " Value"
                        .Cells("F6").Value = "Variance"

                        .Cells("A6:F6").Style.HorizontalAlignment = HAlign.Center
                        .Cells("A6:F6").Font.Bold = True
                        .Cells("A6:F6").Font.Underline = True
                        mCellLine = 6
                    End With
                End If
                mExcelSheet.Cells("D:F").Columns.AutoFit()
                mExcelSheet.Cells("D:F").HorizontalAlignment = HAlign.Right
                With mExcelSheet
                    If mWork_rst.Fields("sch").Value > 100 And mWork_rst.Fields("sch").Value < 148 Then
                        ' handle the Loss and Damage records
                        .Cells(mCellLine, 0).Value = Trim(mWork_rst.Fields("Sch").Value) & "-L&D"
                    Else
                        .Cells(mCellLine, 0).Value = Get_URCS_Schedule(mWork_rst.Fields("Sch").Value)
                    End If
                    .Cells(mCellLine, 0).HorizontalAlignment = HAlign.Right
                    .Cells(mCellLine, 1).Value = mWork_rst.Fields("Line").Value
                    .Cells(mCellLine, 2).Value = mWork_rst.Fields("Col").Value
                    .Cells(mCellLine, 0, mCellLine, 2).NumberFormat = "0"
                    .Cells(mCellLine, 3).Value = mWork_rst.Fields("Previous_Val").Value
                    .Cells(mCellLine, 4).Value = mWork_rst.Fields("Current_Val").Value
                    .Cells(mCellLine, 5).Value = Calculate_Variance(.Cells(mCellLine, 3).Value, .Cells(mCellLine, 4).Value)
                    .Cells(mCellLine, 3, mCellLine, 4).NumberFormat = "#,##0"
                    .Cells(mCellLine, 5).Style.NumberFormat = "#0.00"

                    mCellLine = mCellLine + 1
                End With
                .MoveNext()
            Loop
        End With

        'Save the file
        mWorkbook.Worksheets(0).Select()
        mWorkbook.SaveAs(txt_Output_FilePath.Text, FileFormat.Excel8)

        ' Housekeeping
        mWork_rst = Nothing
        rst.Close()
        rst = Nothing

        Me.txt_StatusBox.Text = "Done!"
        Text = "Done!"

EndIt:


    End Sub
End Class