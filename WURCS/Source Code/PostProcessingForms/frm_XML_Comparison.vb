Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Reflection
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Xml

Public Class frm_XML_Comparison

    Private Sub btn_Test_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Test_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "XML Files|*.xml|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Test_FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub btn_Base_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Base_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "XML Files|*.xml|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Base_FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub btn_Output_File_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Output_File.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = False
        fd.Filter = "Excel Files|*.xlsx|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Output_FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub btn_Return_To_Menu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_Menu.Click
        Dim frmNew As New frm_Post_Processing_Menu
        frmNew.Show()
        Me.Close()
    End Sub

    Private Sub frm_XML_Comparison_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click
        Dim mOverwrite As Boolean
        Dim BaseFileStream As StreamReader, TestFileStream As StreamReader
        Dim BaseCodes(0, 973) As String, BaseVals(0, 973) As Double
        Dim TestCodes(0, 973) As String, TestVals(0, 973) As Double
        Dim OutArray(973, 5) As String
        Dim eCodes(973) As String, Descs(973) As String
        Dim BaseRoads() As String, TestRoads() As String
        Dim mBaseString As String, mTestString As String
        Dim mBaseRoadsFound As Integer, mTestRoadsFound As Integer
        Dim mCharPos As Integer, mCellPos As Integer
        Dim mXML_RRName As String
        Dim mNumValue As String
        Dim mLineValue As String
        Dim mCellNum As String
        Dim mCellValue As String
        Dim m_ECode As String
        Dim mLooper As Integer, mArrayPos As Integer, mRoads As Integer
        Dim oXL As Excel.Application = Nothing
        Dim oWBs As Excel.Workbooks = Nothing
        Dim oWB As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim oCells As Excel.Range = Nothing
        Dim rst As New ADODB.Recordset
        Dim mSQLStr As String

        mOverwrite = False
        mArrayPos = 1
        mBaseRoadsFound = 0
        mTestRoadsFound = 0
        mXML_RRName = ""
        ReDim BaseRoads(0)
        ReDim TestRoads(0)

        ' Check to see that the user has selected all the files
        If IsNothing(Me.txt_Base_FilePath.Text) Then
            MsgBox("You must select a Base text file.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        If IsNothing(Me.txt_Test_FilePath.Text) Then
            MsgBox("You must select a Test text file", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        If IsNothing(Me.txt_Output_FilePath.Text) Then
            MsgBox("You must select an output file destination.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        ' Make sure that they haven't selected the same file to check against itself
        If Me.txt_Base_FilePath.Text = Me.txt_Test_FilePath.Text Then
            MsgBox("The Test source file and the Base file cannot be the same file.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        ' If the output file exists verify that the user wants to overwrite it.  If not, end this.
        If My.Computer.FileSystem.FileExists(Me.txt_Output_FilePath.Text) Then
            If MsgBox("File already exists!  Overwrite?", vbYesNo, "Warning!") = vbYes Then
                My.Computer.FileSystem.DeleteFile(Me.txt_Output_FilePath.Text)
            Else
                GoTo EndIt
            End If
        End If

        OpenADOConnection(Get_Database_Name_From_SQL("1", "ECode"))

        'Initialize the return recordset and run the query to load it
        rst = New ADODB.Recordset

        mSQLStr = "SELECT ecode, etitle_and_desc FROM " & Get_Table_Name_From_SQL("1", "ECode")
        rst.Open(mSQLStr, Global_Variables.gbl_ADOConnection)

        'Load the eCodes and Descs to the arrays
        rst.MoveFirst()
        mLooper = 1
        Do While Not rst.EOF
            eCodes(mLooper) = rst.Fields("ecode").Value
            Descs(mLooper) = rst.Fields("etitle_and_desc").Value
            rst.MoveNext()
            mLooper = mLooper + 1
        Loop
        rst.Close()
        rst = Nothing
        
        ' OK, now that that is out of the way, open the base file
        BaseFileStream = New StreamReader(Me.txt_Base_FilePath.Text)

        ' Load the Railroad name values to the Baseroads array
        mBaseString = BaseFileStream.ReadLine()
        Do While Not BaseFileStream.EndOfStream
            mBaseString = mBaseString.Trim
            Select Case mBaseString.Substring(0, 2)
                Case "<R" ' Railroad Name line - we need this
                    ' Have to get the railroad name value
                    ' Start by removing the '"Railroad Name="' string
                    mBaseString = mBaseString.Substring(16)
                    ' Then locate the trailing quote mark and get the character up to that point - 1
                    mCharPos = InStr(mBaseString.ToString, Chr(34)) - 1
                    ' save the railroad name
                    mXML_RRName = mBaseString.Substring(0, mCharPos)
                    ReDim Preserve BaseRoads(BaseRoads.Length)
                    BaseRoads(BaseRoads.Length - 1) = mXML_RRName
                    mBaseRoadsFound = mBaseRoadsFound + 1
            End Select
            mBaseString = BaseFileStream.ReadLine()
        Loop

        ' close the base file and open the test file
        BaseFileStream.Close()
        TestFileStream = New StreamReader(Me.txt_Test_FilePath.Text, True)

        ' read the railroad records and see if each road/region can be found in the array
        mTestString = TestFileStream.ReadLine()
        Do While Not TestFileStream.EndOfStream
            mTestString = mTestString.Trim
                Select mTestString.Substring(0, 2)
                Case "<R" ' Railroad Name line - we need this
                    ' Have to get the railroad name value
                    ' Start by removing the '"Railroad Name="' string
                    mTestString = mTestString.Substring(16)
                    ' Then locate the trailing quote mark and get the character up to that point - 1
                    mCharPos = InStr(mTestString.ToString, Chr(34)) - 1
                    ' save the railroad name
                    mXML_RRName = mTestString.Substring(0, mCharPos)
                    ReDim Preserve TestRoads(TestRoads.Length)
                    TestRoads(TestRoads.Length - 1) = mXML_RRName
                    mTestRoadsFound = mTestRoadsFound + 1
            End Select
            mTestString = TestFileStream.ReadLine()
        Loop


        ' Close the test file
        TestFileStream.Close()

        'Resize the arrays
        ReDim BaseCodes(mBaseRoadsFound, 973)
        ReDim BaseVals(mBaseRoadsFound, 973)
        ReDim TestCodes(mTestRoadsFound, 973)
        ReDim TestVals(mTestRoadsFound, 973)

        ' Load the BaseVals array using the position value in Baseroads
        BaseFileStream = New StreamReader(Me.txt_Base_FilePath.Text)

        mBaseString = BaseFileStream.ReadLine()
        Do While Not BaseFileStream.EndOfStream
            mBaseString = mBaseString.Trim
            Select Case mBaseString.Substring(0, 2)
                Case "<R" ' Railroad Name line - we need this
                    ' Have to get the railroad name value
                    ' Start by removing the '"Railroad Name="' string
                    mBaseString = mBaseString.Substring(16)
                    ' Then locate the trailing quote mark and get the character up to that point - 1
                    mCharPos = InStr(mBaseString.ToString, Chr(34)) - 1
                    ' Find the RR in the 0 position of the BaseCodes array
                    mXML_RRName = mBaseString.Substring(0, mCharPos)
                    ' Find the road in the array.  Default to 0 if not found
                    mArrayPos = Array1DFindFirst(BaseRoads, mXML_RRName, 0)
                    If mArrayPos > 0 Then
                        Me.txt_StatusBox.Text = "Loading Base Data for " & mXML_RRName
                        Me.Refresh()
                        ' let's make sure that the environment is reset
                        For mLooper = 1 To 973
                            BaseVals(mArrayPos, mLooper) = 0
                        Next
                    End If
                Case "<E" ' Data line - only process if the road was found in the array
                    If mArrayPos > 0 Then
                        ' remove the first character
                        mBaseString = mBaseString.Substring(1)
                        ' Parse the Ex values and remove it from the string
                        mNumValue = mBaseString.Substring(0, 2)
                        mBaseString = mBaseString.Substring(2)
                        ' Drop the "Py" values from the string
                        mBaseString = mBaseString.Substring(2)
                        ' get the line number info and drop them
                        mLineValue = mBaseString.Substring(0, 4)
                        mBaseString = mBaseString.Substring(4)
                        ' Now we loop through the remainder of the line, parsing out the cells
                        Do While InStr(mBaseString, "C") > 0
                            ' drop any leading space
                            mBaseString = mBaseString.Trim
                            ' if the next character is "C", we're good
                            If mBaseString.Substring(0, 1) = "C" Then
                                ' find the "=" sign
                                mCharPos = InStr(mBaseString.ToString, "=") - 1
                                ' save the cell address number - C1, C2, etc
                                mCellNum = mBaseString.Substring(0, mCharPos)
                                ' remove the characters prior to the value
                                mCharPos = InStr(mBaseString.ToString, Chr(34))
                                mBaseString = mBaseString.Substring(mCharPos)
                                ' save the value and then remove it along with any trailing space
                                mCharPos = InStr(mBaseString.ToString, Chr(34)) - 1
                                mCellValue = mBaseString.Substring(0, mCharPos)
                                mBaseString = mBaseString.Substring(mCharPos + 1)
                                ' Put the Ecode together for this cell
                                m_ECode = mNumValue & mLineValue & mCellNum
                                ' Find the ecode in the array so we know where to put the data
                                mCellPos = Array1DFindFirst(eCodes, m_ECode, 0)
                                ' Save the data to the arrays
                                BaseVals(mArrayPos, mCellPos) = CDbl(mCellValue)
                            End If
                        Loop
                    End If
            End Select
            mBaseString = BaseFileStream.ReadLine()
        Loop
        BaseFileStream.Close()

        ' Load the TestVals array using the position value in Baseroads
        ' This will make each road to be in the same position in each array - 1 to 1 match
        TestFileStream = New StreamReader(Me.txt_Test_FilePath.Text)

        mTestString = TestFileStream.ReadLine()
        Do While Not TestFileStream.EndOfStream
            mTestString = mTestString.Trim
            Select Case mTestString.Substring(0, 2)
                Case "<R" ' Railroad Name line - we need this
                    ' Have to get the railroad name value
                    ' Start by removing the '"Railroad Name="' string
                    mTestString = mTestString.Substring(16)
                    ' Then locate the trailing quote mark and get the character up to that point - 1
                    mCharPos = InStr(mTestString.ToString, Chr(34)) - 1
                    ' Find the RR in the 0 position of the BaseCodes array
                    mXML_RRName = mTestString.Substring(0, mCharPos)
                    ' Find the road in the array.  Default to 0 if not found
                    mArrayPos = Array1DFindFirst(BaseRoads, mXML_RRName, 0)
                    If mArrayPos > 0 Then
                        Me.txt_StatusBox.Text = "Loading Test Data for " & mXML_RRName
                        Me.Refresh()
                        ' let's make sure that the environment is reset
                        For mLooper = 1 To 973
                            TestVals(mArrayPos, mLooper) = 0
                        Next
                    End If
                Case "<E" ' Data line - only process if the road was found in the array
                    If mArrayPos > 0 Then
                        ' remove the first character
                        mTestString = mTestString.Substring(1)
                        ' Parse the Ex values and remove it from the string
                        mNumValue = mTestString.Substring(0, 2)
                        mTestString = mTestString.Substring(2)
                        ' Drop the "Py" values from the string
                        mTestString = mTestString.Substring(2)
                        ' get the line number info and drop them
                        mLineValue = mTestString.Substring(0, 4)
                        mTestString = mTestString.Substring(4)
                        ' Now we loop through the remainder of the line, parsing out the cells
                        Do While InStr(mTestString, "C") > 0
                            ' drop any leading space
                            mTestString = mTestString.Trim
                            ' if the next character is "C", we're good
                            If mTestString.Substring(0, 1) = "C" Then
                                ' find the "=" sign
                                mCharPos = InStr(mTestString.ToString, "=") - 1
                                ' save the cell address number - C1, C2, etc
                                mCellNum = mTestString.Substring(0, mCharPos)
                                ' remove the characters prior to the value
                                mCharPos = InStr(mTestString.ToString, Chr(34))
                                mTestString = mTestString.Substring(mCharPos)
                                ' save the value and then remove it along with any trailing space
                                mCharPos = InStr(mTestString.ToString, Chr(34)) - 1
                                mCellValue = mTestString.Substring(0, mCharPos)
                                mTestString = mTestString.Substring(mCharPos + 1)
                                ' Put the Ecode together for this cell
                                m_ECode = mNumValue & mLineValue & mCellNum
                                ' Find the ecode in the array so we know where to put the data
                                mCellPos = Array1DFindFirst(eCodes, m_ECode, 0)
                                ' Save the data to the arrays
                                TestVals(mArrayPos, mCellPos) = CDbl(mCellValue)
                            End If
                        Loop
                    End If
            End Select
            mTestString = TestFileStream.ReadLine()
        Loop
        TestFileStream.Close()

        ' At this point we have all we need in memory

        Try
            Me.txt_StatusBox.Text = "Creating & Loading Excel File..."
            Me.Refresh()
            ' Open the Excel App
            oXL = New Excel.Application
            oXL.Visible = False
            ' Create the new Workbook
            oWBs = oXL.Workbooks
            oWB = oWBs.Add()

            ' Set the active sheet's name
            oSheet = oWB.ActiveSheet
            oSheet.Name = "Summary"
            ' Set the data to the Summary page
            oCells = oSheet.Cells

            oCells(1, 1) = "URCS XML Comparison Summary"
            oCells(3, 1) = "Base File:"
            oCells(4, 1) = "Test File:"
            oCells(6, 1) = "Date/Time:"
            oCells(8, 1) = "RRs/Regs in Base File:"
            oCells(9, 1) = "RRs/Regs in Test File:"

            oCells(3, 2) = Me.txt_Base_FilePath.Text
            oCells(4, 2) = Me.txt_Test_FilePath.Text
            oCells(6, 2) = CStr(Now)
            oCells(8, 2) = mBaseRoadsFound.ToString()
            oCells(9, 2) = mTestRoadsFound.ToString()

            oCells = oSheet.Range("A1")
            oCells.Font.Bold = True
            oCells.Font.Size = 14

            oSheet.Columns("A:B").ColumnWidth = 20


            ' Determine if there is a constraint on the loop as one file might be smaller than the other
            If mBaseRoadsFound >= mTestRoadsFound Then
                mRoads = mBaseRoadsFound
            Else
                mRoads = mTestRoadsFound
            End If

            ' Now we create/write the sheet for each road/region
            For mLooper = 1 To mRoads
                oSheet = oWB.Worksheets.Add(After:=oWB.Worksheets(oWB.Worksheets.Count))
                oSheet = oWB.ActiveSheet
                oSheet.Name = BaseRoads(mLooper)
                'load the data to the OutArray to dump into Excel
                For mCellPos = 1 To 973
                    OutArray(mCellPos, 0) = eCodes(mCellPos)
                    OutArray(mCellPos, 1) = Descs(mCellPos)
                    OutArray(mCellPos, 2) = BaseVals(mLooper, mCellPos)
                    OutArray(mCellPos, 3) = TestVals(mLooper, mCellPos)
                    If TestVals(mLooper, mCellPos) = 0 And BaseVals(mLooper, mCellPos) = 0 Then
                        OutArray(mCellPos, 4) = 0
                    ElseIf (BaseVals(mLooper, mCellPos) = 0 And TestVals(mLooper, mCellPos) <> 0) Then
                        OutArray(mCellPos, 4) = 1
                    ElseIf (BaseVals(mLooper, mCellPos) <> 0 And TestVals(mLooper, mCellPos) = 0) Then
                        OutArray(mCellPos, 4) = -1
                    Else
                        OutArray(mCellPos, 4) = _
                            CDbl(TestVals(mLooper, mCellPos)) / CDbl(BaseVals(mLooper, mCellPos)) - 1
                    End If
                Next
                ' Dump it into Excel
                oSheet.Range("A3").Resize(974, 5).Value = OutArray
                ' Set up and Bold the headers
                oCells = oSheet.Cells
                oCells(1, 1) = "Variances for " & BaseRoads(mLooper)
                oCells(3, 1) = "ECode"
                oCells(3, 2) = "eTitle_and_Desc"
                oCells(3, 3) = "Base Value"
                oCells(3, 4) = "Test Value"
                oCells(3, 5) = "Variance Percentage"
                oCells = oSheet.Range("A1", "E3")
                oCells.Font.Bold = True
                oCells = oSheet.Range("A3", "E3")
                oCells.Font.Underline = True
                oCells = oSheet.Range("E4", "E977")
                oCells.NumberFormat = "0.00%"

                ' And resize all of the columns to autofit
                oCells = oSheet.Range("A4", "E977")
                ' And change the format from text to numbers
                oCells.Value = oCells.Value
                ' And resize all of the columns to autofit
                oCells.EntireColumn.AutoFit()

            Next

            ' Set the summary sheet as default when the file opens in Excel for the first time
            CType(oXL.ActiveWorkbook.Sheets(1), Excel.Worksheet).Activate()

            ' We're done!  Save it and close excel
            oWB.SaveAs(Me.txt_Output_FilePath.Text)
            oSheet = Nothing
            oWB = Nothing
            oXL.Quit()
            oXL = Nothing

            Me.txt_StatusBox.Text = "Done!"

        Catch ex As System.Exception
            Console.WriteLine("Solution1.AutomateExcel throws the error: {0}", _
                              ex.Message)
        Finally

            ' Clean up the unmanaged Excel COM resources by explicitly call  
            ' Marshal.FinalReleaseComObject on all accessor objects.  
            ' See http://support.microsoft.com/kb/317109. 

            If Not oCells Is Nothing Then
                Marshal.FinalReleaseComObject(oCells)
                oCells = Nothing
            End If
            If Not oSheet Is Nothing Then
                Marshal.FinalReleaseComObject(oSheet)
                oSheet = Nothing
            End If
            If Not oWB Is Nothing Then
                Marshal.FinalReleaseComObject(oWB)
                oWB = Nothing
            End If
            If Not oWBs Is Nothing Then
                Marshal.FinalReleaseComObject(oWBs)
                oWBs = Nothing
            End If
            If Not oXL Is Nothing Then
                Marshal.FinalReleaseComObject(oXL)
                oXL = Nothing
            End If

        End Try

EndIt:

    End Sub
End Class