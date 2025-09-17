Imports System.Data.SqlClient
Public Class frm_XML_Data_Loader

    Private Sub frm_XML_Data_Loader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rst As ADODB.Recordset
        Dim mStrsql As String

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        'Get the years table location
        Gbl_URCS_Years_TableName = Get_Table_Name_From_SQL("1", "URCS_Years")

        ' Load the Year combobox from the SQL database
        ' Open the SQL connection using the global variable holding the connection string
        OpenADOConnection(Gbl_Controls_Database_Name)

        ' Execute the query
        rst = SetRST()
        mStrsql = "SELECT urcs_year FROM " & Gbl_URCS_Years_TableName

        rst.Open(mStrsql, gbl_ADOConnection)
        rst.MoveFirst()

        Do While Not rst.EOF
            cmb_URCSYear.Items.Add(rst.Fields("urcs_year").Value)
            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

    End Sub

    Private Sub btn_Return_To_Menu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_Menu.Click
        ' Open the Post Processing Menu
        Dim frmNew As New frm_Post_Processing_Menu
        frmNew.Show()
        ' Close this form
        Me.Close()
    End Sub

    Private Sub btn_Input_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "XML Files|*.xml|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click

        Dim mInString As String
        Dim sr As StreamReader
        Dim mRailroad As String
        Dim mXML_Name As String
        Dim mXML_Title As String
        Dim mRRICC As String
        Dim mNumValue As String
        Dim mLineValue As String
        Dim mCellNum As String
        Dim mCellValue As String
        Dim m_ECode As String
        Dim m_eCode_Id As String
        Dim mCharPos As Integer
        Dim mRecs As Integer

        mRailroad = ""
        mRRICC = ""
        mXML_Name = ""
        mXML_Title = ""
        mRecs = 0

        ' Check to make sure that the user has selected a year and an input file
        If IsNothing(Me.cmb_URCSYear.Text) Then
            MsgBox("You must select a year value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        If txt_Input_FilePath.TextLength = 0 Then
            MsgBox("You must select an Input File value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        ' Open the filestream for reading the input file
        sr = New StreamReader(txt_Input_FilePath.Text)

        ' Time to parse Washburn's crappy XML file

        ' Check to see if we're at EOF
        Do While Not sr.EndOfStream
            mInString = sr.ReadLine()
            mInString = mInString.Trim
            Select Case mInString.Substring(0, 2)
                Case "<?" ' xml formatting line data - nothing to do here
                Case "<!" ' comment line - again - nothing to do here.
                Case "<U" ' unitcost header - ignore this as well
                Case "</" ' End of the railroad or the unitcost - ignore it
                Case "<R" ' Railroad Name line - we need this
                    ' Have to get the railroad name value
                    ' Start by removing the '"Railroad Name="' string
                    mInString = mInString.Substring(16)
                    ' Then locate the trailing quote mark and get the character up to that point - 1
                    mCharPos = InStr(mInString.ToString, Chr(34)) - 1
                    mXML_Name = mInString.Substring(0, mCharPos)
                    ' Now that we have the road name, get the RRICC from the lookup table
                    mRRICC = Get_RRICC_By_Short_Name(mXML_Name)
                    ' And also get the Railroad Title from the file
                    ' Locate the next quote mark
                    mCharPos = InStr(mInString.ToString, "Title=") + 6
                    ' drop the data up to that point
                    mInString = mInString.Substring(mCharPos)
                    ' locate the trailing quote mark
                    mCharPos = InStr(mInString.ToString, Chr(34)) - 1
                    ' Grab the title from the string
                    mXML_Title = mInString.Substring(0, mCharPos)
                Case "<E" ' Data line - we definately need to process this
                    ' remove the first character
                    mInString = mInString.Substring(1)
                    ' Parse the Ex values and remove it from the string
                    mNumValue = mInString.Substring(0, 2)
                    mInString = mInString.Substring(2)
                    ' Drop the "Py" values from the string
                    mInString = mInString.Substring(2)
                    ' get the line number info and drop them
                    mLineValue = mInString.Substring(0, 4)
                    mInString = mInString.Substring(4)
                    ' Now we loop through the remainder of the line, parsing out the cells
                    Do While InStr(mInString, "C") > 0
                        ' drop any leading space
                        mInString = mInString.Trim
                        ' if the next character is "C", we're good
                        If mInString.Substring(0, 1) = "C" Then
                            ' find the "=" sign
                            mCharPos = InStr(mInString.ToString, "=") - 1
                            ' save the cell address number - C1, C2, etc
                            mCellNum = mInString.Substring(0, mCharPos)
                            ' remove the characters prior to the value
                            mCharPos = InStr(mInString.ToString, Chr(34))
                            mInString = mInString.Substring(mCharPos)
                            ' save the value and then remove it along with any trailing space
                            mCharPos = InStr(mInString.ToString, Chr(34)) - 1
                            mCellValue = mInString.Substring(0, mCharPos)
                            mInString = mInString.Substring(mCharPos + 1)
                            ' Put the Ecode together for this cell
                            m_ECode = mNumValue & mLineValue & mCellNum
                            ' Now that we have all of the data for this cell,
                            ' get the ecode info from the UR_ECode table
                            m_eCode_Id = Get_ECode_Id_By_ECode(m_ECode)
                            ' Now we have all the data for the SQL table
                            ' Perform the write/update
                            Write_XML_Data_SQL( _
                                Me.cmb_URCSYear.Text, _
                                mRRICC, _
                                mXML_Name, _
                                mXML_Title, _
                                m_eCode_Id, _
                                m_ECode, _
                                mCellValue)
                            mRecs = mRecs + 1
                            If mRecs Mod 100 = 0 Then
                                Me.txt_StatusBox.Text = "Working - Loaded " & CStr(mRecs) & " records..."
                                Me.Refresh()
                                System.Windows.Forms.Application.DoEvents()
                            End If
                        End If
                    Loop

            End Select
        Loop
        Me.txt_StatusBox.Text = "Done!"
        sr.Close()
EndIt:

    End Sub
End Class