Imports System.Xml
Public Class Convert_To_XML_Legacy

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Converts this object to an XML legacy load. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub Convert_To_XML_Legacy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        mDataTable = Get_URCS_Years_Table()

        For mLooper = 1 To mDataTable.Rows.Count
            cmb_URCS_Year.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
        Next

        mDataTable = Nothing

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button return to data prep menu click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        Dim frmNew As New frm_MainMenu
        frmNew.Show()
        Me.Close()
    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button output file entry click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Output_File_Entry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New FolderBrowserDialog

        If Me.cmb_URCS_Year.Text = "" Then
            MsgBox("Error - You must select a year to process!", vbOKOnly, "Error!")
        Else
            fd.Description = "Select the folder where the UMF File is located."

            If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.txt_Output_FilePath.Text = fd.SelectedPath.ToString
            End If
        End If

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Button execute click. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="sender">   Source of the event. </param>
    ''' <param name="e">        Event information. </param>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub btn_Execute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Execute.Click
        Dim fs, outfile
        Dim XML_Doc As XmlDocument

        'These are historical URCS control codes
        Dim RRurcsid() As Integer = {8, 10, 20, 22, 30, 35, 37, 44, 47, 49}
        Dim RRurcsName() As String = {"CN", "NS", "CSXT", "BN", "KCS", "CP", "UP", "EAST", "WEST", "NATIONAL"}
        Dim RRegionName() As String = {"REG4", "REG4", "REG4", "REG7", "REG7", "REG7", "REG7", "REG4", "REG7", " "}

        XML_Doc = New XmlDocument
        Try
            XML_Doc.Load(txt_Output_FilePath.ToString & "\XML Files\All Railroads " & cmb_URCS_Year.Text & ".XML")
        Catch ex As System.Exception
            MsgBox("Error! Cannot create XML file.", vbOKOnly, "Critical Error!")
            GoTo Endit
        End Try

        'Convert the UMF File

        'Convert the Index File

        'Convert the Dictionary

        'open the printfile - overwrite it if it exists
        fs = CreateObject("Scripting.FileSystemObject")
        outfile = fs.createtextfile(Me.txt_Output_FilePath.Text, ForWriting)

        With outfile

        End With
Endit:

    End Sub

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>   Convert urcs input data. </summary>
    '''
    ''' <remarks>   Michael Sanders, 10/27/2020. </remarks>
    '''
    ''' <param name="XML_Document"> The XML document. </param>
    '''
    ''' <returns>   True if it succeeds, false if it fails. </returns>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////

    Function ConvertURCSInputData(ByVal XML_Document As XmlDocument) As Boolean
        Dim XML_Comment As XmlComment
        Dim XML_Declaration As XmlDeclaration
        Dim ProcessYear As XmlElement
        'Dim Railroad As XmlElement
        'Dim RailroadAttribute As XmlAttribute
        Dim Root As XmlElement
        'Dim WorktableAddress As XmlElement
        'Dim WorktableAttribute As XmlAttribute
        Dim YearAttribute As XmlAttribute
        'Dim RRidx As Integer
        Dim oldRRIdx As Integer, RRVal As Integer
        'Dim RRstrName As String
        'Dim RRStrTitle As String
        'Dim RRStrRegion As String
        Dim I As Integer, Year As Integer, RRCnt As Integer
        'Dim datyear As Double, datyear1 As Double, datyear2 As Double, datyear3 As Double, datyear4 As Double
        Dim SR1 As StreamReader
        Dim txtRecord As String, ProcessDate As String
        Dim RRtitle() As String

        'Declare the XML document
        XML_Declaration = XML_Document.CreateXmlDeclaration("1.0", "UTF-16", "yes")
        XML_Document.PrependChild(XML_Declaration)
        'Create the root node
        Root = XML_Document.CreateElement("URCSInputData")
        XML_Document.AppendChild(Root)

        'Open ASCII text input file
        SR1 = New StreamReader(txt_Output_FilePath.ToString & "\URCUMF." & cmb_URCS_Year.Text)
        'Get file control record (1st record in file)
        txtRecord = SR1.ReadLine
        Year = CInt(txtRecord.Substring(0, 4))
        RRCnt = CInt(txtRecord.Substring(39, 2))
        ProcessDate = txtRecord.Substring(txtRecord.LastIndexOf(" ") + 1, txtRecord.Length - 1)

        If RRCnt < 2 Then
            MsgBox("Error! Cannot determine number of railroads in file/data error.", vbOKOnly, "Critical Error!")
            Return False
        End If
        RRCnt = RRCnt - 1 'removes National from processing
        ReDim RRtitle(RRCnt)

        ' Write processing comment
        XML_Comment = XML_Document.CreateComment("This file Created: " & Today.Date & " From URCS File Created on " & _
                                                 ProcessDate)
        Root.AppendChild(XML_Comment)

        ' Read the railroad names from the input file
        For I = 0 To RRCnt
            txtRecord = SR1.ReadLine
            RRtitle(I) = txtRecord.Substring(3, txtRecord.Length - 3)
        Next

        'read past all system records
        For I = 0 To 5256
            SR1.ReadLine()
        Next

        ' Set up processing year documentation
        ProcessYear = XML_Document.CreateElement("DataYears")
        Root.AppendChild(ProcessYear)
        YearAttribute = XML_Document.CreateAttribute("Year")
        YearAttribute.Value = Year
        ProcessYear.Attributes.Append(YearAttribute)
        YearAttribute = XML_Document.CreateAttribute("Year1")
        YearAttribute.Value = Year - 1
        ProcessYear.Attributes.Append(YearAttribute)
        YearAttribute = XML_Document.CreateAttribute("Year2")
        YearAttribute.Value = Year - 2
        ProcessYear.Attributes.Append(YearAttribute)
        YearAttribute = XML_Document.CreateAttribute("Year3")
        YearAttribute.Value = Year - 3
        ProcessYear.Attributes.Append(YearAttribute)
        YearAttribute = XML_Document.CreateAttribute("Year4")
        YearAttribute.Value = Year - 4
        ProcessYear.Attributes.Append(YearAttribute)
        oldRRIdx = 0

        'read actual railroad data
        While SR1.Peek <> -1
            txtRecord = SR1.ReadLine
            RRVal = CInt(txtRecord.Substring(0, 2))
            'If RRVal <> oldRRIdx Then
            '    If oldRRIdx > 0 the
            ' RRidx = MatchUCRSIndex(RRVal)

            'End If
        End While

        Return True

    End Function
End Class