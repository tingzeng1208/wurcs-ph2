'**********************************************************************
' Title:        R-1 Balance Report Module
' Author:       Michael Sanders
' Purpose:      Produces the Balance Report to make data integrity checking easier for URCS.
' Revisions:    Initial Creation - 3 August 2015
'               Commented out those Schedules that are not used by URCS - 8 April 2016
' 
' This program is US Government Property - For Official Use Only
'**********************************************************************

Imports System.Data.SqlClient
Public Class frm_R_1_Balance

    Private Sub btn_Return_To_DataPrepMenu_Click(sender As System.Object, e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click

        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_PreProcessing_Checks
        frmNew.Show()
        ' Close this Form
        Me.Close()

    End Sub

    Private Sub frm_R_1_Balance_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

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

    Private Sub btn_Output_File_Entry_Click(sender As System.Object, e As System.EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New FolderBrowserDialog
        Dim Format_Selected As Boolean, Year_Selected As Boolean

        Year_Selected = False
        Format_Selected = False

        If cmb_URCS_Year.Text <> "" Then
            Year_Selected = True
        End If

        If Year_Selected = False Then
            MsgBox("You must set/select the Waybill Year before selecting an output file.", vbOKOnly, "Error!")
            GoTo SkipIt
        End If

        fd.Description = "Select the location in which you want the output report placed."

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Output_FilePath.Text = fd.SelectedPath.ToString & "\" & Me.cmb_URCS_Year.Text
            ' Add the description of the file to the filename
            txt_Output_FilePath.Text = txt_Output_FilePath.Text & "_R1_Balance_Report"
            ' add the extension
            txt_Output_FilePath.Text = txt_Output_FilePath.Text & ".rtf"
        End If

SkipIt:
    End Sub

    Private Sub btn_Execute_Click(sender As System.Object, e As System.EventArgs) Handles btn_Execute.Click
        
        Dim mStrSQL As String

        Dim rst As ADODB.Recordset
        Dim FileWriter As StreamWriter
        Dim bolWrite As Boolean, mFirst As Boolean

        mFirst = True

        'Perform error checking for the form controls
        If Me.cmb_URCS_Year.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo EndIt
        End If

        If Me.txt_Output_FilePath.Text = "" Then
            MsgBox("You must select an output directory to use.", vbOKOnly)
            GoTo EndIt
        End If

        'Now we're ready to go...

        bolWrite = MsgBox("Ready to run report?", vbYesNo)

        If bolWrite = True Then

            'Delete the output file if it exists
            If My.Computer.FileSystem.FileExists(Me.txt_Output_FilePath.Text) Then
                My.Computer.FileSystem.DeleteFile(Me.txt_Output_FilePath.Text)
            End If

            ' Open the RTF output file and write the header
            If (File.Exists(Me.txt_Output_FilePath.Text)) Then
                System.IO.File.Delete(Me.txt_Output_FilePath.Text)
            End If
            FileWriter = File.CreateText(Me.txt_Output_FilePath.Text)

            ' put the header info to the file
            FileWriter.WriteLine("{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}")
            FileWriter.WriteLine("{\header\pard\qr\f0 " & Date.Today.ToString("MM/dd/yyyy") & "\par}")
            FileWriter.Close()

            'Get the list of Railroads
            Global_Variables.Gbl_Class1RailList_TableName = Get_Table_Name_From_SQL("1", "CLASS1RAILLIST")
            rst = SetRST()

            OpenADOConnection(Global_Variables.Gbl_Controls_Database_Name)

            mStrSQL = "Select * from " & Global_Variables.Gbl_Class1RailList_TableName & " WHERE (SHORT_NAME <> 'EAST' AND " & _
                "SHORT_NAME <> 'WEST' AND SHORT_NAME <> 'ALL')"

            rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)
            rst.ActiveConnection = Nothing

            rst.MoveFirst()
            Do While rst.EOF <> True

                ' Step through each of the R1 Schedules
                ' Each subroutine will append its RTF lines to the output file

                ' Write the page header for each road
                FileWriter = File.AppendText(txt_Output_FilePath.Text)
                If mFirst = False Then
                    FileWriter.WriteLine("\page")
                Else
                    mFirst = False
                End If
                FileWriter.WriteLine("{\pard \qc \b UNIFORM RAIL COSTING SYSTEM (URCS) - " & _
                Me.cmb_URCS_Year.Text & "\line")
                FileWriter.WriteLine("R-1 Balance Report for " & rst.Fields("rricc").Value & " - " & rst.Fields("Short_Name").Value & "\line \line")
                FileWriter.WriteLine("Unbalanced lines are in bold with the differing totals in \line parentheses following each source location.\line \line\par}")
                FileWriter.Close()

                txt_StatusBox.Text = "Please Wait - Processing Schedule 200 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule200( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 210 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule210( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                'txt_StatusBox.Text = "Please Wait - Processing Schedule 220 for " & rst.Fields("Short_Name").Value
                'Refresh()

                'R1_Schedule220( _
                '    Me.cmb_URCS_Year.Text, _
                '    rst.Fields("rricc").Value, _
                '    Me.txt_Output_FilePath.Text)

                'txt_StatusBox.Text = "Please Wait - Processing Schedule 240 for " & rst.Fields("Short_Name").Value
                'Refresh()

                'R1_Schedule240( _
                '    Me.cmb_URCS_Year.Text, _
                '    rst.Fields("rricc").Value, _
                '    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 245 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule245( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                'txt_StatusBox.Text = "Please Wait - Processing Schedule 332 for " & rst.Fields("Short_Name").Value
                'Refresh()

                'R1_Schedule332( _
                '    Me.cmb_URCS_Year.Text, _
                '    rst.Fields("rricc").Value, _
                '    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 335 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule335( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                'txt_StatusBox.Text = "Please Wait - Processing Schedule 339 for " & rst.Fields("Short_Name").Value
                'Refresh()

                'R1_Schedule339( _
                '    Me.cmb_URCS_Year.Text, _
                '    rst.Fields("rricc").Value, _
                '    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 340 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule340( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 342 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule342( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                'txt_StatusBox.Text = "Please Wait - Processing Schedule 350 for " & rst.Fields("Short_Name").Value
                'Refresh()

                'R1_Schedule350( _
                '    Me.cmb_URCS_Year.Text, _
                '    rst.Fields("rricc").Value, _
                '    Me.txt_Output_FilePath.Text)

                'txt_StatusBox.Text = "Please Wait - Processing Schedule 351 for " & rst.Fields("Short_Name").Value
                'Refresh()

                'R1_Schedule351( _
                '    Me.cmb_URCS_Year.Text, _
                '    rst.Fields("rricc").Value, _
                '    Me.txt_Output_FilePath.Text)

                'txt_StatusBox.Text = "Please Wait - Processing Schedule 352A for " & rst.Fields("Short_Name").Value
                'Refresh()

                'R1_Schedule352A( _
                '    Me.cmb_URCS_Year.Text, _
                '    rst.Fields("rricc").Value, _
                '    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 352B for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule352B( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 410 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule410( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 412 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule412( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 414 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule414( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 415 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule415( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                'txt_StatusBox.Text = "Please Wait - Processing Schedule 416 for " & rst.Fields("Short_Name").Value
                'Refresh()

                'R1_Schedule416( _
                '    Me.cmb_URCS_Year.Text, _
                '    rst.Fields("rricc").Value, _
                '    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 417 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule417( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                'txt_StatusBox.Text = "Please Wait - Processing Schedule 450 for " & rst.Fields("Short_Name").Value
                'Refresh()

                'R1_Schedule450( _
                '    Me.cmb_URCS_Year.Text, _
                '    rst.Fields("rricc").Value, _
                '    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 510 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule510( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 700 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule700( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                'txt_StatusBox.Text = "Please Wait - Processing Schedule 702 for " & rst.Fields("Short_Name").Value
                'Refresh()

                'R1_Schedule702( _
                '    Me.cmb_URCS_Year.Text, _
                '    rst.Fields("rricc").Value, _
                '    Me.txt_Output_FilePath.Text)

                'txt_StatusBox.Text = "Please Wait - Processing Schedule 710 for " & rst.Fields("Short_Name").Value
                'Refresh()

                'R1_Schedule710( _
                '    Me.cmb_URCS_Year.Text, _
                '    rst.Fields("rricc").Value, _
                '    Me.txt_Output_FilePath.Text)

                txt_StatusBox.Text = "Please Wait - Processing Schedule 755 for " & rst.Fields("Short_Name").Value
                Refresh()

                R1_Schedule755( _
                    Me.cmb_URCS_Year.Text, _
                    rst.Fields("rricc").Value, _
                    Me.txt_Output_FilePath.Text)

                rst.MoveNext()
            Loop

            ' Lastly, open the output file and close out the RTF code
            FileWriter = File.AppendText(Me.txt_Output_FilePath.Text)
            FileWriter.WriteLine("}}")
            FileWriter.Close()

            txt_StatusBox.Text = "Done!"
            Refresh()

            ' Close the connection and clean up
            rst.Close()
            rst = Nothing
            
        End If

EndIt:

    End Sub

    Sub R1_Schedule200( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        ' At design time this schedule was identified as SCH 2 in the TRANS database. but we'll
        ' verify it before proceeding

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mSumLine As String

        ' Find out what URCS Code is used for Sch 200
        mURCSCode = Get_URCS_Code("200")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 200\par}")
            ' Now that we have the RRICC, the schedule and the year, we can begin checking the values of the R1 data
            ' Get lines 1 thru 13's sum from SQL

            ' Set the database connection to the URCSyyyy database
            Global_Variables.Gbl_URCSProdYear_Database_Name = "URCS" & cmb_URCS_Year.Text
            OpenADOConnection(Global_Variables.Gbl_URCSProdYear_Database_Name)

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 1, 13)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 14)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L1 Col B thru L13 Col B (" & mSum1.ToString & ") <> L14 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L1 Col B thru L13 Col B = L14 Col B\par}")
            End If

            ' Get lines 15 thru 22's sum from SQL and compare it to line 23
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 15, 22, 23))

            ' Get lines 24 thru 27's sum from SQL and compare it to line 28
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 24, 27, 28))

            'Check the Totals to see that they match line 29
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 14)
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 23)
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 28)

            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 29)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L14 Col B + L23 Col B + L28 Col B (" & mSum1.ToString & ") <> L29 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L14 Col B + L23 Col B + L28 Col B = L29 Col B\par}")
            End If

            ' Get lines 30 thru 39's sum from SQL and compare it to line 40
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 30, 39, 40))

            ' Get lines 41 thru 50's sum from SQL and compare it to line 51
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 41, 50, 51))

            ' Get lines 53 thru 54's sum from SQL and compare it to line 52
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 53, 54)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 52)

            If mSum1 <> mSum2 Then
                .WriteLine(SumsDontMatch("B", 53, 54, mSum1.ToString, 52, mSum2.ToString))
            Else
                .WriteLine(SumsMatch("B", 53, 54, 52))
            End If

            ' We need to save the value from line 52
            mSum1 = mSum2

            ' Get lines 55 thru 59's sum from SQL and add line 52's value
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 55, 59)

            ' If line 60 > 0 subtract it
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 60)

            If mSum2 > 0 Then
                mSum1 = mSum1 - mSum2
            End If

            mSumLine = "61"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 61)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b (L52 Col B + (L55 Col B thru L59 Col B) - L60 Col B) + L61 Col B (" & mSum1.ToString & ") <> L62 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab (L52 Col B + (L55 Col B thru L59 Col B) - L60 Col B) + L61 Col B = L62 Col B\par}")
            End If

            ' Finally, we have to see if the subtotals match the grand total
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 40)
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 51)
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 61)

            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 62)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L40 Col B + L51 Col B + L61 Col B (" & mSum1.ToString & ") <> L62 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L40 Col B + L51 Col B + L61 Col B = L62 Col B\par}")
            End If

            ' Crosscheck to Sch 330 for lines 24 and 25
            mURCSCode = Get_URCS_Code("200")
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 24)
            mURCSCode = Get_URCS_Code("330")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", 30)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b Crosscheck Error - L24 Col B (" & mSum1.ToString & ") <> Sch 330 L30 Col H (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L24 Col B = Sch 330 L30 Col H\par}")
            End If

            mURCSCode = Get_URCS_Code("200")
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 25)
            mURCSCode = Get_URCS_Code("330")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", 39)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b Crosscheck Error - L25 Col B (" & mSum1.ToString & ") <> Sch 330 L39 Col H (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L25 Col B = Sch 330 L39 Col H\par}")
            End If

            .WriteLine("{\pard\Line\Line\par}")
        End With

        Filewriter.Close()


    End Sub

    Sub R1_Schedule210( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String, mURCSCode2 As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mSumLine As String
        Dim mcolNum As Integer, mlineNum As Integer

        ' Find out what URCS Code is used for Sch 200
        mURCSCode = Get_URCS_Code("210")
        ' Find out what URCS Code is used for Sch 410
        mURCSCode2 = Get_URCS_Code("410")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 210\par}")
            ' Now that we have the RRICC, the schedule and the year, we can begin checking the values of the R1 data
            ' Get lines 1 thru 9's sum from SQL and compare them to line 10
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 1, 9, 10))

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 10, 12, 13))

            ' Subtract L14 from L13
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 13)
            mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 14)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 15)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L13 Col B - L14 Col B (" & mSum1.ToString & ") <> L15 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L13 Col B - L14 Col B = L15 Col B \par}")
            End If

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 16, 26, 27))

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 15)
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 27)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 28)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L15 Col B + L27 Col B (" & mSum1.ToString & ") <> L28 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L15 Col B + L27 Col B = L28 Col B\par}")
            End If

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 29, 35, 36))

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 28)
            mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 36)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 37)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L28 Col B - L36 Col B (" & mSum1.ToString & ") <> L37 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L28 Col B - L36 Col B = L37 Col B\par}")
            End If

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 38, 41, 42))

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 37)
            mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 42)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 43)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L37 Col B - L42 Col B (" & mSum1.ToString & ") <> L43 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L37 Col B - L42 Col B = L43 Col B\par}")
            End If

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 43, 45, 46))

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 47, 50, 51))

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 46)
            mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 51)

            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 52)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L46 Col B - L51 Col B (" & mSum1.ToString & ") <> L52 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L46 Col B - L51 Col B = L52 Col B\par}")
            End If

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 52, 54, 55))

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 56, 58, 59))

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 55)
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 59)
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 60)

            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 61)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L55 Col B + L59 Col B + L60 Col B (" & mSum1.ToString & ") <> L61 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L55 Col B + L59 Col B + L60 Col B = L61 Col B \par}")
            End If

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 62)
            mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 63)
            mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 64)
            mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 65)
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 66)

            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 67)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L62 Col B - L63 Col B - L64 Col B - L65 Col B + L66 Col B (" & mSum1.ToString & ") <> L67 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L62 Col B - L63 Col B - L64 Col B - L65 Col B + L66 Col B = L67 Col B \par}")
            End If

            ' Perform Crosschecks

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 15)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 62)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L15 Col B (" & mSum1.ToString & ") <> L62 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L15 Col B = L62 Col B \par}")
            End If

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 47, 49, 63))

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 50)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 64)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L50 Col B (" & mSum1.ToString & ") <> L64 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L50 Col B = L64 Col B \par}")
            End If

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 14)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode2, "H", 620)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L14 Col B (" & mSum1.ToString & ") <> Sch 410 L20 Col H (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L14 Col B = Sch 410 L620 Col H\par}")
            End If

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "D", 14)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode2, "F", 620)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L14 Col D (" & mSum1.ToString & ") <> Sch 410 L620 Col F (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L14 Col D = Sch 410 L620 Col F\par}")
            End If

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "E", 14)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode2, "G", 620)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L14 Col E (" & mSum1.ToString & ") <> Sch 410 L20C6 (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L14 Col E = Sch 410 L620 Col G\par}")
            End If

            ' Check for invalid data
            For mlineNum = 16 To 37
                For mcolNum = 3 To 4
                    mSumLine = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C" & CStr(mcolNum), mlineNum)
                    If mSumLine <> 0 Then
                        .WriteLine("{\pard\tab \b L" & CStr(mlineNum) & "C" & CStr(mcolNum) & " <> 0 - Unexpected/Invalid Value\par}")
                    End If
                Next
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule220( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String, mStartline As String, mEndLine As String, mSumLine As String

        ' Find out what URCS Code is used for Sch 220
        mURCSCode = Get_URCS_Code("220")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 220\par}")

            mCol = "C1"
            mStartline = "3"
            mEndLine = "5"
            mSumLine = "6"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mStartline = "7"
            mEndLine = "12"
            mSumLine = "13"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mSumLine = "6"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "13"
            mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "14"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L6 Col B - L13 Col B (" & mSum1.ToString & ") <> L14 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L6 Col B - L13 Col B = L14 Col B \par}")
            End If

            mSumLine = "1"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "2"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "14"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "15"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L1 Col B + L2 Col B + L14 Col B (" & mSum1.ToString & ") <> L15 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L1 Col B + L2 Col B + L14 Col B = L15 Col B \par}")
            End If

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule240( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String, mStartline As String, mEndLine As String, mSumLine As String

        ' Find out what URCS Code is used for Sch 220
        mURCSCode = Get_URCS_Code("240")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 240\par}")

            mCol = "C1"
            mStartline = "1"
            mEndLine = "8"
            mSumLine = "9"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mStartline = "10"
            mEndLine = "18"
            mSumLine = "19"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mStartline = "19"
            mEndLine = "20"
            mSumLine = "21"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mStartline = "22"
            mEndLine = "28"
            mSumLine = "29"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mStartline = "30"
            mEndLine = "35"
            mSumLine = "36"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mSumLine = "21"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "29"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "36"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mSumLine = "37"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L21 Col B + L29 Col B + L36 Col B (" & mSum1.ToString & ") <> L37 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L21 Col B + L29 Col B + L36 Col B = L37 Col B \par}")
            End If

            mStartline = "37"
            mEndLine = "38"
            mSumLine = "39"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule245( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String, mURCSCode1 As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mStartline As String, mEndLine As String, mSumLine As String

        ' Find out what URCS Code is used for Sch 245
        mURCSCode = Get_URCS_Code("245")
        ' And for Sch 200
        mURCSCode1 = Get_URCS_Code("200")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 245\par}")

            ' Get Sch 245, L1C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 1)

            'Get Sch 200, L5C1
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "B", 5)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L1 Col B (" & mSum1.ToString & ") <> Sch 200 L5 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L1 Col B = Sch 200 L5 Col B \par}")
            End If

            ' Get Sch 245, L2C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 2)

            'Get Sch 200, L6C1
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "B", 6)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L2 Col B (" & mSum1.ToString & ") <> Sch 200 L6 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L2 Col B = Sch 200 L6 Col B \par}")
            End If

            mStartline = "1"
            mEndLine = "3"
            mSumLine = "4"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 1, 3, 4))

            mURCSCode1 = Get_URCS_Code("210")

            ' Get Sch 245, L5C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 5)

            'Get Sch 210, L13C1
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "B", 13)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L5 Col B (" & mSum1.ToString & ") <> Sch 210 L13 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L5 Col B = Sch 210 L13 Col B \par}")
            End If

            ' Get Sch 245, L6C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 6)

            'Get Sch 410, L121 Col H + L122 Col H + L123 Col H + L127 Col H + L128 Col H + L129 Col H + L133 Col H + L134 Col H + L135 Col H + L208 Col H + L210 Col H + L212 Col H + L227 Col H + L229 Col H + L231 Col H + L312 Col H + L314 Col H + L316C1
            mURCSCode1 = Get_URCS_Code("410")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 121, 123)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 127, 129)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 133, 135)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 208)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 210)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 212)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 227)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 229)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 231)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 312)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 314)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 316)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L6 Col B (" & mSum1.ToString & ") <> " & _
                           "Sch 410 L121 Col H + L122 Col H + L123 Col H + L127 Col H + L128 Col H + L129 Col H + L133 Col H + L134 Col H + L135 Col H + L208 Col H + L210 Col H + L212 Col H + L227 Col H + L229 Col H + L231 Col H + L312 Col H + L314 Col H + L316 Col H (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L6 Col H = Sch 410 L121 Col H + L122 Col H + L123 Col H + L127 Col H + L128 Col H + L129 Col H + L133 Col H + L134 Col H + L135 Col H + L208 Col H + L210 Col H + L212 Col H + L227 Col H + L229 Col H + L231 Col H + L312 Col H + L314 Col H + L316 Col H \par}")
            End If

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 5, 6, 7))

            ' Get Sch 245, L11C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 11)

            'Get Sch 200, L31C1
            mURCSCode1 = Get_URCS_Code("200")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "B", 31)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L11 Col B (" & mSum1.ToString & ") <> Sch 200 L31 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L11 Col B = Sch 200 L31 Col B \par}")
            End If

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "B", 11, 14, 15))

            ' Get Sch 245, L16C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 16)

            'Get Sch 210, L14C1
            mURCSCode1 = Get_URCS_Code("210")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "B", 14)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L16 Col B (" & mSum1.ToString & ") <> Sch 210 L14 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L16 Col B = Sch 210 L14 Col B \par}")
            End If

            ' Get Sch 245, L17C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", "17")

            'Get Sch 410, L136 Col H + L137 Col H + L138 Col H + L213 Col H + L232 Col H + L317 Col H
            mURCSCode1 = Get_URCS_Code("410")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 136, 138)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 213)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 232)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "H", 317)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L17 Col B (" & mSum1.ToString & ") <> Sch 410 L136 Col H + L137 Col H + L138 Col H + L213 Col H + L232 Col H + L317 Col H (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L17 Col B = Sch 410 L136 Col H + L137 Col H + L138 Col H + L213 Col H + L232 Col H + L317 Col H\par}")
            End If

            mSumLine = "18"
            ' Get Sch 245, L18C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 18)

            ' Get Sch 245, L16 Col B + L6 Col B - L17C1
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 16)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 6)
            mSum2 = mSum2 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 17)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L18 Col B (" & mSum1.ToString & ") <> L16 Col B + L6 Col B - L17 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L18 Col B = L16 Col B + L6 Col B - L17 Col B \par}")
            End If

            ' Get Sch 245, L20C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 20)

            ' Get Sch 245, L15 Col B + L19C1
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 15)
            mSum2 = mSum2 / Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 19)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L20 Col B (" & mSum1.ToString & ") <> L15 Col B / L19 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L20 Col B = L15 Col B + L19 Col B \par}")
            End If

            ' Get Sch 245, L21C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 21)

            ' Get Sch 245, L10 Col B - L20C1
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 10)
            mSum2 = mSum2 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 20)

            If mSum2 < 0 Then
                mSum2 = 0
            End If

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L21 Col B (" & mSum1.ToString & ") <> L10 Col B - L20 Col B (" & mSum2.ToString & ")\par}")
            Else
                If mSum2 = 0 Then
                    .WriteLine("{\pard\tab L21 Col B = L10 Col B - L20 Col B (Result is negative or 0)\par}")
                Else
                    .WriteLine("{\pard\tab L21 Col B = L10 Col B - L20 Col B \par}")
                End If
            End If

            ' Get Sch 245, L22C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 22)

            ' Get Sch 245, L21 Col B x L19C1
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 21)
            mSum2 = mSum2 * (Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 19))

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L22 Col B (" & mSum1.ToString & ") <> L21 Col B x L19 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L22 Col B = L21 Col B x L19 Col B \par}")
            End If

            mURCSCode1 = Get_URCS_Code("200")

            mSumLine = "23"
            ' Get Sch 245, L23C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 23)

            ' Get Sch 200, L1 Col B + L2C1
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode1, "B", 1, 2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L23 Col B (" & mSum1.ToString & ") <> Sch 200 L1 Col B + L2 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L23 Col B = Sch 200 L1 Col B + L2 Col B \par}")
            End If

            ' Get Sch 245, L27C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 27)

            ' Get Sch 245, L25 Col B - L26C1
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 25)
            mSum2 = mSum2 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 26)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L27 Col B (" & mSum1.ToString & ") <> L25 Col B - L26 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L27 Col B = L25 Col B - L26 Col B \par}")
            End If

            mSumLine = "28"
            ' Get Sch 245, L28C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 28)

            ' Get Sch 245, L24 Col B + L27C1
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 24)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", 27)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L28 Col B (" & mSum1.ToString & ") <> L24 Col B + L27 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L28 Col B = L24 Col B + L27 Col B \par}")
            End If

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule332( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String
        Dim mStartline As String, mEndLine As String, mSumLine As String

        ' Find out what URCS Code is used for Sch 332
        mURCSCode = Get_URCS_Code("332")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 332\par}")

            mCol = "C1"

            mStartline = "1"
            mEndLine = "29"
            mSumLine = "30"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mCol = "C2"

            mStartline = "1"
            mEndLine = "29"
            mSumLine = "30"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mCol = "C4"

            mStartline = "1"
            mEndLine = "29"
            mSumLine = "30"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mCol = "C5"

            mStartline = "1"
            mEndLine = "29"
            mSumLine = "30"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mCol = "C1"

            mStartline = "31"
            mEndLine = "59"
            mSumLine = "60"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mCol = "C2"

            mStartline = "31"
            mEndLine = "59"
            mSumLine = "60"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mCol = "C4"

            mStartline = "31"
            mEndLine = "59"
            mSumLine = "60"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mCol = "C5"

            mStartline = "31"
            mEndLine = "38"
            mSumLine = "39"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mSumLine = "40"
            ' Get Sch 332, L40C1
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            ' Get Sch 332, L30 Col B + L39C1
            mSumLine = "30"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "39"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L40 Col B (" & mSum1.ToString & ") <> L30 Col B + L39 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L40 Col B = L30 Col B + L39 Col B \par}")
            End If

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule335( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String, mThisCol As String
        Dim mStartline As String, mEndLine As String, mSumLine As String
        Dim mLooper As Integer

        mThisCol = ""

        ' Find out what URCS Code is used for Sch 335
        mURCSCode = Get_URCS_Code("335")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 335\par}")

            For mLooper = 1 To 6
                mCol = "C" & CStr(mLooper)

                mStartline = "1"
                mEndLine = "29"
                mSumLine = "30"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 6
                mCol = "C" & CStr(mLooper)

                mStartline = "31"
                mEndLine = "39"
                mSumLine = "40"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 6
                ' Get Sch 335, L30Cx+L39Cx
                mCol = "C" & CStr(mLooper)

                mThisCol = ColText(mCol)

                mSumLine = "41"
                ' Get Sch 335, L41Cx
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                mSumLine = "30"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "40"
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L41" & mThisCol & "(" & mSum1.ToString & ") <> L30" & mThisCol & "+ L40" & mThisCol & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L41" & mThisCol & "= L30" & mThisCol & "+ L40" & mThisCol & "\par}")
                End If
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule339( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String, mThisCol As String
        Dim mStartline As String, mEndLine As String, mSumLine As String
        Dim mLooper As Integer

        mThisCol = ""

        ' Find out what URCS Code is used for Sch 335
        mURCSCode = Get_URCS_Code("339")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 339\par}")

            For mLooper = 1 To 6
                mCol = "C" & CStr(mLooper)

                mStartline = "1"
                mEndLine = "29"
                mSumLine = "30"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 6
                mCol = "C" & CStr(mLooper)

                mStartline = "31"
                mEndLine = "39"
                mSumLine = "40"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 6
                ' Get Sch 339, L30Cx+L40Cx
                mCol = "C" & CStr(mLooper)

                mThisCol = ColText(mCol)

                mSumLine = "41"
                ' Get Sch 332, L40C1
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                mSumLine = "30"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "40"
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L41" & mThisCol & "(" & mSum1.ToString & ") <> L30" & mThisCol & "+ L40" & mThisCol & " (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L41" & mThisCol & "= L30" & mThisCol & "+ L40" & mThisCol & "\par}")
                End If
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule340( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String, mThisCol As String
        Dim mLooper As Integer

        mThisCol = ""

        ' Find out what URCS Code is used for Sch 335
        mURCSCode = Get_URCS_Code("340")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 340\par}")

            For mLooper = 1 To 2
                mCol = "C" & CStr(mLooper)

                mThisCol = ColText(mCol)

                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, 1, 29, 30))
            Next

            For mLooper = 1 To 2
                mCol = "C" & CStr(mLooper)

                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, 31, 39, 40))
            Next

            For mLooper = 1 To 2
                ' Get Sch 340, L30Cx+L40Cx
                mCol = "C" & CStr(mLooper)

                ' Get Sch 340, L40Cx
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, 41)

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, 30)
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, 40)

                Select Case mLooper
                    Case 1
                        mThisCol = "Col B"
                    Case Else
                        mThisCol = "Col C"
                End Select

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L41 " & mThisCol & " (" & mSum1.ToString & ") <> L30 " & mThisCol & " + L40 " & mThisCol & " (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L41 " & mThisCol & " = L30 " & mThisCol & " + L40 " & mThisCol & "\par}")
                End If
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule342( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String, mThisCol As String
        Dim mStartline As String, mEndLine As String, mSumLine As String
        Dim mLooper As Integer

        mThisCol = ""

        ' Find out what URCS Code is used for Sch 335
        mURCSCode = Get_URCS_Code("342")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 342\par}")

            For mLooper = 1 To 6
                mCol = "C" & CStr(mLooper)

                mThisCol = ColText(mCol)

                mStartline = "1"
                mEndLine = "28"
                mSumLine = "29"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 6
                mCol = "C" & CStr(mLooper)

                mStartline = "30"
                mEndLine = "37"
                mSumLine = "38"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 6
                ' Get Sch 342, L29Cx+L38Cx
                mCol = "C" & CStr(mLooper)

                mSumLine = "39"
                ' Get Sch 340, L39Cx
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                mSumLine = "29"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "38"
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L39" & ColText(mCol).ToString & " (" & mSum1.ToString & ") <> L29" & ColText(mCol).ToString & " + L38" & ColText(mCol).ToString & " (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L39" & ColText(mCol).ToString & " = L29" & ColText(mCol).ToString & " + L38" & ColText(mCol).ToString & "\par}")
                End If
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule350( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String, mThisCol As String
        Dim mStartline As String, mEndLine As String, mSumLine As String
        Dim mLooper As Integer

        mThisCol = ""

        ' Find out what URCS Code is used for Sch 335
        mURCSCode = Get_URCS_Code("350")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 350\par}")

            For mLooper = 1 To 2
                mCol = "C" & CStr(mLooper)

                mThisCol = ColText(mCol)

                mStartline = "1"
                mEndLine = "28"
                mSumLine = "29"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 2
                mCol = "C" & CStr(mLooper)

                mStartline = "30"
                mEndLine = "37"
                mSumLine = "38"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 2
                ' Get Sch 350, L29Cx+L38Cx
                mCol = "C" & CStr(mLooper)

                mSumLine = "39"
                ' Get Sch 350, L39Cx
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                mSumLine = "29"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "38"
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L39" & mThisCol & "(" & mSum1.ToString & ") <> L29" & mThisCol & "+ L38" & mThisCol & " (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L39" & mThisCol & "= L29" & mThisCol & "+ L 38" & mThisCol & "\par}")
                End If
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule351( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String, mThisCol As String
        Dim mStartline As String, mEndLine As String, mSumLine As String
        Dim mLooper As Integer

        mThisCol = ""

        ' Find out what URCS Code is used for Sch 335
        mURCSCode = Get_URCS_Code("351")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 351\par}")

            For mLooper = 1 To 6
                mCol = "C" & CStr(mLooper)

                mThisCol = ColText(mCol)

                mStartline = "1"
                mEndLine = "28"
                mSumLine = "29"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 6
                mCol = "C" & CStr(mLooper)

                mStartline = "30"
                mEndLine = "37"
                mSumLine = "38"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 6
                ' Get Sch 351, L29Cx+L38Cx
                mCol = "C" & CStr(mLooper)

                mSumLine = "39"
                ' Get Sch 351, L39Cx
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                mSumLine = "29"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "38"
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L39" & mThisCol & "(" & mSum1.ToString & ") <> L29" & mThisCol & "+ L38" & mThisCol & " (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L39" & mThisCol & "= L29" & mThisCol & "+ L38" & mThisCol & "\par}")
                End If
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule352A( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mCol As String
        Dim mStartline As String, mEndLine As String, mSumLine As String
        Dim mLooper As Integer


        ' Find out what URCS Code is used for Sch 335
        mURCSCode = Get_URCS_Code("352A")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 352A\par}")

            For mLooper = 1 To 3
                mCol = "C" & CStr(mLooper)

                mStartline = "1"
                mEndLine = "30"
                mSumLine = "31"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule352B( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String, mThisCol As String
        Dim mStartline As String, mEndLine As String, mSumLine As String
        Dim mLooper As Integer

        mThisCol = ""

        ' Find out what URCS Code is used for Sch 352B
        mURCSCode = Get_URCS_Code("352B")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 352B\par}")

            For mLooper = 1 To 4
                mCol = "C" & CStr(mLooper)

                mThisCol = ColText(mCol)

                mStartline = "1"
                mEndLine = "30"
                mSumLine = "31"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 4
                mCol = "C" & CStr(mLooper)

                mStartline = "32"
                mEndLine = "39"
                mSumLine = "40"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mLooper = 1 To 4
                ' Get Sch 352B, L31Cx+L40Cx+L41Cx+L42Cx+L43Cx
                mCol = "C" & CStr(mLooper)

                mThisCol = ColText(mCol)

                mSumLine = "44"
                ' Get Sch 352B, L44Cx
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                mSumLine = "31"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "40"
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "41"
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "42"
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "43"
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L44" & mThisCol & "(" & mSum1.ToString & ") <> L30" & mThisCol & _
                               "+ L40" & mThisCol & "+ L41" & mThisCol & "+ L42" & mThisCol & "+ L43" & mThisCol & " (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L39" & mThisCol & "= L30" & mThisCol & "+ L40" & mThisCol & _
                               "+ L41" & mThisCol & "+ L42" & mThisCol & "+ L43" & mThisCol & "\par}")
                End If
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule410( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String, mCol2 As String, mThisCol As String
        Dim mStartline As String, mEndLine As String, mSumLine As String, mSumLine2 As String
        Dim mColLooper As Integer, mLineLooper As Integer

        mThisCol = ""

        ' Find out what URCS Code is used for Sch 410
        mURCSCode = Get_URCS_Code("410")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 410\par}")

            ' Calculate/compare the values horizontally
            For mLineLooper = 1 To 151
                mSumLine = mLineLooper
                mSum1 = 0

                For mColLooper = 1 To 4
                    mCol = "C" & CStr(mColLooper)
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                Next

                mCol = "C5"
                ' Get Sch 410, LxC5
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col F (" & mSum2.ToString & ") <> L" & mLineLooper & " Col B" & _
                               "+ L" & mLineLooper & " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col F = L" & mLineLooper & " Col B + L" & mLineLooper & _
                               " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E\par}")
                End If

                mCol = "C6"
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                mCol = "C7"
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col H (" & mSum1.ToString & ") <> L" & mLineLooper & " Col F" & _
                               " + L" & mLineLooper & " Col G (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col H = L" & mLineLooper & " Col F + L" & mLineLooper & " Col G\par}")
                End If

            Next

            'Calculate/compare the values vertically
            For mColLooper = 1 To 7

                mCol = "C" & CStr(mColLooper)

                mStartline = "1"
                mEndLine = "120"
                mSum1 = 0
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "121"
                mEndLine = "123"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "124"
                mEndLine = "126"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "127"
                mEndLine = "129"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "130"
                mEndLine = "132"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "133"
                mEndLine = "135"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "136"
                mEndLine = "141"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "142"
                mEndLine = "144"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "145"
                mEndLine = "150"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mSumLine = "151"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L1 thru L150" & ColText(mCol).ToString & " (" & mSum1.ToString & ") <> L151" & _
                        ColText(mCol).ToString & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L1 thru L150" & ColText(mCol).ToString & " = L151" & ColText(mCol).ToString & "\par}")
                End If

            Next

            ' Calculate/compare the values horizontally
            For mLineLooper = 201 To 219
                mSum1 = 0
                mSumLine = mLineLooper
                For mColLooper = 1 To 4
                    mCol = "C" & CStr(mColLooper)
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                Next

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col F (" & mSum1.ToString & ") <> Col B" & _
                               "+ Col C + Col D + Col E (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col F = Col B + Col C + Col D + Col E\par}")
                End If

                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", mSumLine)

                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & "Col H (" & mSum1.ToString & ") <> L" & mLineLooper & " Col F" & _
                               "+ L" & mLineLooper & " Col G (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col H = L" & mLineLooper & " Col F + L" & mLineLooper & " Col G\par}")
                End If

            Next

            'Calculate/compare the values vertically
            For mColLooper = 1 To 7

                mCol = "C" & CStr(mColLooper)

                mStartline = "201"
                mEndLine = "207"
                mSum1 = 0
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "208"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "209"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "210"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "211"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "212"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "213"
                mEndLine = "214"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "215"
                mEndLine = "216"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "217"
                mEndLine = "218"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mSumLine = "219"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L201 thru L218" & ColText(mCol).ToString & " (" & mSum1.ToString & ") <> L219" & _
                        ColText(mCol).ToString & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L201 thru L218" & ColText(mCol).ToString & " = L219" & ColText(mCol).ToString & "\par}")
                End If

            Next

            ' Calculate/compare the values horizontally
            For mLineLooper = 220 To 238
                mSum1 = 0
                mSumLine = mLineLooper
                For mColLooper = 1 To 4
                    mCol = "C" & CStr(mColLooper)
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                Next

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col F (" & mSum1.ToString & ") <> L" & mLineLooper & " Col B" & _
                            " + L" & mLineLooper & " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col F = L" & mLineLooper & " Col B + L" & mLineLooper & _
                               " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E\par}")
                End If

                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", mSumLine)

                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col H (" & mSum1.ToString & ") <> L" & mLineLooper & " Col F" & _
                                " + L" & mLineLooper & " Col G (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col H = L" & mLineLooper & " Col F + L" & mLineLooper & " Col G\par}")
                End If

            Next

            'Calculate/compare the values vertically
            For mColLooper = 1 To 7

                mCol = "C" & CStr(mColLooper)

                mStartline = "220"
                mEndLine = "226"
                mSum1 = 0
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "227"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "228"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "229"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "230"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "231"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "232"
                mEndLine = "233"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "234"
                mEndLine = "235"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "236"
                mEndLine = "237"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mSumLine = "238"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L220 thru L237" & ColText(mCol).ToString & " (" & mSum1.ToString & ") <> L238" & _
                        ColText(mCol).ToString & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L220 thru L237" & ColText(mCol).ToString & " = L238" & ColText(mCol).ToString & "\par}")
                End If

            Next

            ' Calculate/compare the values horizontally
            For mLineLooper = 301 To 323
                mSum1 = 0
                mSumLine = mLineLooper
                For mColLooper = 1 To 4
                    mCol = "C" & CStr(mColLooper)
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                Next

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col F (" & mSum1.ToString & ") <> L" & mLineLooper & " Col B " & _
                               " + L" & mLineLooper & " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & "C4 (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col F = L" & mLineLooper & " Col B + L" & mLineLooper & _
                               " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E\par}")
                End If

                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", mSumLine)

                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col H (" & mSum1.ToString & ") <> L" & mLineLooper & " Col F " & _
                               " + L" & mLineLooper & " Col G (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col H = L" & mLineLooper & " Col F + L" & mLineLooper & " Col G\par}")
                End If

            Next

            'Calculate/compare the values vertically
            For mColLooper = 1 To 7

                mCol = "C" & CStr(mColLooper)

                mStartline = "301"
                mEndLine = "311"
                mSum1 = 0
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "312"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "313"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "314"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "315"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "316"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "317"
                mEndLine = "318"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "319"
                mEndLine = "320"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "321"
                mEndLine = "322"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mSumLine = "323"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L301 thru L322" & ColText(mCol).ToString & " (" & mSum1.ToString & ") <> L323" & _
                        ColText(mCol).ToString & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L301 thru L322" & ColText(mCol).ToString & " = L323" & ColText(mCol).ToString & "\par}")
                End If

            Next

            For mColLooper = 1 To 7
                ' Get the totals for lines 219Cx, 238Cx and 323Cx
                mSumLine = "219"
                mCol = "C" & CStr(mColLooper)
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "238"
                mCol = "C" & CStr(mColLooper)
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "323"
                mCol = "C" & CStr(mColLooper)
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                'Get the value for the total from L324Cx
                mSumLine = "324"
                mCol = "C" & CStr(mColLooper)
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                mThisCol = ColText(mCol)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L219" & mThisCol & "+ L238" & mThisCol & "+ L323" & mThisCol & " (" & _
                               mSum1.ToString & ") <> L324" & mThisCol & ")\par}")
                Else
                    .WriteLine("{\pard\tab L219" & mThisCol & " + L238" & mThisCol & " + L323" & mThisCol & _
                               " = L324" & mThisCol & "\par}")
                End If
            Next

            ' Calculate/compare the values horizontally
            For mLineLooper = 401 To 419
                mSum1 = 0
                mSumLine = mLineLooper
                For mColLooper = 1 To 4
                    mCol = "C" & CStr(mColLooper)
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                Next

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col F (" & mSum1.ToString & ") <> L" & mLineLooper & " Col B " & _
                               " + L" & mLineLooper & " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & "C4 (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col F = L" & mLineLooper & " Col B + L" & mLineLooper & _
                               " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E\par}")
                End If

                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", mSumLine)

                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col H (" & mSum1.ToString & ") <> L" & mLineLooper & " Col F " & _
                               " + L" & mLineLooper & " Col G (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col H = L" & mLineLooper & " Col F + L" & mLineLooper & " Col G\par}")
                End If

            Next

            'Calculate/compare the values vertically
            For mColLooper = 1 To 7

                mCol = "C" & CStr(mColLooper)

                mStartline = "401"
                mEndLine = "416"
                mSum1 = 0
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "417"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "418"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mSumLine = "419"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L401 thru L418" & ColText(mCol).ToString & " (" & mSum1.ToString & ") <> L419" & _
                        ColText(mCol).ToString & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L401 thru L418" & ColText(mCol).ToString & " = L419" & ColText(mCol).ToString & "\par}")
                End If

            Next

            ' Calculate/compare the values horizontally
            For mLineLooper = 420 To 435
                mSum1 = 0
                mSumLine = mLineLooper
                For mColLooper = 1 To 4
                    mCol = "C" & CStr(mColLooper)
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                Next

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col F (" & mSum1.ToString & ") <> L" & mLineLooper & " Col B " & _
                               " + L" & mLineLooper & " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col F = L" & mLineLooper & " Col B + L" & mLineLooper & _
                               " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E\par}")
                End If

                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", mSumLine)

                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col H (" & mSum1.ToString & ") <> L" & mLineLooper & " Col F " & _
                               " + L" & mLineLooper & " Col G (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col H = L" & mLineLooper & " Col F + L" & mLineLooper & " Col G\par}")
                End If

            Next

            'Calculate/compare the values vertically
            For mColLooper = 1 To 7

                mCol = "C" & CStr(mColLooper)

                mStartline = "420"
                mEndLine = "432"
                mSum1 = 0
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "433"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "434"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mSumLine = "435"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L420 thru L434" & ColText(mCol).ToString & " (" & mSum1.ToString & ") <> L435" & _
                        ColText(mCol).ToString & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L420 thru L434" & ColText(mCol).ToString & " = L435" & ColText(mCol).ToString & "\par}")
                End If

            Next

            ' Calculate/compare the values horizontally
            For mLineLooper = 501 To 506
                mSum1 = 0
                mSumLine = mLineLooper
                For mColLooper = 1 To 4
                    mCol = "C" & CStr(mColLooper)
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                Next

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col F (" & mSum1.ToString & ") <> L" & mLineLooper & " Col B " & _
                               " + L" & mLineLooper & " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col F = L" & mLineLooper & " Col B + L" & mLineLooper & _
                               " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E\par}")
                End If

                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", mSumLine)

                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col H (" & mSum1.ToString & ") <> L" & mLineLooper & " Col F " & _
                               " + L" & mLineLooper & " Col G (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col H = L" & mLineLooper & " Col F + L" & mLineLooper & " Col G\par}")
                End If

            Next

            'Calculate/compare the values vertically
            For mColLooper = 1 To 7

                mCol = "C" & CStr(mColLooper)

                mStartline = "501"
                mEndLine = "505"
                mSumLine = "506"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            Next

            ' Calculate/compare the values horizontally
            For mLineLooper = 507 To 517
                mSum1 = 0
                mSumLine = mLineLooper
                For mColLooper = 1 To 4
                    mCol = "C" & CStr(mColLooper)
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                Next

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col F (" & mSum1.ToString & ") <> L" & mLineLooper & " Col B " & _
                               " + L" & mLineLooper & " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col F = L" & mLineLooper & " Col B + L" & mLineLooper & _
                               " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E\par}")
                End If

                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", mSumLine)

                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col H (" & mSum1.ToString & ") <> L" & mLineLooper & " Col F " & _
                               " + L" & mLineLooper & " Col G (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col H = L" & mLineLooper & " Col F + L" & mLineLooper & " Col G\par}")
                End If

            Next

            'Calculate/compare the values vertically
            For mColLooper = 1 To 7

                mCol = "C" & CStr(mColLooper)

                mStartline = "507"
                mEndLine = "514"
                mSum1 = 0
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "515"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "516"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mSumLine = "517"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L507 thru L516" & ColText(mCol).ToString & " (" & mSum1.ToString & ") <> L517" & _
                        ColText(mCol).ToString & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L507 thru L516" & ColText(mCol).ToString & " = L517" & ColText(mCol).ToString & "\par}")
                End If

            Next

            ' Calculate/compare the values horizontally
            For mLineLooper = 518 To 527
                mSum1 = 0
                mSumLine = mLineLooper
                For mColLooper = 1 To 4
                    mCol = "C" & CStr(mColLooper)
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                Next

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col F (" & mSum1.ToString & ") <> L" & mLineLooper & " Col B " & _
                               " + L" & mLineLooper & " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col F = L" & mLineLooper & " Col B + L" & mLineLooper & _
                               " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E\par}")
                End If

                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", mSumLine)

                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col H (" & mSum1.ToString & ") <> L" & mLineLooper & " Col F " & _
                               " + L" & mLineLooper & " Col G (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col H = L" & mLineLooper & " Col F + L" & mLineLooper & " Col G\par}")
                End If

            Next

            'Calculate/compare the values vertically
            For mColLooper = 1 To 7

                mCol = "C" & CStr(mColLooper)

                mStartline = "518"
                mEndLine = "524"
                mSum1 = 0
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "525"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "526"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mSumLine = "527"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L518 thru L526" & ColText(mCol).ToString & " (" & mSum1.ToString & ") <> L527" & _
                        ColText(mCol).ToString & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L518 thru L526" & ColText(mCol).ToString & " = L527" & ColText(mCol).ToString & "\par}")
                End If

            Next

            For mColLooper = 1 To 7
                ' Get the totals for lines 419Cx, 435Cx, 506Cx, 517Cx and 527Cx
                mSumLine = "419"
                mCol = "C" & CStr(mColLooper)
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "435"
                mCol = "C" & CStr(mColLooper)
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "506"
                mCol = "C" & CStr(mColLooper)
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "517"
                mCol = "C" & CStr(mColLooper)
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "527"
                mCol = "C" & CStr(mColLooper)
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                'Get the value for the total from L528Cx
                mSumLine = "528"
                mCol = "C" & CStr(mColLooper)
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                mThisCol = ColText(mCol)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L419" & mThisCol & _
                               " + L435" & mThisCol & _
                               " + L506" & mThisCol & _
                               " + L517" & mThisCol & _
                               " + L527" & mThisCol & " (" & _
                               mSum1.ToString & ") <> L528" & mThisCol & ")\par}")
                Else
                    .WriteLine("{\pard\tab L419" & mThisCol & _
                               " + L435" & mThisCol & _
                               " + L506" & mThisCol & _
                               " + L517" & mThisCol & _
                               " + L527" & mThisCol & _
                               " = L528" & mThisCol & "\par}")
                End If
            Next

            ' Calculate/compare the values horizontally
            For mLineLooper = 601 To 619
                mSum1 = 0
                mSumLine = mLineLooper
                For mColLooper = 1 To 4
                    mCol = "C" & CStr(mColLooper)
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                Next

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col F (" & mSum1.ToString & ") <> L" & mLineLooper & " Col B " & _
                               " + L" & mLineLooper & " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col F = L" & mLineLooper & " Col B + L" & mLineLooper & _
                               " Col C + L" & mLineLooper & " Col D + L" & mLineLooper & " Col E\par}")
                End If

                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", mSumLine)

                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper & " Col H (" & mSum1.ToString & ") <> L" & mLineLooper & " Col F " & _
                               " + L" & mLineLooper & " Col G (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper & " Col H = L" & mLineLooper & " Col F + L" & mLineLooper & " Col G\par}")
                End If

            Next

            'Calculate/compare the values vertically
            For mColLooper = 1 To 7

                mCol = "C" & CStr(mColLooper)

                mStartline = "601"
                mEndLine = "616"
                mSum1 = 0
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine)

                mStartline = "617"
                mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mStartline = "618"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)

                mSumLine = "619"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L601 thru L618" & ColText(mCol).ToString & " (" & mSum1.ToString & ") <> L619" & _
                        ColText(mCol).ToString & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L601 thru L618" & ColText(mCol).ToString & " = L619" & ColText(mCol).ToString & "\par}")
                End If

            Next

            For mColLooper = 1 To 7
                ' Get the totals for lines 151Cx, 324Cx, 528Cx, 619Cx
                mSumLine = "151"
                mCol = "C" & CStr(mColLooper)
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "324"
                mCol = "C" & CStr(mColLooper)
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "528"
                mCol = "C" & CStr(mColLooper)
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "619"
                mCol = "C" & CStr(mColLooper)
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                'Get the value for the total from L528Cx
                mSumLine = "620"
                mCol = "C" & CStr(mColLooper)
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                mThisCol = ColText(mCol)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L151" & mThisCol & _
                               " + L324" & mThisCol & _
                               " + L528" & mThisCol & _
                               " + L619" & mThisCol & " (" & _
                               mSum1.ToString & ") <> L528" & mThisCol & ")\par}")
                Else
                    .WriteLine("{\pard\tab L151" & mThisCol & _
                               " + L324" & mThisCol & _
                               " + L528" & mThisCol & _
                               " + L619" & mThisCol & _
                               " = L620" & mThisCol & "\par}")
                End If
            Next

            'Now we crosscheck to other schedules
            mCol = "C7"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "620"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol = "C1"
            mURCSCode = Get_URCS_Code("210")
            mSumLine = "14"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L620 Col H (" & mSum1.ToString & ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L14 Col B" & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L620 Col H = Sch " & Get_URCS_Schedule(mURCSCode).ToString & " L14 Col B \par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "620"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C3"
            mURCSCode = Get_URCS_Code("210")
            mSumLine2 = "14"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol).ToString & " = Sch " & Get_URCS_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2).ToString & "\par}")
            End If

            mCol = "C6"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "620"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C4"
            mURCSCode = Get_URCS_Code("210")
            mSumLine2 = "14"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol).ToString & " (" & mSum1.ToString & _
                           ") <> Sch " & Get_URCS_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2).ToString & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol).ToString & " = Sch " & Get_URCS_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2).ToString & "\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "231"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C1"
            mURCSCode = Get_URCS_Code("414")
            mSumLine2 = "19"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            mCol2 = "C2"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            mCol2 = "C3"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & " Cols B thru D" & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & " Cols B thru D\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "230"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C4"
            mURCSCode = Get_URCS_Code("414")
            mSumLine2 = "19"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            mCol2 = "C5"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            mCol2 = "C6"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & " Cols E thru G" & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & " Cols E thru G\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "507"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C9"
            mURCSCode = Get_URCS_Code("417")
            mSumLine2 = "1"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "508"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C9"
            mURCSCode = Get_URCS_Code("417")
            mSumLine2 = "2"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & "= Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "509"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C9"
            mURCSCode = Get_URCS_Code("417")
            mSumLine2 = "3"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "510"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C9"
            mURCSCode = Get_URCS_Code("417")
            mSumLine2 = "4"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "511"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C9"
            mURCSCode = Get_URCS_Code("417")
            mSumLine2 = "5"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "512"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C9"
            mURCSCode = Get_URCS_Code("417")
            mSumLine2 = "6"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "513"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C9"
            mURCSCode = Get_URCS_Code("417")
            mSumLine2 = "7"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "514"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C9"
            mURCSCode = Get_URCS_Code("417")
            mSumLine2 = "8"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "515"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C9"
            mURCSCode = Get_URCS_Code("417")
            mSumLine2 = "9"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "516"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C9"
            mURCSCode = Get_URCS_Code("417")
            mSumLine2 = "10"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "517"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C9"
            mURCSCode = Get_URCS_Code("417")
            mSumLine2 = "11"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            'mCol = "C1"
            'mURCSCode = Get_URCS_Code("410")
            'mSumLine = "4"
            'mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            'mCol2 = "C1"
            'mURCSCode = Get_URCS_Code("210")
            'mSumLine2 = "47"
            'mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            'If mSum1 <> mSum2 Then
            '    .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
            '               ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
            '               " (" & mSum2.ToString & ")\par}")
            'Else
            '    .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
            '               " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            'End If

            mSum1 = 0
            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            For mLineLooper = 136 To 138
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mLineLooper)
            Next

            mCol2 = "C1"
            mURCSCode = Get_URCS_Code("412")
            mSumLine2 = "29"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L" & mSumLine.ToString & ColText(mCol) & " (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L" & mSumLine.ToString & ColText(mCol) & " = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mSum1 = 0
            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            For mLineLooper = 118 To 123
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mLineLooper)
            Next
            For mLineLooper = 130 To 135
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mLineLooper)
            Next

            mCol2 = "C2"
            mURCSCode = Get_URCS_Code("412")
            mSumLine2 = "29"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L118 Col F thru L123 Col F and L130 Col F thru L135 Col F (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L" & mSumLine2.ToString & ColText(mCol2) & _
                           " (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L118 Col F thru L123 Col F and L130 Col F thru L135 Col F = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L" & mSumLine2.ToString & ColText(mCol2) & "\par}")
            End If

            mSum1 = 0
            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "207"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "208"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "211"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "212"
            mSum1 = mSum1 + Math.Abs(Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine))

            mCol2 = "C5"
            mURCSCode = Get_URCS_Code("415")
            mSumLine2 = "5"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            mSumLine2 = "38"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L207 Col F + L208 Col F + L211 Col F + L212 Col F (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L5 Col F + L38 Col F (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L207 Col F + L208 Col F + L211 Col F + L212 Col F = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L5 Col F + L38 Col F\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "226"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "227"
            mSum1 = mSum1 + Math.Abs(Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine))

            mCol2 = "C5"
            mURCSCode = Get_URCS_Code("415")
            mSumLine2 = "24"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            mSumLine2 = "39"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine2)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L226 Col F + L227 Col F (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L24 Col F + L39 Col F (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L226 Col F + L227 Col F = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L24 Col F + L39 Col F\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "311"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "312"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "315"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "316"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol2 = "C5"
            mURCSCode = Get_URCS_Code("415")
            mSumLine2 = "32"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            mSumLine2 = "35"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine2)
            mSumLine2 = "36"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine2)
            mSumLine2 = "37"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine2)
            mSumLine2 = "40"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine2)
            mSumLine2 = "41"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine2)
            mURCSCode = Get_URCS_Code("414")
            mSumLine2 = "24"
            For mColLooper = 1 To 3
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next
            For mColLooper = 4 To 6
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L311 Col F + L312 Col F + L315 Col F + L316 Col F (" & mSum1.ToString & _
                           ") <> Sch 415 L32 Col F + L35 Col F + L36 Col F + L37 Col F + L40 Col F + L41 Col F - Sch 414 L24 Col B thru Col D + L24 Col E to Col G (" & _
                           mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L311 Col F + L312 Col F + L315 Col F + L316C5 = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " Sch 415 L32 Col F + L35 Col F + L36 Col F + L37 Col F + L40 Col F + L41C5 - Sch 414 L24 Col B thru Col D + L24 Col E to Col G\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "213"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mURCSCode = Get_URCS_Code("415")
            mSumLine2 = "5"
            mSum2 = 0
            For mColLooper = 2 To 3
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next
            mSumLine2 = "38"
            For mColLooper = 2 To 3
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L213 Col F (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L5 Col C + L5 Col D + L38 Col C + L38 Col D (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L213 Col F = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L5 Col C + L5 Col D + L38 Col C + L38 Col D\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "232"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mURCSCode = Get_URCS_Code("415")
            mSumLine2 = "24"
            mSum2 = 0
            For mColLooper = 2 To 3
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next
            mSumLine2 = "39"
            For mColLooper = 2 To 3
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L232 Col F (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L24 Col C + L24 Col D + L39 Col C + L39 Col D (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L232 Col F = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L24 Col C + L24 Col D + L39 Col C + L39 Col D\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "317"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mURCSCode = Get_URCS_Code("415")
            mSumLine2 = "32"
            mSum2 = 0
            For mColLooper = 2 To 3
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next
            mSumLine2 = "35"
            For mColLooper = 2 To 3
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next
            mSumLine2 = "36"
            For mColLooper = 2 To 3
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next
            mSumLine2 = "37"
            For mColLooper = 2 To 3
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next
            mSumLine2 = "40"
            For mColLooper = 2 To 3
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next
            mSumLine2 = "41"
            For mColLooper = 2 To 3
                mCol2 = "C" & mColLooper.ToString
                mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine2)
            Next

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L317 Col F (" & mSum1.ToString & _
                           ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L32 Col C + L32 Col D + L35 Col C + L35 Col D + L36 Col C + L36 Col D + L37 Col C + L37 Col D + L40 Col C + L40 Col D + L41 Col C + L41 Col D (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L317C5 = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L32 Col C + L32 Col D + L35 Col C + L35 Col D + L36 Col C + L36 Col D + L37 Col C + L37 Col D + L40 Col C + L40 Col D + L41 Col C + L41 Col D\par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "202"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "203"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "216"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol = "C1"
            mURCSCode = Get_URCS_Code("415")
            mSumLine = "5"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "38"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            If mSum1 < mSum2 Then
                mCol = "C5"
                mURCSCode = Get_URCS_Code("410")
                mSumLine = "216"
                If (mSum2 - mSum1) > Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine) Then
                    mURCSCode = Get_URCS_Code("415")
                    .WriteLine("{\pard\tab \b L202 Col F + L203 Col F + L216 Col F (" & mSum1.ToString & _
                           ") Out of Range with Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L5 Col B + L36 Col B (" & mSum2.ToString & ")\par}")
                Else
                    mURCSCode = Get_URCS_Code("415")
                    .WriteLine("{\pard\tab L202 Col F + L203 Col F + L216 Col F Within Range to Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L5 Col B + L36 Col B \par}")
                End If
            Else
                mURCSCode = Get_URCS_Code("415")
                .WriteLine("{\pard\tab L202 Col F + L203 Col F + L216 Col F = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                       " L5 Col B + L36 Col B \par}")
            End If

            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSumLine = "221"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "222"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "235"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol = "C1"
            mURCSCode = Get_URCS_Code("415")
            mSumLine = "24"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSumLine = "39"
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            If mSum1 < mSum2 Then
                mCol = "C5"
                mURCSCode = Get_URCS_Code("410")
                mSumLine = "235"
                If (mSum2 - mSum1) > Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine) Then
                    mURCSCode = Get_URCS_Code("415")
                    .WriteLine("{\pard\tab \b L221 Col F + L222 Col F + L235 Col F (" & mSum1.ToString & _
                           ") Out of Range with Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L24 Col B + L39 Col B (" & mSum2.ToString & ")\par}")
                Else
                    mURCSCode = Get_URCS_Code("415")
                    .WriteLine("{\pard\tab L221 Col F + L222 Col F + L235 Col F Within Range to Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L24 Col B + L39 Col B \par}")
                End If
            Else
                mURCSCode = Get_URCS_Code("415")
                .WriteLine("{\pard\tab L221 Col F + L222 Col F + L235 Col F = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                       " L24 Col B + L39 Col B \par}")
            End If

            mSum1 = 0
            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "302", "307")
            mSumLine = "320"
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mCol = "C1"
            mURCSCode = Get_URCS_Code("415")
            mSumLine = "32"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "35", "37")
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "40", "41")

            If mSum1 < mSum2 Then
                mCol = "C5"
                mURCSCode = Get_URCS_Code("410")
                mSumLine = "320"
                If (mSum2 - mSum1) > Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine) Then
                    mURCSCode = Get_URCS_Code("415")
                    .WriteLine("{\pard\tab \b L302 Col F thru L307 Col F + L320 Col F (" & mSum1.ToString & _
                           ") Out of Range with Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L35 Col B thru L37 Col B + L40 Col B + L41 Col B (" & mSum2.ToString & ")\par}")
                Else
                    mURCSCode = Get_URCS_Code("415")
                    .WriteLine("{\pard\tab L302 Col F thru L307 Col F + L320C5 Within Range to Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                           " L35C1 thru L37 Col B + L40 Col B + L41 Col B \par}")
                End If
            Else
                mURCSCode = Get_URCS_Code("415")
                .WriteLine("{\pard\tab L302 Col F thru L307 Col F + L320 Col F = Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                       " L35C1 thru L37 Col B + L40 Col B + L41 Col B \par}")
            End If

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule412( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String
        Dim mStartline As String, mEndLine As String, mSumLine As String
        Dim mColLooper As Integer

        ' Find out what URCS Code is used for Sch 220
        mURCSCode = Get_URCS_Code("412")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 412\par}")

            For mColLooper = 1 To 3
                mCol = "C" & mColLooper.ToString
                mStartline = "1"
                mEndLine = "28"
                mSumLine = "29"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            mCol = "C1"
            mSumLine = "29"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

            mSum2 = 0
            mCol = "C5"
            mURCSCode = Get_URCS_Code("410")
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "136", "138")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L29 Col B (" & mSum1.ToString & _
                           ") <> Sch 410 L136 Col F thru L318 Col F (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L29 Col B = Sch 410 L136 Col F thru L318 Col F\par}")
            End If

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule414( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String, mCol2 As String
        Dim mStartline As String, mEndLine As String, mSumLine As String
        Dim mColLooper As Integer

        ' Find out what URCS Code is used for Sch 220
        mURCSCode = Get_URCS_Code("414")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 414\par}")

            For mColLooper = 1 To 6
                mCol = "C" & mColLooper.ToString
                mStartline = "1"
                mEndLine = "18"
                mSumLine = "19"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mColLooper = 1 To 6
                mCol = "C" & mColLooper.ToString
                mStartline = "20"
                mEndLine = "23"
                mSumLine = "24"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))
            Next

            For mColLooper = 1 To 6
                mCol = "C" & mColLooper.ToString
                mSumLine = "19"
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
                mSumLine = "24"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                mSumLine = "25"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L19" & ColText(mCol) & "+" & " L24" & ColText(mCol) & " (" & mSum1.ToString & _
                               ") <> Sch " & get_Urcs_Schedule(mURCSCode).ToString & " L25" & ColText(mCol) & " (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L19" & ColText(mCol) & "+" & " L24" & ColText(mCol) & "= Sch " & get_Urcs_Schedule(mURCSCode).ToString & _
                               " L25" & ColText(mCol) & "\par}")
                End If

            Next

            mSum1 = 0
            For mColLooper = 1 To 3
                mURCSCode = Get_URCS_Code("414")
                mCol = "C" & mColLooper.ToString
                mSumLine = "19"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            Next

            mURCSCode = Get_URCS_Code("410")
            mCol2 = "C5"
            mSumLine = "231"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L19 Col B thru L19 Col D" & " (" & mSum1.ToString & _
                           ") <> Sch 410 L231" & ColText(mCol2) & "(" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L19 Col B thru L19 Col D = Sch 410 L231" & ColText(mCol2) & "\par}")
            End If

            mSum1 = 0

            For mColLooper = 4 To 6
                mURCSCode = Get_URCS_Code("414")
                mCol = "C" & mColLooper.ToString
                mSumLine = "19"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mSumLine)
            Next

            mURCSCode = Get_URCS_Code("410")
            mCol2 = "C5"
            mSumLine = "230"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol2, mSumLine)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L19 Col E thru L19 Col G" & " (" & mSum1.ToString & _
                           ") <> Sch 410 L230" & ColText(mCol2) & "(" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L19 Col E thru L19 Col G = Sch 410 L230" & ColText(mCol2) & "\par}")
            End If

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule415( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String
        Dim mColLooper As Integer

        ' Find out what URCS Code is used for Sch 415
        mURCSCode = Get_URCS_Code("415")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 415\par}")

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, "1", "4", "5"))
            Next

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, "6", "23", "24"))
            Next

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, "25", "31", "32"))
            Next

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, "33", "34", "35"))
            Next

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, "36", "41", "42"))
            Next

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "5")
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "24")
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "32")
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "35")
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "42")

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "43")

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L5" & ColText(mCol) & _
                               "+ L24" & ColText(mCol) & _
                               "+ L32" & ColText(mCol) & _
                               "+ L35" & ColText(mCol) & _
                               "+ L42" & ColText(mCol) & "(" & mSum1.ToString & _
                               ") <> L43" & ColText(mCol) & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L5" & ColText(mCol) & _
                               "+ L24" & ColText(mCol) & _
                               "+ L32" & ColText(mCol) & _
                               "+ L35" & ColText(mCol) & _
                               "+ L42" & ColText(mCol) & _
                               ") = L43" & ColText(mCol) & "\par}")
                End If

            Next

            mURCSCode = Get_URCS_Code("415")
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", "5")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", "38")

            mURCSCode = Get_URCS_Code("410")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "202", "203")
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "216")
            
            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L5 Col B + L38 Col B (" & mSum1.ToString & ") <> " & _
                           "Sch 410 L202 Col F + L203 Col F + L216 Col F (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L5 Col B + L38 Col B =  " & _
                           "Sch 410 L202 Col F + L203 Col F + L216 Col F")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", "24")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", "39")

            mURCSCode = Get_URCS_Code("410")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C5", "221", "222")
            mSum2 = mSum2 + Math.Abs(Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C5", "235"))

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L24 Col B + L39 Col B (" & mSum1.ToString & ") <> " & _
                           "Sch 410 L221 Col F + L222 Col F + L235 Col F (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L24 Col B + L39 Col B =  " & _
                           "Sch 410 L221 Col F + L222 Col F + L235 Col F")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", "32")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", "35", "37")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "B", "40", "41")

            mURCSCode = Get_URCS_Code("410")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "302", "307")
            mSum2 = mSum2 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "320")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L32 Col B + L35 thru 37 Col B + L40 and L41 Col B (" & mSum1.ToString & ") <> " & _
                           "Sch 410 L302 thru L307 Col F + L320 Col F (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L32 Col B + L35 thru 37 Col B + L40 and L41 Col B =  " & _
                           "Sch 410 L302 thru L307 Col F + L320 Col F")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mCol = "C2"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C", "5")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "D", "5")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C", "38")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "D", "38")

            mURCSCode = Get_URCS_Code("410")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "213")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L5 Cols C and D + L38 Cols C and D (" & mSum1.ToString & ") <> " & _
                           "Sch 410 L213 Col F (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L5 Cols C and D + L38 Cols C and D = Sch 410 L213 Col F")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mCol = "C2"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C", "24")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "D", "24")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C", "39")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "D", "39")

            mURCSCode = Get_URCS_Code("410")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "232")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L24 Cols C and D + L39 Cols C and D (" & mSum1.ToString & ") <> " & _
                           "Sch 410 L232 Col F (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L24 Cols C and D + L39 Cols C and D = Sch 410 L232 Col F")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C", "32")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "D", "32")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C", "35", "37")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "D", "35", "37")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C", "40", "41")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "D", "40", "41")

            mURCSCode = Get_URCS_Code("410")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "317")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L32 Cols C and D + L35 the L37 Cols C and D + L40 and L41 Cols C and D (" & mSum1.ToString & ") <> " & _
                           "Sch 410 L317 Col F (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L32 Cols C and D + L35 the L37 Cols C and D + L40 and L41 Cols C and D = Sch 410 L317 Col F")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "32")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "38")

            mURCSCode = Get_URCS_Code("410")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "207", "208")
            mSum2 = mSum2 + Math.Abs(Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "211", "212"))

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L32 Col F + L38 Col F (" & mSum1.ToString & ") <> " & _
                           "Sch 410 L207 and L208 Col F + L211 and L 212 Col F (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L32 Col F + L38 Col F = Sch 410 L207 and L208 Col F + L211 and L 212 Col F")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "24")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "39")

            mURCSCode = Get_URCS_Code("410")
            mSum2 = Math.Abs(Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "F", "226", "227"))

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L24 Col F + L39 Col F (" & mSum1.ToString & ") <> " & _
                           "Sch 410 L226 Col F + L227 Col F (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L24 Col F + L39 Col F = Sch 410 L226 Col F + L227 Col F")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "I", "5")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "J", "5")

            mURCSCode = Get_URCS_Code("335")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "31")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L5 Col I + Col J (" & mSum1.ToString & ") <> " & _
                           "Sch 335 L31 Col G (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L5 Col I + Col J = Sch 335 L31 Col G")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "I", "24")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "J", "24")

            mURCSCode = Get_URCS_Code("335")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "32")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L24 Col I + Col J (" & mSum1.ToString & ") <> " & _
                           "Sch 335 L32 Col G (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L24 Col I + Col J = Sch 335 L32 Col G")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "I", "32")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "J", "32")

            mURCSCode = Get_URCS_Code("335")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "34")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L32 Col I + Col J (" & mSum1.ToString & ") <> " & _
                           "Sch 335 L34 Col G (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L32 Col I + Col J = Sch 335 L34 Col G")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "I", "35")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "J", "35")

            mURCSCode = Get_URCS_Code("335")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "35")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L35 Col I + Col J (" & mSum1.ToString & ") <> " & _
                           "Sch 335 L35 Col G (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L35 Col I + Col J = Sch 335 L35 Col G")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "I", "36")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "J", "36")

            mURCSCode = Get_URCS_Code("335")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "33")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L36 Col I + Col J (" & mSum1.ToString & ") <> " & _
                           "Sch 335 L33 Col G (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L36 Col I + Col J = Sch 335 L33 Col G")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "I", "37")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "J", "37")

            mURCSCode = Get_URCS_Code("335")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "38")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L37 Col I + Col J (" & mSum1.ToString & ") <> " & _
                           "Sch 335 L38 Col G (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L37 Col I + Col J = Sch 335 L38 Col G")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "I", "41")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "J", "41")

            mURCSCode = Get_URCS_Code("335")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "36", "37")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L41 Col I + Col J (" & mSum1.ToString & ") <> " & _
                           "Sch 335 L36 & L37 Col G (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L41 Col I + Col J = Sch 335 L36 & L37 Col G")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "I", "38")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "J", "38")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "I", "39")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "J", "39")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "I", "40")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "J", "40")

            mURCSCode = Get_URCS_Code("335")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "26")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L38 thru L40 Cols I + Col J (" & mSum1.ToString & ") <> " & _
                           "Sch 335 L26 Col G (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L38 thru L40 Cols I + Col J = Sch 335 L26 Col G")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "5")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "5")

            mURCSCode = Get_URCS_Code("330")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "31")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L5 Col G + Col H (" & mSum1.ToString & ") <> " & _
                           "Sch 330 L31 Col H (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L5 Col G + Col H = Sch 330 L31 Col H")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "24")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "24")

            mURCSCode = Get_URCS_Code("330")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "32")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L24 Col G + Col H (" & mSum1.ToString & ") <> " & _
                           "Sch 330 L32 Col H (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L24 Col G + Col H = Sch 330 L32 Col H")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "32")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "32")

            mURCSCode = Get_URCS_Code("330")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "34")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L32 Col G + Col H (" & mSum1.ToString & ") <> " & _
                           "Sch 330 L34 Col H (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L32 Col G + Col H = Sch 330 L34 Col H")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "35")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "35")

            mURCSCode = Get_URCS_Code("330")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "35")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L35 Col G + Col H (" & mSum1.ToString & ") <> " & _
                           "Sch 330 L35 Col H (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L35 Col G + Col H = Sch 330 L35 Col H")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "36")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "36")

            mURCSCode = Get_URCS_Code("330")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "33")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L36 Col G + Col H (" & mSum1.ToString & ") <> " & _
                           "Sch 330 L33 Col H (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L36 Col G + Col H = Sch 330 L33 Col H")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "37")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "37")

            mURCSCode = Get_URCS_Code("330")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "38")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L37 Col G + Col H (" & mSum1.ToString & ") <> " & _
                           "Sch 330 L38 Col H (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L37 Col G + Col H = Sch 330 L38 Col H")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "41")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "41")

            mURCSCode = Get_URCS_Code("330")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "36", "37")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L41 Col G + Col H (" & mSum1.ToString & ") <> " & _
                           "Sch 330 L36 & L37 Col H (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L41 Col G + Col H = Sch 330 L36 & L37 Col H")
            End If

            .WriteLine("\par}")

            mURCSCode = Get_URCS_Code("415")

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "38")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "38")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "39")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "39")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "G", "40")
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "40")

            mURCSCode = Get_URCS_Code("330")
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", "27")

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L38 thru L40 Cols G + Col H (" & mSum1.ToString & ") <> " & _
                           "Sch 330 L27 Col H (" & mSum2.ToString & ")")
            Else
                .WriteLine("{\pard\tab L38 thru L40 Cols G + Col H = Sch 330 L27 Col H")
            End If

            .WriteLine("\par}")

            .WriteLine("{\pard\Line\Line\par}")
            
        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule416( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String
        Dim mColLooper As Integer

        ' Find out what URCS Code is used for Sch 220
        mURCSCode = Get_URCS_Code("416")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 416\par}")

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, "1", "4", "5"))
            Next

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, "6", "9", "10"))
            Next

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, "11", "14", "15"))
            Next

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, "16", "19", "20"))
            Next

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, "21", "24", "25"))
            Next

            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "5")
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "10")
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "15")
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "20")
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "25")

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, "26")

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L5" & ColText(mCol) & _
                               "+ L10" & ColText(mCol) & _
                               "+ L15" & ColText(mCol) & _
                               "+ L20" & ColText(mCol) & _
                               "+ L25" & ColText(mCol) & " (" & mSum1.ToString & _
                               ") <> L26" & ColText(mCol) & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L5" & ColText(mCol) & _
                               "+ L10" & ColText(mCol) & _
                               "+ L15" & ColText(mCol) & _
                               "+ L20" & ColText(mCol) & _
                               "+ L25" & ColText(mCol) & _
                               ") = L26" & ColText(mCol) & "\par}")
                End If

            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule417( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String
        Dim mColLooper As Integer, mLineLooper As Integer

        ' Find out what URCS Code is used for this schedule
        mURCSCode = Get_URCS_Code("417")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 417\par}")

            'Check the totals vertically
            For mColLooper = 1 To 9
                mCol = "C" & mColLooper.ToString
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, "1", "10", "11"))
            Next

            'Check the totals horizontally
            For mLineLooper = 1 To 11
                'mStartline = mLineLooper
                mSum1 = 0
                For mColLooper = 1 To 8
                    mCol = "C" & mColLooper.ToString
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mLineLooper)
                Next
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "J", mLineLooper)
                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper.ToString & " Col B thru Col I (" & mSum1.ToString & _
                               ") <> Col J (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper.ToString & " Col B thru Col I = Col J\par}")
                End If

            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule450( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String
        Dim mStartline As String, mEndLine As String, mSumLine As String
        Dim mLineLooper As Integer

        ' Find out what URCS Code is used for this schedule
        mURCSCode = Get_URCS_Code("450")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 450\par}")

            mCol = "C1"
            mStartline = "2"
            mEndLine = "3"
            mSumLine = "4"
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mCol, mStartline, mEndLine, mSumLine))

            mLineLooper = "1"
            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mLineLooper)
            For mLineLooper = 5 To 9
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mLineLooper)
            Next

            mLineLooper = "10"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mLineLooper)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L1 Col B + (L5" & ColText("C1") & "thru L9" & ColText("C1") & ")" & " (" & mSum1.ToString & _
                           ") <> L10" & ColText("C1") & "(" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L1 Col B + (L5" & ColText("C1") & "thru L9" & ColText("C1") & ") = L10 Col B \par}")
            End If

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule510( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String, mURCSCode2 As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mSum1Sng As Single, mSum2Sng As Single
        Dim mSumLine As String, mSumLine2 As String
        Dim mLineLooper As Integer

        ' Find out what URCS Code is used for this schedule
        mURCSCode = Get_URCS_Code("510")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 510\par}")

            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C1", 1, 8, 9))
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C1", 10, 11, 12))
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C1", 16, 17, 9))
            .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C1", 25, 26, 21))

            mURCSCode2 = Get_URCS_Code("200")
            mSumLine2 = 0
            For mLineLooper = 1 To 8
                mSumLine = mLineLooper
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", mSumLine)

                Select Case mLineLooper
                    Case 1
                        mSumLine2 = 30
                    Case 2
                        mSumLine2 = 39
                    Case 3
                        mSumLine2 = 41
                    Case 4
                        mSumLine2 = 42
                    Case 5
                        mSumLine2 = 43
                    Case 6
                        mSumLine2 = 44
                    Case 7
                        mSumLine2 = 45
                    Case 8
                        mSumLine2 = 46
                End Select

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode2, "C1", mSumLine2)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper.ToString & " Col B (" & mSum1.ToString & _
                               ") <> Sch 200 L" & mSumLine2.ToString & " Col B " & " (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper.ToString & " Col B = Sch 200 L" & mSumLine2.ToString & " Col B \par}")
                End If

            Next

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 9)
            mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 12)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 15)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L9 Col B - L12 Col B (" & mSum1.ToString & _
                           ") <> L15 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L9 Col B - L12 Col B = L15 Col B \par}")
            End If

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 13)
            mSum1 = mSum1 * Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 15)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 10)

            If mSum2 <> 0 Then
                mSum1 = mSum1 / mSum2
            Else
                mSum1 = 0
            End If

            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 16)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b (L13 Col B * L15 Col B)/L10 Col B (" & mSum1.ToString & _
                           ") <> L16 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab (L13C1*L15C1)/L10 Col B = L16 Col B \par}")
            End If

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 13)
            mSum1 = mSum1 * Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 15)
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 10)

            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 16)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b (L13 Col B * L15 Col B) + L10 Col B (" & mSum1.ToString & _
                           ") <> L16 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab (L13 Col B * L15 Col B) + L10 Col B = L16 Col B \par}")
            End If

            mURCSCode2 = Get_URCS_Code("210")
            mSumLine2 = 0
            For mLineLooper = 18 To 20
                mSumLine = mLineLooper
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", mSumLine)

                Select Case mLineLooper
                    Case 18
                        mSumLine2 = 42
                    Case 19
                        mSumLine2 = 44
                    Case 20
                        mSumLine2 = 22
                End Select

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode2, "C1", mSumLine2)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper.ToString & " Col B (" & mSum1.ToString & _
                               ") <> Sch 210 L" & mSumLine2.ToString & " Col B " & " (" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper.ToString & " Col B = Sch 210 L" & mSumLine2.ToString & " Col B \par}")
                End If

            Next

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 18)

            If Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 19) > 0 Then
                mSum1 = mSum1 / Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 19)
            End If

            mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 20)

            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 21)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b (L18 Col B / L19 Col B) - L20 Col B (" & mSum1.ToString & _
                           ") <> L21 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab (L18 Col B / L19 Col B) - L20 Col B = L21 Col B \par}")
            End If

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 21)
            mSum1 = mSum1 - Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 22, 23)

            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 24)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L21 Col B - (L22 Col B + L23 Col B) (" & mSum1.ToString & _
                           ") <> L24 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L21 Col B - (L22 Col B + L23 Col B) = L24 Col B \par}")
            End If

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 24)
            If Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 13) > 0 Then
                mSum1 = mSum1 * (Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 13) / 10000)
            Else
                mSum1 = 0
            End If
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 22)

            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 25)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L22 Col B + (L24 Col B * L13 Col B) (" & mSum1.ToString & _
                           ") <> L25 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L22 Col B + (L24 Col B * L13 Col B) = L25 Col B \par}")
            End If

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 24)
            If Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 14) > 0 Then
                mSum1 = mSum1 * (Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 14) / 10000)
            Else
                mSum1 = 0
            End If
            mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 23)

            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 26)

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L23 Col B + (L24 Col B * L14 Col B) (" & mSum1.ToString & _
                           ") <> L26 Col B (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L23 Col B + (L24 Col B * L14 Col B) = L26 Col B \par}")
            End If

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 25)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 16)

            If mSum2 <> 0 Then
                mSum1Sng = mSum1 / mSum2
            Else
                mSum1Sng = 0
            End If

            mSum2Sng = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 27) / 10000

            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L25 Col B / L16 Col B (" & CStr(mSum1Sng) & _
                           ") <> L27 Col B (" & CStr(mSum2) & ")\par}")
            Else
                .WriteLine("{\pard\tab L25 Col B / L16 Col B = L27 Col B \par}")
            End If

            mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 26)
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 17)

            If mSum2 <> 0 Then
                mSum1Sng = mSum1 / mSum2
            Else
                mSum1Sng = 0
            End If

            mSum2Sng = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C1", 28) / 10000

            If Math.Round(mSum1Sng, 4) <> Math.Round(mSum2Sng, 4) Then
                .WriteLine("{\pard\tab \b L26 Col B / L17 Col B (" & Math.Round(CDbl(mSum1Sng), 4) & _
                           ") <> L28 Col B (" & Math.Round(CDbl(mSum2Sng), 4) & ")\par}")
            Else
                .WriteLine("{\pard\tab L26 Col B / L17 Col B = L28 Col B \par}")
            End If

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule700( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String
        Dim mStartline As String
        Dim mColLooper As Integer

        ' Find out what URCS Code is used for this schedule
        mURCSCode = Get_URCS_Code("700")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 700\par}")

            'Check the totals horizontally

            mStartline = 57
            mSum1 = 0
            For mColLooper = 1 To 6
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, mStartline)
            Next
            mCol = "C7"
            mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "H", mStartline)
            If mSum1 <> mSum2 Then
                .WriteLine("{\pard\tab \b L57" & ColText("C1") & "thru" & ColText("C6") & "(" & mSum1.ToString & _
                           ") <> Col H (" & mSum2.ToString & ")\par}")
            Else
                .WriteLine("{\pard\tab L57" & ColText("C1") & "thru" & ColText("C6") & "= Col H \par}")
            End If

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule702( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mCol As String
        Dim mStartline As String, mSumLine As String
        Dim mColLooper As Integer, mLineLooper As Integer

        ' Find out what URCS Code is used for this schedule
        mURCSCode = Get_URCS_Code("702")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 702\par}")

            'Check the totals vertically
            For mColLooper = 1 To 7
                mSum1 = 0
                For mLineLooper = 1 To 31
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mColLooper, mLineLooper)
                Next
                mSumLine = 32
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mColLooper, mSumLine)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L1" & ColText("C" & mColLooper.ToString) & "thru L31" & ColText("C" & mColLooper.ToString) & "(" & mSum1.ToString & _
                               ") <> L32" & ColText("C" & mColLooper.ToString) & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L1" & ColText("C" & mColLooper.ToString) & "thru L31" & ColText("C" & mColLooper.ToString) & "= L32" & _
                               ColText("C" & mColLooper.ToString) & "\par}")
                End If

            Next

            'Check the totals horizontally
            For mLineLooper = 1 To 31
                mStartline = mLineLooper
                mSum1 = 0
                For mColLooper = 1 To 6
                    mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mColLooper, mStartline)
                Next
                mCol = "C7"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mCol, mStartline)
                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L" & mLineLooper.ToString & ColText("C1") & "thru" & ColText("C6") & "(" & mSum1.ToString & _
                               ") <>" & ColText("C7") & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L" & mLineLooper.ToString & ColText("C1") & "thru" & ColText("C6") & "=" & ColText("C7") & "\par}")
                End If
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule710( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mStartline As String, mEndLine As String, mSumLine As String
        Dim mColLooper As Integer

        ' Find out what URCS Code is used for this schedule
        mURCSCode = Get_URCS_Code("710")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 710\par}")

            For mColLooper = 1 To 11
                mStartline = "1"
                mEndLine = "3"
                mSumLine = "4"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))

                mStartline = "5"
                mEndLine = "6"
                mSumLine = "7"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))

                mStartline = "8"
                mEndLine = "9"
                mSumLine = "10"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))

                mStartline = "11"
                mEndLine = "13"
                mSumLine = "14"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))

                mStartline = "14"
                mEndLine = "15"
                mSumLine = "16"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))

                mStartline = "17"
                mEndLine = "22"
                mSumLine = "23"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))

                mStartline = "24"
                mEndLine = "27"
                mSumLine = "28"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))

                mSumLine = "23"
                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mColLooper, mSumLine)
                mSumLine = "28"
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mColLooper, mSumLine)
                mSumLine = "29"
                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mColLooper, mSumLine)
                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L23" & ColText("C" & mColLooper.ToString) & "+ L28" & ColText("C" & mColLooper.ToString) & "(" & mSum1.ToString & _
                               ") <> L29" & ColText("C" & mColLooper.ToString) & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L23" & ColText("C" & mColLooper.ToString) & "+ L28" & ColText("C" & mColLooper.ToString) & "= " & _
                               "L29" & ColText("C" & mColLooper.ToString) & "\par}")
                End If

                mStartline = "30"
                mEndLine = "34"
                mSumLine = "35"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))

            Next

            For mColLooper = 1 To 13
                mStartline = "36"
                mEndLine = "52"
                mSumLine = "53"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))

                mStartline = "53"
                mEndLine = "54"
                mSumLine = "55"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))

                mStartline = "56"
                mEndLine = "57"
                mSumLine = "58"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))

                mStartline = "59"
                mEndLine = "69"
                mSumLine = "70"
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, mColLooper, mStartline, mEndLine, mSumLine))
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub

    Sub R1_Schedule755( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mOutPutFileName As String)

        Dim mURCSCode As String
        Dim Filewriter As StreamWriter
        Dim mSum1 As Long, mSum2 As Long
        Dim mColLooper As Integer

        ' Find out what URCS Code is used for this schedule
        mURCSCode = Get_URCS_Code("755")

        ' Open the output file
        Filewriter = File.AppendText(mOutPutFileName)

        With Filewriter
            .WriteLine("{\pard \b \ul Sch 755\par}")

            For mColLooper = 1 To 2
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 2, 4, 5))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 5, 6, 7))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 8, 10, 11))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 11, 13, 14))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 15, 29, 30))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 31, 45, 46))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 47, 63, 64))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 65, 81, 82))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 85, 87, 88))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 98, 103, 104))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 105, 106, 107))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 108, 109, 110))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 111, 112, 113))

                mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 110)
                mSum1 = mSum1 + Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 113)

                mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 114)

                If mSum1 <> mSum2 Then
                    .WriteLine("{\pard\tab \b L110" & ColText("C" & mColLooper.ToString.ToString) & " + L113" & ColText("C" & mColLooper.ToString.ToString) & "(" & mSum1.ToString & _
                               ") <> L114" & ColText("C" & mColLooper.ToString.ToString) & "(" & mSum2.ToString & ")\par}")
                Else
                    .WriteLine("{\pard\tab L110" & ColText("C" & mColLooper.ToString.ToString) & " + L113" & ColText("C" & mColLooper.ToString.ToString) & "= " & _
                               "L114" & ColText("C" & mColLooper.ToString.ToString) & "\par}")
                End If

                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 126, 128, 129))
                .WriteLine(CompareSums(mYear, mRRICC, mURCSCode, "C" & mColLooper.ToString, 130, 132, 133))
            Next

            .WriteLine("{\pard\Line\Line\par}")

        End With

        Filewriter.Close()

    End Sub


    Function CompareSums( _
                      ByVal mYear As String, _
                      ByVal mRRICC As String, _
                      ByVal mURCSCode As String, _
                      ByVal mCol As String, _
                      ByVal mStartLine As Integer, _
                      ByVal mEndLine As Integer, _
                      ByVal mSumLine As Integer) As String

        Dim mSum1 As Long, mSum2 As Long
        Dim mColumn As String

        mColumn = ConvertColumn(mCol)

        mSum1 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mColumn, mStartLine, mEndLine)
        mSum2 = Get_Trans_Values_Sum(mYear, mRRICC, mURCSCode, mColumn, mSumLine)

        If mSum1 <> mSum2 Then
            CompareSums = SumsDontMatch(mColumn, mStartLine, mEndLine, mSum1.ToString, mSumLine, mSum2.ToString)
        Else
            CompareSums = SumsMatch(mCol, mStartLine, mEndLine, mSumLine)
        End If

    End Function


    Function SumsDontMatch( _
                      ByVal mCol As String, _
                      ByVal mStartLine As String, _
                      ByVal mEndLine As String, _
                      ByVal mSum1 As String, _
                      ByVal mSumLine As String, _
                      ByVal mSum2 As String) As String

        Dim mThisCol As String = ""

        If IsNumeric(mCol) Then
            mCol = "C" & mCol
        End If

        mThisCol = ColText(mCol)

        SumsDontMatch = "{\pard\tab \b L" & mStartLine & mThisCol & " thru L" & mEndLine & mThisCol & "(" & mSum1 & ") <> L" & mSumLine & mThisCol & "(" & mSum2 & ").\par}"

    End Function

    Function SumsMatch( _
                      ByVal mCol As String, _
                      ByVal mStartLine As String, _
                      ByVal mEndLine As String, _
                      ByVal mSumLine As String) As String

        Dim mThisCol As String = ""

        If IsNumeric(mCol) Then
            mCol = "C" & mCol
        End If

        mThisCol = ColText(mCol)

        SumsMatch = "{\pard\tab L" & mStartLine & mThisCol & " thru L" & mEndLine & mThisCol & "= L" & mSumLine & mThisCol & "\par}"

    End Function


End Class