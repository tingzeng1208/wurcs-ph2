Imports System.Text
Public Class STCC_Code_Table_Loader


    Private Sub btn_Return_To_DataPrepMenu_Click(sender As System.Object, e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the Data Prep Menu Form
        Dim frmNew As New frm_URCS_Waybill_Data_Prep()
        frmNew.Show()
        ' Close this form       
        Me.Close()
    End Sub

    Private Sub btn_Input_File_Entry_Click(sender As System.Object, e As System.EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "Text Files|*.txt|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub STCC_Code_Table_Loader_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Center the form on the user's screen
        Me.CenterToScreen()
    End Sub

    Private Sub btn_Execute_Click(sender As System.Object, e As System.EventArgs) Handles btn_Execute.Click

        Dim rst As New ADODB.Recordset

        Dim mStrSQL As StringBuilder
        Dim mSr As StreamReader

        Dim bolWrite As Integer
        Dim mInString As String
        Dim mThisRec As Integer
        Dim mThisSTCC As String, mLastSTCC As String
        Dim mWorkStr1 As String, mWorkStr2 As String

        If txt_Input_FilePath.TextLength = 0 Then
            MsgBox("You must select an Input File value.", MsgBoxStyle.OkOnly)
            GoTo EndIt
        End If

        ' ask the user if he wants to load x records for this RRID
        bolWrite = MsgBox("Are you sure you want to Write this data?", vbYesNo)

        If bolWrite = vbYes Then

            Gbl_Controls_Database_Name = My.Settings.Controls_DB

            ' Open/Check the SQL connection
            OpenADOConnection(Gbl_Controls_Database_Name)

            Gbl_STCC_Codes_TableName = Get_Table_Name_From_SQL("1", "EXTENDED_STCC")

            'If the table does not exist, create it
            If VerifyTableExist(Gbl_Controls_Database_Name, Gbl_STCC_Codes_TableName) = False Then
                Create_STCC_Table()
            End If

            'Open the text file for reading
            mSr = New StreamReader(txt_Input_FilePath.Text)
            mThisRec = 0
            mLastSTCC = "00"

            Do While Not mSr.EndOfStream
                mThisRec = mThisRec + 1
                If mThisRec Mod 100 = 0 Then
                    Me.txt_StatusBox.Text = "Working - Loaded " & CStr(mThisRec) & " records..."
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                End If

                mInString = mSr.ReadLine()

                'In the data file, the HazMat Response code is the STCC Code in the 3949 length record
                mThisSTCC = mInString.Substring(0, 7)

                'Check to see if a record exists for this STCC
                mStrSQL = New StringBuilder
                mStrSQL.Append("SELECT * FROM ")
                mStrSQL.Append(Gbl_STCC_Codes_TableName)
                mStrSQL.Append(" WHERE STCC = '" & mThisSTCC & "'")

                ' Check/Open the SQL connection
                OpenADOConnection(Gbl_Controls_Database_Name)

                rst = SetRST()
                rst.Open(mStrSQL.ToString, Global_Variables.gbl_ADOConnection)

                If rst.RecordCount = 0 Then
                    Select Case mInString.Length
                        Case Is > 3500   'Hazmat Listing
                            ' Build the append sql statement for the HazMat record
                            mStrSQL = Build_Insert_HazMat_STCC_Rec(mInString, Gbl_STCC_Codes_TableName)

                        Case Else   ' Non-HazMat record
                            ' Build the append sql statement for the non-HazMat record (2825 record length)
                            mStrSQL = Build_Insert_Non_HazMat_STCC_Rec(mInString, Gbl_STCC_Codes_TableName)
                    End Select

                    ' Check/Open the SQL connection
                    OpenADOConnection(Gbl_Controls_Database_Name)

                    ' Execute the statement
                    Global_Variables.gbl_ADOConnection.Execute(mStrSQL.ToString)

                Else
                    ' A record exists.  Is it later than the record we're reading?
                    ' Get the date from the record
                    Select Case mInString.Length
                        Case Is > 3500   'Hazmat Listing
                            mStrSQL = Build_Update_HazMat_STCC_Rec(mInString, Gbl_STCC_Codes_TableName)
                        Case Else
                            If InStr(txt_Input_FilePath.Text, "STCC Headers") > 0 Then
                                mStrSQL = New StringBuilder
                                mStrSQL.Append("UPDATE " & Gbl_STCC_Codes_TableName & " SET ")
                                ' The header file has multiple lines in some cases
                                If mLastSTCC = mThisSTCC Then
                                    mWorkStr1 = rst.Fields("alpha_desc").Value
                                    mWorkStr2 = Trim(mInString.Substring(282, 25))
                                    ' we have a continuance of the description - ignore more than 5 continuances
                                    If InStr(rst.Fields("alpha_desc").Value, "-") > 0 Then
                                        mWorkStr1 = mWorkStr1.Substring(0, Len(mWorkStr1) - 1)
                                    Else
                                        mWorkStr2 = " " & mWorkStr2
                                    End If
                                    mStrSQL.Append("Alpha_Desc = '" & mWorkStr1 & mWorkStr2 & "'")
                                Else
                                    mStrSQL.Append("Alpha_Desc = '" & Trim(mInString.Substring(282, 25)) & "'")
                                End If
                                mStrSQL.Append(" WHERE STCC = '" & Trim(mInString.Substring(0, 7)) & "'")
                            Else
                                mStrSQL = Build_Update_Non_HazMat_STCC_Rec(mInString, Gbl_STCC_Codes_TableName)
                            End If
                    End Select
                End If

                ' Check/Open the SQL connection
                OpenADOConnection(Gbl_Controls_Database_Name)

                ' Execute the statement
                Global_Variables.gbl_ADOConnection.Execute(mStrSQL.ToString)

                rst.Close()
                rst = Nothing

                mLastSTCC = mThisSTCC

            Loop

            Me.txt_StatusBox.Text = "Done!"
            Me.Refresh()

            'Clean up
            mSr.Close()
            mSr = Nothing

        End If

EndIt:

    End Sub
End Class