Imports System.Data.SqlClient
Public Class frm_AddMakeWholeToXML
    Private Sub frm_AddMakeWholeToXML_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CenterToScreen()

    End Sub

    Private Sub btn_Input_File_Entry_Click(sender As Object, e As EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "DAT Files|*.dat|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Residual_FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub btn_Output_File_Entry_Click(sender As Object, e As EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = False
        fd.Filter = "XML Files|*.xml|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_XML_FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub btn_Execute_Click(sender As Object, e As EventArgs) Handles btn_Execute.Click
        Dim mValues(9, 18) As Decimal
        Dim Residual(18) As Decimal
        Dim mIndustrySwitchingResidual(2) As Decimal
        Dim mStationClericalResidual(2) As Decimal
        Dim mInterchangeSWTResidual(2) As Decimal
        Dim mIISwitchingResidual(2) As Decimal
        Dim mMileageResidual(2) As Decimal
        Dim strInline As String, mWorkStr As String, mWorkFilePathName As String
        Dim mRR_Num As Integer, mValueNum As Integer, mC13 As Decimal
        Dim mRRStr As String

        Dim StreamIn As StreamReader
        Dim FileOut As FileStream = Nothing
        Dim StreamOut As StreamWriter

        If Me.txt_XML_FilePath.Text = "" Then
            MsgBox("You must select a DAT file to read.", vbOKOnly)
            GoTo EndIt
        End If

        If Me.txt_XML_FilePath.Text = "" Then
            MsgBox("You must select an XML file to update.", vbOKOnly)
            GoTo EndIt
        End If

        ' Load the RR_IDs and Short_Names to the arrays from SQL for lookup use
        Gbl_Controls_Database_Name = "URCS_Controls"   'Default value for Control Tables

        gbl_SQLConnection = New SqlConnection
        gbl_SQLConnection.ConnectionString = LoadSQLConnectionStringForDB(Gbl_Controls_Database_Name)
        gbl_SQLConnection.Open()

        'mDataTable = Get_CLASS1RAILLIST_Table()
        'For i = 1 To mDataTable.Rows.Count - 1
        '    mRR_IDs(i) = CInt(mDataTable.Rows(i - 1)("RR_ID").ToString)
        '    mShort_Names(i) = mDataTable.Rows(i - 1)("SHORT_NAME").ToString
        'Next

        ' Open the input file for reading
        StreamIn = New StreamReader(txt_Residual_FilePath.Text)
        ' read/skip the first line as we don't need it for processing
        strInline = StreamIn.ReadLine

        ' Read until EOF
        Do While StreamIn.Peek() >= 0
            strInline = StreamIn.ReadLine
            mWorkStr = strInline
            ' Get the RR_ID
            mRR_Num = Mid(mWorkStr, 1, InStr(mWorkStr, ",") - 1)
            ' Any number > 9 can be ignored, so we'll exit the loop
            If mRR_Num > 9 Then
                Exit Do
            End If
            ' remove the RR_ID from the mworkstr
            mWorkStr = Mid(mWorkStr, InStr(mWorkStr, ",") + 1)
            ' Load the values to the mValues array for this road
            mValueNum = 1
            Do While InStr(mWorkStr, ",") > 0
                mValues(mRR_Num, mValueNum) = CDec(Mid(mWorkStr, 1, InStr(mWorkStr, ",") - 1))
                mWorkStr = Mid(mWorkStr, InStr(mWorkStr, ",") + 1)
                mValueNum = mValueNum + 1
            Loop
            ' Get the last value for this road
            mValues(mRR_Num, mValueNum) = CDec(mWorkStr)
        Loop

        ' Close the input file
        StreamIn.Close()

        ' Copy the XML file as a backup before modification
        mWorkFilePathName = Replace(txt_XML_FilePath.Text, ".xml", "-Original.xml")
        My.Computer.FileSystem.CopyFile(txt_XML_FilePath.Text, mWorkFilePathName)

        ' Open the backup as read-only input
        StreamIn = New StreamReader(mWorkFilePathName)

        ' Open the output filestream for writing
        FileOut = New FileStream(txt_XML_FilePath.Text, FileMode.Truncate)
        StreamOut = New StreamWriter(FileOut, System.Text.Encoding.Unicode)

        ' Read the input file, line by line
        Do While StreamIn.Peek() >= 0
            mWorkStr = StreamIn.ReadLine
            mC13 = 0
            If InStr(mWorkStr, "<Railroad Name=") > 0 Then
                ' We're at a RR and we need to capture it
                mRRStr = Mid(mWorkStr, InStr(mWorkStr, Chr(34)) + 1)
                mRRStr = Mid(mRRStr, 1, InStr(mRRStr, Chr(34)) - 1)
                ' Now we have the Short Name and can get the RR_ID
                mRR_Num = Get_RRID_By_Short_Name(mRRStr)

                'Calculate the values for this road
                If (mValues(mRR_Num, 11) > 0) Then
                    mIndustrySwitchingResidual(2) = mValues(mRR_Num, 1) / mValues(mRR_Num, 11)
                End If
                If (mValues(mRR_Num, 13) > 0) Then
                    mIndustrySwitchingResidual(1) = (mValues(mRR_Num, 6) / mValues(mRR_Num, 13)) + mIndustrySwitchingResidual(2)
                End If
                If (mValues(mRR_Num, 12) > 0) Then
                    mStationClericalResidual(2) = mValues(mRR_Num, 8) / mValues(mRR_Num, 12)
                End If
                If (mValues(mRR_Num, 14) > 0) Then
                    mStationClericalResidual(1) = (mValues(mRR_Num, 7) / mValues(mRR_Num, 14)) + mStationClericalResidual(2)
                End If
                If (mValues(mRR_Num, 15) > 0) Then
                    mInterchangeSWTResidual(2) = mValues(mRR_Num, 2) / mValues(mRR_Num, 15)
                End If
                If (mValues(mRR_Num, 16) > 0) Then
                    mInterchangeSWTResidual(1) = (mValues(mRR_Num, 4) / mValues(mRR_Num, 16)) + mInterchangeSWTResidual(2)
                End If
                If (mValues(mRR_Num, 17) > 0) Then
                    mIISwitchingResidual(2) = mValues(mRR_Num, 3) / mValues(mRR_Num, 17)
                End If
                If (mValues(mRR_Num, 18) > 0) Then
                    mIISwitchingResidual(1) = (mValues(mRR_Num, 5) / mValues(mRR_Num, 18)) + mIISwitchingResidual(2)
                End If
                If (mValues(mRR_Num, 17) > 0) Then
                    mMileageResidual(2) = mValues(mRR_Num, 9) / mValues(mRR_Num, 17)
                End If
                If (mValues(mRR_Num, 18) > 0) Then
                    mMileageResidual(1) = (mValues(mRR_Num, 10) / mValues(mRR_Num, 18)) + mMileageResidual(2)
                End If
                ' Always write the mWorkstr value to the file
                StreamOut.WriteLine(mWorkStr)
            Else
                ' Substitute the lines that we need to update in the file
                If InStr(mWorkStr, "<E2P3L301") > 0 Then
                    mWorkStr = "    <E2P3L301 C1=" & Chr(34) & mIndustrySwitchingResidual(1).ToString & Chr(34) & " "
                    mWorkStr = mWorkStr & "C2=" & Chr(34) & mIndustrySwitchingResidual(2).ToString & Chr(34) & " />"
                End If
                If InStr(mWorkStr, "<E2P3L302") > 0 Then
                    mWorkStr = "    <E2P3L302 C1=" & Chr(34) & mStationClericalResidual(1).ToString & Chr(34) & " "
                    mWorkStr = mWorkStr & "C2=" & Chr(34) & mStationClericalResidual(2).ToString & Chr(34) & " />"
                End If
                If InStr(mWorkStr, "<E2P3L303") > 0 Then
                    mWorkStr = "    <E2P3L303 C1=" & Chr(34) & mInterchangeSWTResidual(1).ToString & Chr(34) & " "
                    mWorkStr = mWorkStr & "C2=" & Chr(34) & mInterchangeSWTResidual(2).ToString & Chr(34) & " />"
                End If
                If InStr(mWorkStr, "<E2P3L304") > 0 Then
                    mWorkStr = "    <E2P3L304 C1=" & Chr(34) & mIISwitchingResidual(1).ToString & Chr(34) & " "
                    mWorkStr = mWorkStr & "C2=" & Chr(34) & mIISwitchingResidual(2).ToString & Chr(34) & " />"
                End If
                If InStr(mWorkStr, "<E2P3L305") > 0 Then
                    mWorkStr = "    <E2P3L305 C1=" & Chr(34) & mMileageResidual(1).ToString & Chr(34) & " "
                    mWorkStr = mWorkStr & "C2=" & Chr(34) & mMileageResidual(2).ToString & Chr(34) & " />"
                End If
                ' Always write the mWorkstr value to the file
                StreamOut.WriteLine(mWorkStr)
            End If
        Loop

        StreamIn.Close()
        StreamOut.Close()
        FileOut.Close()

        txt_Status.Text = "Done!"
        Refresh()
EndIt:

    End Sub

    Private Sub btn_Return_To_PostProcessingMenu_Click(sender As Object, e As EventArgs) Handles btn_Return_To_PostProcessingMenu.Click
        ' Open the post processing Menu Form
        Dim frmNew As New frm_Post_Processing_Menu
        frmNew.Show()
        ' Close this Form
        Me.Close()
    End Sub
End Class