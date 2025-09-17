Imports System.Data.SqlClient
Imports System.Text
Public Class QuarterlyMaskingUtil

    Private Sub btn_Return_To_Menu_Click(sender As System.Object, e As System.EventArgs) Handles btn_Return_To_Menu.Click
        Dim frmNew As New frm_WBGeneratorsAndUtilitiesMenu
        frmNew.Show()
        Me.Close()
    End Sub

    Private Sub btn_Text_File_Entry_Click(sender As System.Object, e As System.EventArgs) Handles btn_Text_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "TXT Files|*.txt|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Input_Text_FilePath.Text = fd.FileName
        End If
    End Sub

    Private Sub btn_Output_File_Entry_Click(sender As System.Object, e As System.EventArgs) Handles btn_Output_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = False
        fd.Filter = "TXT Files|*.txt|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_Output_FilePath.Text = fd.FileName
        End If

    End Sub

    Private Sub btn_Execute_Click(sender As System.Object, e As System.EventArgs) Handles btn_Execute.Click

        Dim mSQLCmd As SqlCommand
        Dim fileWriter As StreamWriter
        Dim stringReader As String
        Dim mOutString As StringBuilder
        Dim RecNum As Integer, TotalRecs As Integer
        Dim mWriteIt As Boolean
        Dim mLooper As Integer
        Dim mRateFlg As String
        Dim mSTCC_W49 As String
        Dim mU_Rev As Integer
        Dim mDate As String
        Dim mAcctYear As String
        Dim mAcctMonth As String
        Dim mFSAC As String
        Dim mWbNum As String
        Dim mUCarNum As String
        Dim mTCNum As String
        Dim mCar_Init As String
        Dim mYear As String
        Dim mThisStr As String
        Dim mReport_RR As String

        TotalRecs = 0
        RecNum = 0
        mLooper = 0

        txt_StatusBox.Text = "Processing..."
        Refresh()

        LoadArrayData()

        Gbl_EIA_STCCs_TableName = Get_Table_Name_From_SQL("1", "EIA_STCCS")
        Gbl_Ordnance_STCCs = Get_Table_Name_From_SQL("1", "ORDNANCE_STCCS")

        TotalRecs = File.ReadAllLines(txt_Input_Text_FilePath.Text).Length

        fileWriter = New StreamWriter(txt_Output_FilePath.Text, False, Encoding.ASCII)
        RecNum = 0

        Using reader As StreamReader = New StreamReader(txt_Input_Text_FilePath.Text)
            stringReader = reader.ReadLine

            Do While (Not stringReader Is Nothing)
                mOutString = New StringBuilder
                RecNum = RecNum + 1
                mWriteIt = False

                ' Let the user know where we are
                If RecNum Mod 1000 = 0 Then
                    txt_StatusBox.Text = "Processing Line " & RecNum.ToString & " of " & TotalRecs.ToString
                    Application.DoEvents()
                    Refresh()
                End If

                mReport_RR = Mid(stringReader, 146, 3)
                mRateFlg = Mid(stringReader, 120, 1)
                mSTCC_W49 = Mid(stringReader, 54, 7)
                mU_Rev = Mid(stringReader, 79, 9)
                mDate = Mid(stringReader, 13, 2) & "/" & Mid(stringReader, 15, 2) & "/" & Mid(stringReader, 17, 2)
                mYear = Mid(stringReader, 17, 2)
                mAcctYear = Mid(stringReader, 21, 2) 'Dropped "20" preface 6/11/2021 MRS
                mAcctMonth = Mid(stringReader, 19, 2)
                mFSAC = Mid(stringReader, 229, 5)
                mWbNum = Mid(stringReader, 7, 6)
                mUCarNum = Mid(stringReader, 31, 6)
                mTCNum = Mid(stringReader, 48, 6)
                mCar_Init = Mid(stringReader, 27, 1)

                If Trim(mTCNum) = "" Then
                    mTCNum = 0
                End If

                'adjust the Acct_Year as needed
                'If Val(mAcctYear >= 80) Then
                '    mAcctYear = 1900 + Val(mAcctYear).ToString
                'Else
                '    mAcctYear = 2000 + Val(mAcctYear).ToString
                'End If

                'Adjust mAcctYear as necessary
                If Val(mAcctYear) > 80 Then
                    mAcctYear = Val(Val(mAcctYear) + 1900).ToString
                Else
                    mAcctYear = Val(Val(mAcctYear) + 2000).ToString
                End If

                ' If this is an EIA run, only write the records with STCCs they've requested
                If rdo_Unmask_Some_49s_For_EIA.Checked = True Then
                    mSQLCmd = New SqlCommand
                    mSQLCmd.CommandType = CommandType.Text
                    OpenSQLConnection(My.Settings.Waybills_DB)
                    mSQLCmd.Connection = gbl_SQLConnection
                    mSQLCmd.CommandText = "select COUNT(*) from " & Gbl_EIA_STCCs_TableName & " WHERE STCC = " & mSTCC_W49

                    If mSQLCmd.ExecuteScalar > 0 Then
                        mOutString = New StringBuilder
                        mOutString.Append(Mid(stringReader, 1, 78))
                        If mRateFlg = "1" Then
                            ' This record is masked
                            Select Case mReport_RR

                                Case 105, 482    'CP
                                    mU_Rev = UnmaskCPValue(0, mSTCC_W49, mU_Rev)

                                Case 131    'CNW
                                'Do nothing - No way to unmask CN data without State value

                                Case 400    'KCS
                                'Do nothing - they rely upon our masking which does not apply to monthly/quarterly

                                Case 555    'NS
                                    mU_Rev = UnmaskNSValue(0, mDate, mAcctYear, mAcctMonth, mFSAC, mWbNum, mUCarNum, mTCNum, mU_Rev)

                                Case 712    'CSX
                                    mU_Rev = UnmaskCSXValue(0, mSTCC_W49, mWbNum, mAcctMonth, mAcctYear, mU_Rev)

                                Case 777    'BNSF
                                    mU_Rev = UnmaskBNSFValue(0, mCar_Init, mAcctMonth, mU_Rev)

                                Case 802  'UP
                                    mU_Rev = UnmaskUPValue(0, Val(mSTCC_W49), Year(mDate), mWbNum, mU_Rev)

                            End Select

                            mThisStr = mU_Rev.ToString
                            Do While Len(mThisStr) < 9
                                mThisStr = "0" & mThisStr
                            Loop

                            mOutString.Append(mThisStr)
                            mOutString.Append(Mid(stringReader, 88))
                            mWriteIt = True
                        Else
                            ' The record is unmasked, but in the EIA table
                            mOutString = New StringBuilder
                            mOutString.Append(stringReader)
                            mWriteIt = True
                        End If
                    Else
                        ' The record is not in the EIA table
                        mWriteIt = False
                    End If
                End If


                mSQLCmd = New SqlCommand
                mSQLCmd.CommandType = CommandType.Text
                OpenSQLConnection(My.Settings.Waybills_DB)
                mSQLCmd.Connection = gbl_SQLConnection
                mSQLCmd.CommandText = "select COUNT(*) from " & Gbl_Ordnance_STCCs & " WHERE STCC = " & mSTCC_W49

                If mSQLCmd.ExecuteScalar <> 0 Then
                    ' we mask ordnance STCCs
                    Select Case Mid(mSTCC_W49, 1, 2)
                        Case "19"
                            mSTCC_W49 = "1900000"
                        Case "49"
                            mSTCC_W49 = "4900000"
                    End Select
                End If

                ' If masked data is requested, set the rate flags to 0 and mask KCS records 
                If rdo_Masked.Checked = True Then
                    mThisStr = ""
                    mOutString = New StringBuilder
                    mOutString.Append(Mid(stringReader, 1, 53))
                    mOutString.Append(mSTCC_W49)
                    mOutString.Append(Mid(stringReader, 61, 18))
                    mThisStr = mU_Rev.ToString
                    'If CInt(mReport_RR) = 400 Then   'Mask KCS revenues
                    '    mThisStr = MaskGenericValue(mSTCC_W49, mDate, mYear, mU_Rev)
                    'End If
                    Select Case CInt(mReport_RR)
                        Case 105, 482, 131, 555, 712, 777, 802
                            'do nothing for these class 1 roads
                        Case Else
                            If mRateFlg = "1" Then
                                mThisStr = MaskGenericValue(mSTCC_W49, mDate, mAcctYear, mU_Rev)
                            End If
                    End Select
                    Do While Len(mThisStr) < 9
                        mThisStr = "0" & mThisStr
                    Loop
                    mOutString.Append(mThisStr)
                    mOutString.Append(Mid(stringReader, 88, 32))
                    If mRateFlg = "1" Then
                        mOutString.Append("0")
                    Else
                        mOutString.Append(Mid(stringReader, 120, 1))
                    End If
                    mOutString.Append(Mid(stringReader, 121))
                    mWriteIt = True
                End If

                ' if unmasked data is requested, unmask the revenue if the rate flag is 1
                If rdo_Unmasked.Checked = True Then

                    mOutString = New StringBuilder
                    mOutString.Append(Mid(stringReader, 1, 53))
                    mOutString.Append(mSTCC_W49)
                    mOutString.Append(Mid(stringReader, 61, 18))
                    If mRateFlg = "1" Then
                        ' This record is masked
                        Select Case mReport_RR

                            Case 105, 482    'CP
                                mU_Rev = UnmaskCPValue(0, mSTCC_W49, mU_Rev)

                            Case 400    'KCS
                                'Do nothing - they rely upon our masking which does not apply to monthly/quarterly

                            Case 555    'NS
                                mU_Rev = UnmaskNSValue(0, mDate, mAcctYear, mAcctMonth, mFSAC, mWbNum, mUCarNum, mTCNum, mU_Rev)

                            Case 712    'CSX
                                mU_Rev = UnmaskCSXValue(0, mSTCC_W49, mWbNum, mAcctMonth, mAcctYear, mU_Rev)

                            Case 777    'BNSF
                                mU_Rev = UnmaskBNSFValue(0, mCar_Init, mAcctMonth, mU_Rev)

                            Case 802  'UP
                                mU_Rev = UnmaskUPValue(0, Val(mSTCC_W49), Year(mDate), mWbNum, mU_Rev)

                        End Select

                    End If

                    mThisStr = mU_Rev.ToString
                    Do While Len(mThisStr) < 9
                        mThisStr = "0" & mThisStr
                    Loop

                    mOutString.Append(mThisStr)
                    mOutString.Append(Mid(stringReader, 88))
                    mWriteIt = True
                End If

                If mWriteIt = True Then
                    fileWriter.WriteLine(mOutString.ToString)
                End If

                stringReader = reader.ReadLine
            Loop
        End Using

        fileWriter.Flush()
        fileWriter.Close()
        fileWriter.Dispose()

        txt_StatusBox.Text = "Done!"
        Refresh()

    End Sub

    Private Sub QuarterlyMaskingUtil_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        Me.Refresh()
    End Sub

    Private Sub rdo_Masked_CheckedChanged(sender As Object, e As EventArgs) Handles rdo_Masked.CheckedChanged
        txt_Output_FilePath.Text = txt_Input_Text_FilePath.Text
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, ".", "_Masked.")
    End Sub

    Private Sub rdo_Unmasked_CheckedChanged(sender As Object, e As EventArgs) Handles rdo_Unmasked.CheckedChanged
        txt_Output_FilePath.Text = txt_Input_Text_FilePath.Text
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, ".", "_Unmasked.")
    End Sub

    Private Sub rdo_Unmask_Some_49s_For_EIA_CheckedChanged(sender As Object, e As EventArgs) Handles rdo_Unmask_Some_49s_For_EIA.CheckedChanged
        txt_Output_FilePath.Text = txt_Input_Text_FilePath.Text
        txt_Output_FilePath.Text = Replace(txt_Output_FilePath.Text, ".", "_Unmasked_EIA.")
    End Sub
End Class