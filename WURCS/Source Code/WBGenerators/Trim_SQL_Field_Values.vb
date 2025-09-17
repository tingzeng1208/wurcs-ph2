Imports System.Data.SqlClient
Public Class Trim_SQL_Field_Values

    Private Sub Trim_SQL_Field_Values_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
        Dim rst As ADODB.Recordset
        Dim mStrSQL As String

        'Set the form so it centers on the user's screen
        Me.CenterToScreen()

        ' Load the Year combobox from the SQL database
        OpenADOConnection(Global_Variables.Gbl_Controls_Database_Name)

        rst = SetRST()
        mStrSQL = "SELECT wb_year FROM " & Global_Variables.Gbl_WB_Years_TableName

        rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)
        rst.MoveFirst()

        Do While Not rst.EOF
            cmb_URCS_Year.Items.Add(rst.Fields("wb_year").Value)
            rst.MoveNext()
        Loop

        rst.Close()
        rst = Nothing

    End Sub

    Private Sub btn_Return_To_DataPrepMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Return_To_DataPrepMenu.Click
        ' Open the WBGenerators And Utilities Menu Form
        Dim frmNew As New frm_WBGeneratorsAndUtilitiesMenu()
        frmNew.Show()
        ' Close this Menu
        Me.Close()
    End Sub

    Private Sub btn_Execute_Click(sender As System.Object, e As System.EventArgs) Handles btn_Execute.Click

        Dim rst As New ADODB.Recordset
        Dim mStrSQL As String
        Dim mbolWrite As Boolean
        Dim mThisRec As Single, mLooper As Integer

        'Perform error checking for the form controls
        If Me.cmb_URCS_Year.Text = "" Then
            MsgBox("You must select a year to process.", vbOKOnly)
            GoTo EndIt
        End If

        'Get the database_name and table_name for the URCS_Years value from the database
        Get_Table_Name_From_SQL(Me.cmb_URCS_Year.Text, "MASKED")
        Global_Variables.Gbl_Masked_TableName = Global_Variables.gbl_Table_Name

        mbolWrite = MsgBox("The number of records returned is " & CStr(Count_Waybills(cmb_URCS_Year.SelectedItem)) & _
                ". Are you sure you want to update the data?", vbYesNo)

        If mbolWrite = True Then

            OpenADOConnection(Global_Variables.Gbl_Waybill_Database_Name)

            txt_StatusBox.Text = "Getting Data From Server"

            mStrSQL = "SELECT * FROM " & Global_Variables.Gbl_Masked_TableName

            rst.Open(mStrSQL, Global_Variables.gbl_ADOConnection)
            rst.MoveFirst()

            mThisRec = 1

            Do While Not rst.EOF

                If mThisRec Mod 100 = 0 Then
                    txt_StatusBox.Text = "Processing record " & CStr(mThisRec) & " of " & CStr(Count_Waybills(cmb_URCS_Year.SelectedItem)) & "..."
                    Refresh()
                End If

                For mLooper = 0 To rst.Fields.Count - 1
                    Select Case UCase(rst.Fields(mLooper).Name)
                        Case "U_CAR_INIT", "TOFC_SERV_CODE", "U_TC_INIT", "WB_ID", _
                        "JCT1", "JCT2", "JCT3", "JCT4", "JCT5", "JCT6", "JCT7", _
                        "CAR_OWN_MARK", "CAR_LESSEE_MARK", "TYPE_WHEEL", "NO_AXLES", _
                        "CAR_TYP", "MECH", "LIC_ST", "O_SPLC", "T_SPLC", "ORR_ALPHA", "JRR1_ALPHA", _
                        "JRR2_ALPHA", "JRR3_ALPHA", "JRR4_ALPHA", "JRR5_ALPHA", "JRR6_ALPHA", _
                        "TRR_ALPHA", "ERROR_FLG", "CAR_OWN", "TOFC_UNIT_TYPE", "DEREG_FLG", _
                        "O_ST", "JCT1_ST", "JCT2_ST", "JCT3_ST", "JCT4_ST", "JCT5_ST", "JCT6_ST", _
                        "JCT7_ST", "T_ST", "INT_HARM_CODE", "INDUS_CLASS", "INTER_SIC", "DOM_CANADA", _
                        "O_FS_TYPE", "T_FS_TYPE", "O_FS_RATEZIP", "T_FS_RATEZIP", "O_RATE_SPLC", _
                        "T_RATE_SPLC", "O_SWLIMIT_SPLC", "T_SWLIMIT_SPLC", "O_CUSTOMS_FLG", _
                        "T_CUSTOMS_FLG", "O_GRAIN_FLG", "T_GRAIN_FLG", "O_RAMP_CODE", "T_RAMP_CODE", _
                        "O_IM_FLG", "T_IM_FLG", "TRANSBORDER_FLG", "ORR_CNTRY", "JRR1_CNTRY", _
                        "JRR2_CNTRY", "JRR3_CNTRY", "JRR4_CNTRY", "JRR5_CNTRY", "JRR6_CNTRY", _
                        "TRR_CNTRY", "O_CENSUS_REG", "T_CENSUS_REG"
                            If Len(Trim(rst.Fields(mLooper).Value)) <> Len(rst.Fields(mLooper).Value) Then
                                ' We have to trim it
                                mStrSQL = "Update dbo." & Global_Variables.Gbl_Masked_TableName & _
                                    " SET " & rst.Fields(mLooper).Name & " = '" & Trim(rst.Fields(mLooper).Value) & "'" & _
                                    " WHERE serial_no = " & CStr(rst.Fields("serial_no").Value)
                                Global_Variables.gbl_ADOConnection.Execute(mStrSQL)
                            End If
                        Case "CS_54"
                            If Len(rst.Fields(mLooper).Value) < 2 Then
                                mStrSQL = "Update dbo." & Global_Variables.Gbl_Masked_TableName & _
                                    " SET " & rst.Fields(mLooper).Name & " = '0" & rst.Fields(mLooper).Value & "'" & _
                                    " WHERE serial_no = " & CStr(rst.Fields("serial_no").Value)
                                Global_Variables.gbl_ADOConnection.Execute(mStrSQL, )
                            End If
                    End Select

                Next mLooper

                mThisRec = mThisRec + 1
                rst.MoveNext()
            Loop

            rst.Close()
            rst = Nothing

        End If

EndIt:

    End Sub
End Class