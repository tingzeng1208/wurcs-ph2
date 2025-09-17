Public Class Legacy_Costing

    Dim NUMEVALS As Integer = 974
    Dim NUMFVALS As Integer = 321

    Private Sub btn_Return_To_MainMenu_Click(sender As Object, e As EventArgs) Handles btn_Return_To_MainMenu.Click
        ' Open the Main Menu Form
        Dim frmNew As New frm_MainMenu
        frmNew.Show()
        ' Close the this Form
        Close()
    End Sub

    Private Sub Legacy_Costing_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim mDataTable As DataTable

        CenterToScreen()

        ' Load the Year combobox from the SQL database
        mDataTable = Get_URCS_Years_Table()

        For mLooper = 0 To mDataTable.Rows.Count - 1
            cmb_URCS_Year_Combobox.Items.Add(mDataTable.Rows(mLooper)("urcs_year").ToString)
        Next

        mDataTable = Nothing

    End Sub

    Private Sub btn_Execute_Click(sender As Object, e As EventArgs) Handles btn_Execute.Click
        Dim mRoads_Table As DataTable
        Dim mWaybills As DataTable
        Dim mWAYRRR As DataRow
        Dim mRoadRegion As Integer
        Dim mStartTime As DateTime
        Dim mThisWaybill As Integer, mTotalWaybillRecords As Integer
        Dim mIntermodal As Boolean
        Dim mCarOwnership As Integer, mFreightCar As Integer
        Dim mSerial_no As Integer
        Dim mSampled_Cars As Integer, mExpandedCars As Integer


        mStartTime = Now
        txt_StatusBox.Text = "Getting Waybills and Railroad data from SQL..."
        Refresh()

        ' Get the Waybills for this year
        mWaybills = GetWaybillData(cmb_URCS_Year_Combobox.Text, True) 'Since we need the rate_flg, we'll get the info unmasked
        mTotalWaybillRecords = mWaybills.Rows.Count

        ' Get the Roads and Regions info, but exclude the national data
        mRoads_Table = Get_Railroads_For_Costing()

        ' Loop thru the Waybills
        For mThisWaybill = 0 To mWaybills.Rows.Count - 1

            ' Let the user know where we are
            If (mThisWaybill Mod 100 = 0 And mThisWaybill > 0) Then
                txt_StatusBox.Text = "Processed " & mThisWaybill.ToString & " of " & mTotalWaybillRecords.ToString & "..."
                Refresh()
                Application.DoEvents()
            End If

            ' Get the serial No
            mSerial_no = mWaybills.Rows(mThisWaybill)("Serial_No")

            ' Get the number of cars sampled and expanded
            mSampled_Cars = mWaybills.Rows(mThisWaybill)("U_Cars")
            mExpandedCars = mWaybills.Rows(mThisWaybill)("Cars")

            'Determine who owns this car - Railroad owned is 0, Privately owned is 1
            'We set this to Railroad Owned = 1, Privately owned is 2
            mCarOwnership = RR_ftcar_Get_Owner(mWaybills(mThisWaybill)("Car_Own_Mark")) + 1

            ' Determine what type of car we have
            mFreightCar = RR_ftcar_Get_CarType(mWaybills(mThisWaybill)("STB_Car_Typ"))

            'Check freight cars for intermodal moves - some unknown car types may be intermodal
            Select Case mFreightCar
                Case 11, 14, 17, 18
                    mIntermodal = RR_ftcar_Get_TOFC(mWaybills(mThisWaybill)("TOFC_Serv_Code"), mCarOwnership, mWaybills(mThisWaybill)("All_Rail_Code"))
                    ' Find the info on the origin railroad
                    mWAYRRR = mRoads_Table.Rows.Find("AARID = " & mWaybills(mThisWaybill)("ORR").ToString)
                    ' Set trailer return factor on origin railroad
                    mRoadRegion = mWAYRRR.Field(Of Integer)("Region_ID")


            End Select


        Next mThisWaybill

    End Sub

    Private Sub btn_Input_File_Entry_Click(sender As Object, e As EventArgs) Handles btn_Input_File_Entry.Click
        Dim fd As New OpenFileDialog

        fd.Multiselect = False
        fd.CheckFileExists = True
        fd.Filter = "XML Files|*.xml|All Files|*.*"

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txt_InputXML_FilePath.Text = fd.FileName
        End If
    End Sub

    Function RR_ftcar_Get_CarType(ByRef mCartype As Integer) As Integer
        RR_ftcar_Get_CarType = 18  ' Unknown by default

        Select Case mCartype
            Case 36
                RR_ftcar_Get_CarType = 1 'Unequipped box car
            Case 37
                RR_ftcar_Get_CarType = 2 '50 ft box car
            Case 38
                RR_ftcar_Get_CarType = 3 'Equipped box car
            Case 39
                RR_ftcar_Get_CarType = 4 'Unequipped genaral service gondola
            Case 40
                RR_ftcar_Get_CarType = 5 'Equipped general service gondola
            Case 41
                RR_ftcar_Get_CarType = 6 'Covered hopper
            Case 42
                RR_ftcar_Get_CarType = 7 'General service covered hopper
            Case 43
                RR_ftcar_Get_CarType = 8 'Open, special service hopper
            Case 44
                RR_ftcar_Get_CarType = 9 'Mechanical refrigerator
            Case 45
                RR_ftcar_Get_CarType = 10 'Non-mechanical refrigerator
            Case 46
                RR_ftcar_Get_CarType = 11 'TOFC Flat
            Case 47
                RR_ftcar_Get_CarType = 12 'Multi level flat
            Case 48
                RR_ftcar_Get_CarType = 13 'General service flat
            Case 49
                RR_ftcar_Get_CarType = 14 'Other flat
            Case 50
                RR_ftcar_Get_CarType = 15 'Tank, less than 22,000 gallons
            Case 51
                RR_ftcar_Get_CarType = 16 'Tank, more than 22,000 gallons
            Case 52
                RR_ftcar_Get_CarType = 17 'All other freight cars
            Case 54
                RR_ftcar_Get_CarType = 17 'Average car-used To be cabooses
        End Select
    End Function

    Function RR_ftcar_Get_Owner(ByRef mCarInitial As String) As Integer
        Dim mWorkstr As String

        RR_ftcar_Get_Owner = 0 'Default Is railroad owned car
        mWorkstr = Trim(mCarInitial)

        'Now using the private car flag to determine if the car was private
        'This code was the historical means for determining private cars
        Select Case mWorkstr
            Case "ABOX", "RBOX", "CSX", "CSXT", "GONX"
                RR_ftcar_Get_Owner = 1
            Case Else
                'Test the last character in initial for X - indicates a private freight
                If Mid(mCarInitial, Len(mCarInitial)) = "X" Then
                    RR_ftcar_Get_Owner = 1
                Else
                    RR_ftcar_Get_Owner = 0
                End If
        End Select

Endit:
    End Function

    Function RR_ftcar_Get_TOFC(ByRef mTOFC_Serv_Code As String, ByRef mPrivateCar As Integer, ByRef mAll_Rail_Code As Integer) As Single

        RR_ftcar_Get_TOFC = 0

        If mAll_Rail_Code = 1 Then
            GoTo Endit  ' Not a TOFC move
        End If

        Select Case mTOFC_Serv_Code
            Case "A"
                RR_ftcar_Get_TOFC = 1.0 ' Actual Intermodal Service Code 15
            Case "B"
                RR_ftcar_Get_TOFC = 2.0 ' 20
            Case "C"
                RR_ftcar_Get_TOFC = 2.25 ' 22
            Case "D"
                RR_ftcar_Get_TOFC = 5.1 ' 25
            Case "E"
                RR_ftcar_Get_TOFC = 5.2 ' 27
            Case "F"
                RR_ftcar_Get_TOFC = 5.3 ' 40
            Case "G"
                RR_ftcar_Get_TOFC = 5.4 ' 42
            Case "H"
                RR_ftcar_Get_TOFC = 1.0 ' 45
            Case "I"
                RR_ftcar_Get_TOFC = 5.5 ' 47
            Case "K"
                RR_ftcar_Get_TOFC = 5.3 ' 60
            Case "L"
                RR_ftcar_Get_TOFC = 5.4 ' 62
            Case "M"
                RR_ftcar_Get_TOFC = 1.0 ' 65
            Case "N"
                RR_ftcar_Get_TOFC = 5.5 ' 67
            Case "O"
                RR_ftcar_Get_TOFC = 5.3 ' 80
            Case "P"
                RR_ftcar_Get_TOFC = 5.4 ' 82
            Case "Q"
                RR_ftcar_Get_TOFC = 1.0 ' 85
            Case "R"
                RR_ftcar_Get_TOFC = 5.5 ' 87
            Case "X"
                RR_ftcar_Get_TOFC = 1.0 ' Unknown Plan
        End Select

        If mAll_Rail_Code = 2 Then
            RR_ftcar_Get_TOFC = 4.0 ' Road Railer movements
        End If

Endit:



    End Function

End Class