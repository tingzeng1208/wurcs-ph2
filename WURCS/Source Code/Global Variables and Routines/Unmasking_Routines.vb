Module Unmasking_Routines

    Function UnmaskGenericValue(
         ByVal mSerial_No As String,
         ByVal mSTCC_W49 As String,
         ByVal mWB_Date As Date,
         ByVal mAcctYear As Integer,
         ByVal mValue As Object) As Object

        Dim mCompare, mCol, mRow As Integer

        If mSerial_No = "" Then
            MsgBox("Error - No serial number value passed to generic unmasking function.",
                   vbOKOnly, "ERROR")
        End If

        If mSTCC_W49 = "" Then
            MsgBox("Error - No STCC_W49 value passed to generic unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If Not IsDate(mWB_Date) Then
            MsgBox("Error - Invalid date value passed to generic unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mAcctYear = 0 Then
            MsgBox("Error - Invalid Account Year value passed to generic unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        ' Get the row value
        Select Case mWB_Date.Day
            Case 1, 11, 21, 31
                mRow = 1
            Case 2, 12, 22
                mRow = 2
            Case 3, 13, 23
                mRow = 3
            Case 4, 14, 24
                mRow = 4
            Case 5, 15, 25
                mRow = 5
            Case 6, 16, 26
                mRow = 6
            Case 7, 17, 27
                mRow = 7
            Case 8, 18, 28
                mRow = 8
            Case 9, 19, 29
                mRow = 9
            Case 10, 20, 30
                mRow = 10
        End Select

        ' Get the column value from the right 2 characters from the STCC code
        mCompare = Val(Right(mSTCC_W49, 2))

        If mCompare >= 0 And mCompare <= 9 Then mCol = 1
        If mCompare >= 10 And mCompare <= 19 Then mCol = 2
        If mCompare >= 20 And mCompare <= 29 Then mCol = 3
        If mCompare >= 30 And mCompare <= 39 Then mCol = 4
        If mCompare >= 40 And mCompare <= 49 Then mCol = 5
        If mCompare >= 50 And mCompare <= 59 Then mCol = 6
        If mCompare >= 60 And mCompare <= 69 Then mCol = 7
        If mCompare >= 70 And mCompare <= 79 Then mCol = 8
        If mCompare >= 80 And mCompare <= 89 Then mCol = 9
        If mCompare >= 90 And mCompare <= 99 Then mCol = 10

        'Calculate based on whether the year is odd or even
        If mAcctYear Mod 2 > 0 Then
            UnmaskGenericValue = mValue / RR_Odd_Factor(mRow, mCol)
        Else
            UnmaskGenericValue = mValue / RR_Even_Factor(mRow, mCol)
        End If

        UnmaskGenericValue = Math.Round(UnmaskGenericValue, 0, MidpointRounding.AwayFromZero)

    End Function

    Function UnmaskCPValue(
        ByVal mserial_no As String,
        ByVal mSTCC_W49 As Long,
         ByVal mVal As Object) As Object

        Dim mCompare As Integer

        If mSTCC_W49.ToString = "" Then
            MsgBox("Error - No STCC_W49 value passed to CP unmasking function for serial Number " & mserial_no,
                   vbOKOnly, "ERROR")
        End If

        mCompare = CInt(mSTCC_W49 / 100000)
        If mCompare <= 19 Then
            UnmaskCPValue = CDec(FormatNumber(mVal / 1.2, 0))
        Else
            UnmaskCPValue = CDec(FormatNumber(mVal / 1.25, 0))
        End If

    End Function

    Function UnmaskCNWValue(
         ByVal mSerial_No As String,
         ByVal mState As String,
         ByVal mCars As Long,
         ByVal mVal As Object) As Object

        Dim mPos As Integer

        If mSerial_No = "" Then
            MsgBox("Error - No serial_number value passed to CN unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mState = "" Then
            MsgBox("Error - No State value passed to CN unmasking function for serial Number " & mSerial_No,
                       vbOKOnly, "ERROR")
        End If

        If mCars = "" Then
            MsgBox("Error - No Cars value passed to CN unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mVal = "" Then
            MsgBox("Error - No masked value passed to CN unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        mPos = Array1DFindFirst(CNWState, mState, 9)

        Select Case mCars
            Case Is <= 50
                UnmaskCNWValue = CDec(FormatNumber(mVal / CNWMult(mPos), 0))
            Case Else
                UnmaskCNWValue = CDec(FormatNumber(mVal / CNWUnit(mPos), 0))
        End Select

    End Function

    Function UnmaskConrailValue(
         ByVal mSerial_No As String,
         ByVal mSTCC_W49 As String,
         ByVal mJF As Integer,
         ByVal mVal As Object) As Object

        Dim mSTCC_Compare As Integer
        Dim mSTCC_Val As Decimal
        Dim mIndex As Integer

        If mSerial_No = "" Then
            MsgBox("Error - No Serial Number value passed to Conrail unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mSTCC_W49 = "" Then
            MsgBox("Error - No STCC49 value passed to Conrail unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mJF = "" Then
            MsgBox("Error - No JF value passed to Conrail unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        mSTCC_Val = Val(mSTCC_W49)

        mSTCC_Compare = Val(Left(Str(mSTCC_Val / 10000), 4))

        Select Case mJF
            Case 0    'Local Move
                mIndex = Array1DFindFirst(ConrailLocalSTCC, mSTCC_Compare, 29)
                UnmaskConrailValue = CDec(FormatNumber(mVal / (1 + ConrailLocalRate(mIndex) / 100), 0))

            Case Else    'Interline Move
                mIndex = Array1DFindFirst(ConrailInterSTCC, mSTCC_Compare, 28)
                UnmaskConrailValue = CDec(FormatNumber(mVal / (1 + ConrailInterRate(mIndex) / 100), 0))
        End Select

    End Function

    Function UnmaskNSValue(
         ByVal mSerial_No As String,
         ByVal mWbDate As Date,
         ByVal mAcctYear As Integer,
         ByVal mAcctMonth As Integer,
         ByVal mFSAC As Long,
         ByVal mWB_num As Long,
         ByVal mCar_num As Long,
         ByVal mContainers As Long,
         ByVal mVal As Object) As Object

        Dim mCarfactor As Object, mDayfactor As Object, mMaskfactor As Object
        Dim mCarsOrContainers As Long, mWB As Long

        If mSerial_No = "" Then
            MsgBox("Error - No Serial Number value passed to NS unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If Not IsDate(mWbDate) Then
            MsgBox("Error - No WBDate value passed to NS unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mAcctYear = 0 Then
            MsgBox("Error - No Account Year value passed to NS unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        'If mSerial_No = 0 Then
        '    MsgBox("Error - No Account month value passed to NS unmasking function for serial Number " & mSerial_No,
        '           vbOKOnly, "ERROR")
        'End If

        If mFSAC = 0 Then
            MsgBox("Error - No FSAC value passed to NS unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mWB_num = 0 Then
            MsgBox("Error - No Waybill Number value passed to NS unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mCar_num = 0 Then
            MsgBox("Error - No Car Number value passed to NS unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        'If mContainers = 0 Then
        '    MsgBox("Error - No Containers Number value passed to NS unmasking function for serial Number " & mSerial_No,
        '           vbOKOnly, "ERROR")
        'End If

        If mVal = 0 Then
            MsgBox("Error - No masked value passed to NS unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If (mAcctYear < 2003) Or ((mAcctYear = 2003) And (mAcctMonth < 10)) Then
            UnmaskNSValue = CDec(mVal / 1.3)
            If mAcctYear >= 1999 Then
                If ((mAcctYear = 1999) And (mAcctMonth >= 6)) Or (mAcctYear > 1999) Then
                    If (mFSAC >= 67000) And (mFSAC <= 75999) Then
                        UnmaskNSValue = CDec(FormatNumber(mVal / 1.35, 0))
                    End If
                End If
            End If
        Else
            If mCar_num = 0 Then
                mCarsOrContainers = mContainers
            Else
                mCarsOrContainers = mCar_num
            End If

            mWB = mWB_num

            If mWB_num < 9999 Then
                mWB = mWB_num + 646519
            End If

            If mCarsOrContainers < 9999 Then
                mCarsOrContainers = mCarsOrContainers + 646519
            End If

            If mWB >= mCarsOrContainers Then
                mCarfactor = CDec(mCarsOrContainers / mWB)
            Else
                mCarfactor = CDec(mWB / mCarsOrContainers)
            End If

            If Month(mWbDate) > mWbDate.Day Then
                mDayfactor = CDec(mWbDate.Day / mWbDate.Month)
            Else
                mDayfactor = CDec(mWbDate.Month / mWbDate.Day)
            End If

            mMaskfactor = CDec((mCarfactor * mDayfactor) + 1.57327)
            UnmaskNSValue = CDec(FormatNumber(mVal / mMaskfactor, 0))
        End If

    End Function

    Function UnmaskCSXValue(
         ByVal mSerial_No As String,
         ByVal mSTCC_W49 As String,
         ByVal mWbNum As Long,
         ByVal mAcctMonth As Integer,
         ByVal mAcctYear As Integer,
         ByVal mValue As Object) As Object

        Dim mSTCC_Compare As String
        Dim mWB As Integer
        Dim mIndex As Integer

        If mSerial_No = "" Then
            MsgBox("Error - No Serial Number value passed to CSX unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mSTCC_W49 = "" Then
            MsgBox("Error - No STCC_W49 value passed to CSX unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mWbNum = 0 Then
            MsgBox("Error - No Waybill Number value passed to CSX unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mAcctYear = 0 Then
            MsgBox("Error - No Account Year value passed to CSX unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mAcctMonth = 0 Then
            MsgBox("Error - No Account month value passed to CSX unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mValue = 0 Then
            MsgBox("Error - No masked value passed to CSX unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        UnmaskCSXValue = 0

        'Get the left 3 digits of the STCC_w49 value
        mSTCC_Compare = Mid(mSTCC_W49, 1, 3)

        If mAcctYear = 1990 Then
            mIndex = Array1DFindFirst(CSX1990STCC, mSTCC_Compare, 44)
            UnmaskCSXValue = CDec(FormatNumber(mValue * CSX1990Rate(mIndex), 0))
        ElseIf (mAcctYear < 1993) Or (mAcctYear = 1993 And mAcctMonth <= 8) Then
            mIndex = Array1DFindFirst(CSX1991STCC, mSTCC_Compare, 48)
            UnmaskCSXValue = CDec(FormatNumber(mValue * CSX1991Rate(mIndex), 0))
        ElseIf mAcctYear < 2020 Then
            'Post August 1993 Processing
            mIndex = Array1DFindFirst(CSX1991STCC, mSTCC_Compare, 48)
            'Get the right 2 digits of the mwbnum value
            mWB = CInt(Right(CStr(mWbNum), 2))
            If mWB < 20 Then
                UnmaskCSXValue = CDec(FormatNumber(mValue * CSXwb00(mIndex), 0))
            ElseIf mWB > 19 And mWB < 40 Then
                UnmaskCSXValue = CDec(FormatNumber(mValue * CSXwb20(mIndex), 0))
            ElseIf mWB > 39 And mWB < 60 Then
                UnmaskCSXValue = CDec(FormatNumber(mValue * CSXwb40(mIndex), 0))
            ElseIf mWB > 59 And mWB < 80 Then
                UnmaskCSXValue = CDec(FormatNumber(mValue * CSXwb60(mIndex), 0))
            ElseIf mWB > 79 Then
                UnmaskCSXValue = CDec(FormatNumber(mValue * CSXwb80(mIndex), 0))
            End If
        ElseIf mAcctYear > 2019 Then
            'Get the right 2 digits of the mwbnum value
            mWB = CInt(Right(CStr(mWbNum), 2))
            'get position of the STCC is in the R_Unmasking_CSX2020 table, zero if not found
            mIndex = Array1DFindFirst(CSX2020_STCC, mSTCC_W49, 0)           ' search for a 7 character match
            If mIndex = 0 Then
                mIndex = Array1DFindFirst(CSX2020_STCC, mSTCC_Compare, 0) ' Seach for a 3 character match
            End If
            If mIndex = 0 Then
                Select Case mWB
                    Case 0 To 19
                        UnmaskCSXValue = mValue * CSX2020_00(1)
                    Case 20 To 39
                        UnmaskCSXValue = mValue * CSX2020_20(1)
                    Case 40 To 59
                        UnmaskCSXValue = mValue * CSX2020_40(1)
                    Case 60 To 79
                        UnmaskCSXValue = mValue * CSX2020_60(1)
                    Case 80 To 99
                        UnmaskCSXValue = mValue * CSX2020_80(1)
                End Select
            Else
                Select Case mWB
                    Case 0 To 19
                        UnmaskCSXValue = mValue * CSX2020_00(mIndex)
                    Case 20 To 39
                        UnmaskCSXValue = mValue * CSX2020_20(mIndex)
                    Case 40 To 59
                        UnmaskCSXValue = mValue * CSX2020_40(mIndex)
                    Case 60 To 79
                        UnmaskCSXValue = mValue * CSX2020_60(mIndex)
                    Case 80 To 99
                        UnmaskCSXValue = mValue * CSX2020_80(mIndex)
                End Select
            End If
        End If

        UnmaskCSXValue = Math.Round(UnmaskCSXValue, 0, MidpointRounding.AwayFromZero)

    End Function

    Function UnmaskBNSFValue(
         ByVal mSerial_No As String,
         ByVal mU_Car_Init_Pos1 As String,
         ByVal m_AcctMonth As Integer,
         ByVal mVal As Object) As Object

        Dim mRow, mCol As Integer

        If mSerial_No = "" Then
            MsgBox("Error - No Serial Number value passed to BNSF unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mU_Car_Init_Pos1 = "" Then
            MsgBox("Error - No U_Car_Init_Pos1 value passed to BNSF unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If m_AcctMonth = 0 Then
            MsgBox("Error - No Account month value passed to BNSF unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mVal = 0 Then
            MsgBox("Error - No masked value passed to BNSF unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If
        mRow = 0
        mCol = 0

        'Find the array row entry
        Select Case m_AcctMonth
            Case 1, 4, 7, 10
                mRow = 1
            Case 2, 5, 8, 11
                mRow = 2
            Case 3, 6, 9, 12
                mRow = 3
        End Select

        'Find the array column entry
        Select Case mU_Car_Init_Pos1
            Case "A", "H", "N", "U"
                mCol = 1
            Case "B", "I", "O", "V"
                mCol = 2
            Case "C", "J", "P", "W"
                mCol = 3
            Case "D", "K", "Q", "X"
                mCol = 4
            Case "E", "L", "R", "Y"
                mCol = 5
            Case "F", "M", "S", "Z"
                mCol = 6
            Case "G", "T"
                mCol = 7
        End Select

        'return the value
        If (mRow = 0) Or (mCol = 0) Then
            UnmaskBNSFValue = 0
        Else
            UnmaskBNSFValue = CDec(FormatNumber(mVal / BNSFunmaskArray(mRow, mCol), 0))
        End If

    End Function

    Function UnmaskUPValue(
         ByVal mSerial_No As String,
         ByVal mSTCC_W49 As Long,
         ByVal mYear As Integer,
         ByVal mwb_num As Long,
         ByVal mVal As Object) As Object

        Dim mPos As Integer, mWaynum_Compare As Integer
        Dim mCol As Integer, mRow As Integer

        UnmaskUPValue = 0

        If mSerial_No = "" Then
            MsgBox("Error - No Serial_No value passed to UP unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mSTCC_W49.ToString = "" Then
            MsgBox("Error - No STCC_W49 value passed to UP unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mYear.ToString = "" Then
            MsgBox("Error - No year value passed to UP unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mwb_num.ToString = "" Then
            MsgBox("Error - No Waybill number value passed to UP unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        If mVal.ToString = "" Then
            MsgBox("Error - No masked value passed to UP unmasking function for serial Number " & mSerial_No,
                   vbOKOnly, "ERROR")
        End If

        'Get the value of the last 2 digits of the Waybill number

        If Len(CStr(mwb_num)) < 3 Then
            mWaynum_Compare = CInt(mwb_num)
        Else
            mWaynum_Compare = CInt(Right(CStr(mwb_num), 2))
        End If

        ' If the value is = 0 then use 100

        If mWaynum_Compare = 0 Then
            mWaynum_Compare = 100
        End If

        mCol = UPWbNum_Col(mWaynum_Compare)

        If mYear <= 2000 Then
            mPos = ArrayFind1993UPRange(mSTCC_W49, 0)
            mRow = UP1993STCC_Row(mPos)
        Else    'Year > 2000
            mPos = ArrayFind2001UPRange(mSTCC_W49, 0)
            mRow = UP2001STCC_Row(mPos)
        End If

        If UPfgrp(mRow, mCol) = 0 Then
            MsgBox("Divide by Zero Error for UP record", vbOKOnly, "ERROR!")
            Application.Exit()
        Else
            UnmaskUPValue = CDec(FormatNumber(mVal / UPfgrp(mRow, mCol), 0))
        End If

    End Function

End Module
