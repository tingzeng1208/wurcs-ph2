Module Masking_Routines

    Function MaskGenericValue(
         ByVal mSTCC_W49 As String,
         ByVal mWB_Date As Date,
         ByVal mAcctYear As Integer,
         ByVal mValue As Decimal) As Integer

        Dim mCompare, mCol, mRow As Integer
        Dim mResult As Decimal

        ' As of early June of 2021, this routine was modified to use the AcctYear value
        ' rather than the Waybill Year.
        ' MRS

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
            mResult = mValue * CDec(RR_Odd_Factor(mRow, mCol))
        Else
            mResult = mValue * CDec(RR_Even_Factor(mRow, mCol))
        End If

        MaskGenericValue = Math.Round(mResult, 0, MidpointRounding.AwayFromZero)
    End Function

End Module
