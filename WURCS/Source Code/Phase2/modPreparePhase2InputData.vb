Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Public Class Phase2InputData

#Region "member variables"

    'global usage vars
    Dim sProcessingYear As String
    Dim iAcodeID As Integer
    Dim sAcode As String
    Dim iAcodeCounter As Integer
    Dim sAColumn As String
    Dim sALine As String
    Dim sLineA As String
    Dim sAPart As String
    Dim sCode As String
    Dim sCarTypeID As String
    Dim iCarTypeID As Integer
    Dim iRRICC As Integer
    Dim iRRICCRegion As Integer
    Dim iRRICCNation As Integer
    Dim iRR_id As Integer
    Dim dValue As Double
    Dim iSCH As Integer
    Dim iLINE As Integer
    Dim iCOL As Integer
    Dim iSCALE As Integer
    Dim iLoadCode As Integer
    Dim sRpt_Sheet As String
    Dim iThisRow As Integer = 0

    'local worktables
    Dim dtTrans As DataTable
    Dim dtDictionary As DataTable
    Dim dtRailroadsToProcess As DataTable

    Dim oDB As DBManager
    Dim sCurrentYear As String
    Public Event StatusUpdated(Status As String, Completed As Boolean)
    Public Event ErrorOccured(Message As String)

#End Region

    Public Sub New(ByVal CurrentYear As String, ByVal DB As DBManager)

        sCurrentYear = CurrentYear
        oDB = DB

    End Sub

    Public Function Prepare() As Boolean

        'Commented out to return actual error to user 1/10/22 - MRS
        'Try

        Dim oRegionValues As Dictionary(Of Integer, Double)

        RaiseEvent StatusUpdated("Started the preparation process", False)

        'Write row to error table to announce beginng of processing
        oDB.HandleError(My.Settings.CurrentYear.ToString, "Status", "Started PreparePhase2InputData Process", "Started PreparePhase2InputData Process", "Started PreparePhase2InputData Process")

        'get local copies of the tables needed
        dtTrans = oDB.GetTransTable(sCurrentYear)
        dtDictionary = oDB.GetDataDictionary(sCurrentYear)
        dtRailroadsToProcess = oDB.GetClass1RailDataToPrepare(sCurrentYear)

        'truncate the A tables
        oDB.TruncateATables()

        'check and adjust trans data
        TransModifications(sCurrentYear)

        'init iEcodeID
        iAcodeID = 1

        'Determine how many records we're dealing with
        iAcodeCounter = dtDictionary.Rows.Count

        'loop through the data dictionary and build the AVALUEs
        For Each dr As DataRow In dtDictionary.Rows

            'increment the row counter
            iThisRow = iThisRow + 1

            'First deal with UR_ACODE
            sAcode = dr("WTALL").ToString()
            sAColumn = dr("WTALL").ToString.Substring(dr("WTALL").ToString.IndexOf("C") + 1, 1)
            sALine = dr("WTALL").ToString.Substring(dr("WTALL").ToString.IndexOf("L") + 1, 3)
            sLineA = dr("WTALL").ToString.Substring(dr("WTALL").ToString.IndexOf("L"), 4)
            sAPart = dr("WTALL").ToString.Substring(0, 2)
            sCode = dr("WTALL").ToString.Substring(dr("WTALL").ToString.IndexOf("C"), 2)

            'Define the worksheet identifier
            If sAPart = "A1" Then
                If Integer.Parse(sALine) >= 101 And Integer.Parse(sALine) <= 160 Then
                    sRpt_Sheet = "A1P1"
                ElseIf Integer.Parse(sALine) >= 201 And Integer.Parse(sALine) <= 216 Then
                    sRpt_Sheet = "A1P2A"
                ElseIf Integer.Parse(sALine) >= 217 And Integer.Parse(sALine) <= 235 Then
                    sRpt_Sheet = "A1P2B"
                ElseIf Integer.Parse(sALine) >= 236 And Integer.Parse(sALine) <= 254 Then
                    sRpt_Sheet = "A1P2C"
                ElseIf Integer.Parse(sALine) >= 301 And Integer.Parse(sALine) <= 324 Then
                    sRpt_Sheet = "A1P3A"
                ElseIf Integer.Parse(sALine) >= 341 And Integer.Parse(sALine) <= 364 Then
                    sRpt_Sheet = "A1P3B"
                ElseIf Integer.Parse(sALine) >= 401 And Integer.Parse(sALine) <= 482 Then
                    sRpt_Sheet = "A1P4"
                ElseIf Integer.Parse(sALine) >= 501 And Integer.Parse(sALine) <= 516 Then
                    sRpt_Sheet = "A1P5A"
                ElseIf Integer.Parse(sALine) >= 521 And Integer.Parse(sALine) <= 536 Then
                    sRpt_Sheet = "A1P5B"
                ElseIf Integer.Parse(sALine) >= 541 And Integer.Parse(sALine) <= 556 Then
                    sRpt_Sheet = "A1P6"
                ElseIf Integer.Parse(sALine) >= 561 And Integer.Parse(sALine) <= 576 Then
                    sRpt_Sheet = "A1P7"
                ElseIf Integer.Parse(sALine) >= 580 And Integer.Parse(sALine) <= 595 Then
                    sRpt_Sheet = "A1P8"
                ElseIf Integer.Parse(sALine) >= 901 And Integer.Parse(sALine) <= 918 Then
                    sRpt_Sheet = "A1P9"
                End If
            ElseIf sAPart = "A2" Then
                If Integer.Parse(sALine) >= 101 And Integer.Parse(sALine) <= 184 Then
                    sRpt_Sheet = "A2P1"
                ElseIf Integer.Parse(sALine) >= 201 And Integer.Parse(sALine) <= 262 Then
                    sRpt_Sheet = "A2P2"
                ElseIf Integer.Parse(sALine) >= 301 And Integer.Parse(sALine) <= 366 Then
                    sRpt_Sheet = "A2P3"
                ElseIf Integer.Parse(sALine) >= 401 And Integer.Parse(sALine) <= 422 Then
                    sRpt_Sheet = "A2P4"
                End If
            ElseIf sAPart = "A3" Then
                If Integer.Parse(sALine) >= 101 And Integer.Parse(sALine) <= 178 Then
                    sRpt_Sheet = "A3P1"
                ElseIf Integer.Parse(sALine) >= 201 And Integer.Parse(sALine) <= 224 Then
                    sRpt_Sheet = "A3P2"
                ElseIf Integer.Parse(sALine) >= 301 And Integer.Parse(sALine) <= 344 Then
                    sRpt_Sheet = "A3P3"
                ElseIf Integer.Parse(sALine) >= 401 And Integer.Parse(sALine) <= 444 Then
                    sRpt_Sheet = "A3P4"
                ElseIf Integer.Parse(sALine) >= 501 And Integer.Parse(sALine) <= 543 Then
                    sRpt_Sheet = "A3P5"
                ElseIf Integer.Parse(sALine) >= 601 And Integer.Parse(sALine) <= 643 Then
                    sRpt_Sheet = "A3P6"
                ElseIf Integer.Parse(sALine) >= 701 And Integer.Parse(sALine) <= 728 Then
                    sRpt_Sheet = "A3P7"
                ElseIf Integer.Parse(sALine) >= 801 And Integer.Parse(sALine) <= 829 Then
                    sRpt_Sheet = "A3P8"
                End If
            ElseIf sAPart = "A4" Then
                If Integer.Parse(sALine) >= 101 And Integer.Parse(sALine) <= 145 Then
                    sRpt_Sheet = "A4P1"
                ElseIf Integer.Parse(sALine) >= 170 And Integer.Parse(sALine) <= 178 Then
                    sRpt_Sheet = "A4P2"
                ElseIf Integer.Parse(sALine) >= 201 And Integer.Parse(sALine) <= 205 Then
                    sRpt_Sheet = "A4P3"
                End If
            ElseIf sAPart = "E2" Then
                sRpt_Sheet = "E2P1"
            End If

            oDB.WriteURAcodeData(sAcode, sAColumn, sALine, sLineA, sAPart, sCode, sRpt_Sheet)

            'Now deal with U_ATABLES 
            iSCH = Integer.Parse(dr("SCH"))
            iLINE = Integer.Parse(dr("LINE"))
            iCOL = Integer.Parse(dr("Column"))
            iSCALE = Integer.Parse(dr("Scaler"))
            iLoadCode = Integer.Parse(dr("LoadCode"))

            Dim i As Integer = Integer.Parse(sCurrentYear - 4)

            Do While i <= Integer.Parse(sCurrentYear)

                oRegionValues = New Dictionary(Of Integer, Double)
                For Each drRR As DataRow In dtRailroadsToProcess.Rows

                    sProcessingYear = i.ToString()
                    iRR_id = Integer.Parse(drRR("RR_ID"))
                    iRRICC = Integer.Parse(drRR("RRICC"))
                    iRRICCRegion = Integer.Parse(drRR("RegionRRICC"))
                    iRRICCNation = Integer.Parse(drRR("NationRRICC"))

                    If Not oRegionValues.Keys.Contains(iRRICCRegion) Then
                        oRegionValues.Add(iRRICCRegion, 0)
                    End If

                    If iRRICC > 900000 And iLoadCode = 0 Then
                        dValue = oRegionValues(iRRICCRegion)
                    Else
                        dValue = GetValue(sProcessingYear)
                        oRegionValues(iRRICCRegion) = oRegionValues(iRRICCRegion) + dValue
                    End If

                    oDB.writeUAValues(sProcessingYear, iRR_id.ToString(), iAcodeID.ToString(), dValue.ToString())

                Next
                i = i + 1
            Loop

            'increment the aCode_id
            iAcodeID = iAcodeID + 1
            RaiseEvent StatusUpdated("Processed data for " & sAcode.ToString _
                                     & " Record " & iThisRow.ToString & " of " & iAcodeCounter.ToString, False)

        Next

        'RaiseEvent StatusUpdated("All Phase 2 input data for year " & sCurrentYear & " has been prepared", False)

        'Commented out 1/10/22 to get actual error to user and not to database - MRS
        'Catch ex As system.exception
        '    'write the error out to the sql error table
        '    oDB.HandleError(sCurrentYear.ToString, "Error", ex.Message.ToString(), ex.StackTrace, "PreparePhase2InputData-New")
        '    RaiseEvent ErrorOccured(ex.Message)

        '    Return False
        'End Try

        ''Write row to error table to announce end of processing
        'oDB.HandleError(sCurrentYear.ToString, "Status", "Ended PreparePhase2InputData Process", "Ended PreparePhase2InputData Process", "Ended PreparePhase2InputData Process")

        Return True

    End Function

#Region "Private Methods"

    ''' <summary>
    ''' Loops through data in the dtTrans table and makes updates to local and SQL data if necessary
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub TransModifications(CurrentYear As String)

        Dim iProcessYear As Integer
        Dim fValue As Double

        'run the check for the current year and the 4 previous years
        For iProcessYear = Integer.Parse(CurrentYear) - 4 To Integer.Parse(CurrentYear)

            'filter to only SCH between 100 and 148
            For Each dr As DataRow In dtTrans.Select("Year = " & iProcessYear.ToString() & " and SCH > 100 and SCH < 148")

                'check C12: Total Carloads Originated & Terminated (CLOT) (A1L144C1)
                fValue = (dr("C1") * 2) + dr("C3") + dr("C5")
                If Not fValue = Double.Parse(dr("C12")) Then
                    dr("C12") = fValue
                    oDB.AdjustUTransValues(12, fValue, iProcessYear.ToString(), dr("RRICC").ToString(), dr("SCH").ToString(), dr("LINE").ToString())
                End If

                'check C13: Total Carlaods Handled (CLOR) (A1L145C1)
                fValue = dr("C1") + dr("C3") + dr("C5") + dr("C7")
                If Not fValue = Double.Parse(dr("C13")) Then
                    dr("C13") = fValue
                    oDB.AdjustUTransValues(13, fValue, iProcessYear.ToString(), dr("RRICC").ToString(), dr("SCH").ToString(), dr("LINE").ToString())
                End If

                'check C14: Total Carloads Interchanged (CLRF) (A1L146C1)
                fValue = dr("C3") + dr("C5") + (dr("C7") * 2)
                If Not fValue = Double.Parse(dr("C14")) Then
                    dr("C14") = fValue
                    oDB.AdjustUTransValues(14, fValue, iProcessYear.ToString(), dr("RRICC").ToString(), dr("SCH").ToString(), dr("LINE").ToString())
                End If
            Next

            'filter to only SCH 420
            For Each dr As DataRow In dtTrans.Select("Year = " & iProcessYear.ToString() & " and SCH = 420")

                'check C12: Total Depreciation Expense (A3L401C2-A3L441C2)
                fValue = dr("C2") + dr("C3")
                If Not fValue = Double.Parse(dr("C12")) Then
                    dr("C12") = fValue
                    oDB.AdjustUTransValues(12, fValue, iProcessYear.ToString(), dr("RRICC").ToString(), dr("SCH").ToString(), dr("LINE").ToString())
                End If
            Next

            'filter to only SCH 33, LINE 57
            For Each dr As DataRow In dtTrans.Select("Year = " & iProcessYear.ToString() & " and SCH = 33 and LINE = 57")

                'check C8: Miles of Running Track (A1L151C1)
                fValue = dr("C1") + dr("C2") + dr("C3") + dr("C4")
                If Not fValue = Double.Parse(dr("C8")) Then
                    dr("C8") = fValue
                    oDB.AdjustUTransValues(8, fValue, iProcessYear.ToString(), dr("RRICC").ToString(), dr("SCH").ToString(), dr("LINE").ToString())
                End If

                'check C9: Miles of Switching Track (A1L154C1)
                fValue = dr("C5") + dr("C6")
                If Not fValue = Double.Parse(dr("C9")) Then
                    dr("C9") = fValue
                    oDB.AdjustUTransValues(9, fValue, iProcessYear.ToString(), dr("RRICC").ToString(), dr("SCH").ToString(), dr("LINE").ToString())
                End If

                'check C10: Miles of Road Track (A1L155C1)
                fValue = dr("C5") + dr("C8")
                If Not fValue = Double.Parse(dr("C10")) Then
                    dr("C10") = fValue
                    oDB.AdjustUTransValues(10, fValue, iProcessYear.ToString(), dr("RRICC").ToString(), dr("SCH").ToString(), dr("LINE").ToString())
                End If
            Next
        Next
    End Sub

    ''' <summary>
    ''' Gets a double value for the data retrieved from the data dictionary
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetValue(ByVal sYear As String) As Double

        Dim d As Double = 0.0
        Dim s As String = ""
        Dim iRRICCcode As Integer = 0

        Select Case iLoadCode
            Case 1, 4, 5 : iRRICCcode = iRRICCRegion
            Case 2 : iRRICCcode = iRRICCNation
            Case 3 : iRRICCcode = iRRICC
            Case Else : iRRICCcode = iRRICC
        End Select

        s += "Year = " & sYear
        s += " And RRICC = " & iRRICCcode.ToString()
        s += " And SCH = " & iSCH.ToString()
        s += " And LINE = " & iLINE.ToString()

        For Each drTrans As DataRow In dtTrans.Select(s)
            Select Case iCOL
                Case 1
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C1")))
                Case 2
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C2")))
                Case 3
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C3")))
                Case 4
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C4")))
                Case 5
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C5")))
                Case 6
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C6")))
                Case 7
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C7")))
                Case 8
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C8")))
                Case 9
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C9")))
                Case 10
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C10")))
                Case 11
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C11")))
                Case 12
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C12")))
                Case 13
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C13")))
                Case 14
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C14")))
                Case 15
                    d = Scale(iSCALE, ReturnDecimal(drTrans("C15")))
            End Select
        Next

        GetValue = Double.Parse(d)

    End Function

    ''' <summary>
    ''' Applies scaling to the value
    ''' </summary>
    ''' <param name="scaler">Scaler value from the Data Dictionary</param>
    ''' <param name="value">original value</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Scale(ByVal scaler As Integer, ByVal value As Double) As Double
        If (scaler = 1) Then
            Return value / 10
        ElseIf (scaler = 2) Then
            Return value / 100
        ElseIf (scaler = 3) Then
            Return value / 1000
        ElseIf (scaler = 4) Then
            Return value / 10000
        ElseIf (scaler = 5) Then
            Return value / 100000
        ElseIf (scaler = 6) Then
            Return value / 1000000
        Else
            Return value
        End If
    End Function

#End Region

End Class
