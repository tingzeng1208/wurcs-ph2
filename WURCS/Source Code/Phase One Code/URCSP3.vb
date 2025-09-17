Imports System.Globalization
Imports System.Threading

Public Class URCSP3

    Public Class GlobalVariables
        Public Shared iYear As String
        Public Shared iWriteLoop1 As Boolean = False
        Public Shared iWriteLoop2 As Boolean = False
        'used in valid year check
        Public Shared iValidText As Boolean = True
        Public Shared CurrentYearMax As Integer
    End Class

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

        'prefill current year
        Dim uscalendar As New GregorianCalendar()
        Me.textEntry_iYear.Text = uscalendar.GetYear(Now)

        'set current year just established
        GlobalVariables.iYear = Me.textEntry_iYear.Text
        GlobalVariables.CurrentYearMax = Me.textEntry_iYear.Text
    End Sub


    Private Sub btn_Cost_WayBillLegacy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Cost_WayBillLegacy.Click
        'cost the way bill sample as legacy
        If GlobalVariables.iValidText Then
            Dim _legacy As New Legacy
            _legacy.Show()
            Me.Hide()
        Else
            MessageBox.Show("Entered Year " & Me.textEntry_iYear.Text & " is invalid, ")
            Me.textEntry_iYear.Clear()
            Me.textEntry_iYear.Focus()
        End If
    End Sub

    Private Sub textEntry_iYear_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles textEntry_iYear.TextChanged
        'check valid year entry
        If IsInputNumeric(textEntry_iYear.Text) Then
            'year entered was valid
            GlobalVariables.iValidText = True
        Else
            'year entered was invalid 
            GlobalVariables.iValidText = False
        End If

    End Sub

    Private Sub btn_Cost_WayBill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Cost_WayBill.Click
        'cost the waybill sample
        If GlobalVariables.iValidText Then
            Dim _production As New Production
            _production.Show()
            Me.Hide()
        Else
            MessageBox.Show("Entered Year " & Me.textEntry_iYear.Text & " is invalid, ")
            Me.textEntry_iYear.Clear()
            Me.textEntry_iYear.Focus()
        End If
    End Sub

    Private Sub btn_exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_exit.Click
        ' exit
        Application.Exit()
    End Sub

    Private Sub cb_iWriteLoop1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_iWriteLoop1.CheckedChanged
        'loop 1
        GlobalVariables.iWriteLoop1 = cb_iWriteLoop1.Checked()
        'If GlobalVariables.iWriteLoop1 Then
        '    MessageBox.Show("checked")
        'Else
        '    MessageBox.Show("unchecked")
        'End If
    End Sub

    Private Sub cb_iWriteLoop2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_iWriteLoop2.CheckedChanged
        'loop2
        GlobalVariables.iWriteLoop2 = cb_iWriteLoop2.Checked()
    End Sub

    Private Function IsInputNumeric(ByVal input As String) As Boolean
        If String.IsNullOrWhiteSpace(input) Then Return False
        'MessageBox.Show("year max is " & GlobalVariables.CurrentYearMax)
        If IsNumeric(input) Then
            Dim nYear As Integer
            nYear = CUInt(input)
            If nYear > GlobalVariables.CurrentYearMax And GlobalVariables.CurrentYearMax > 0 Then Return False
            If nYear > 1000 And nYear < 9999 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If

    End Function

End Class
