Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Threading

Public Class Production : Inherits Form
    Dim ConnString As String

    Private showText As Boolean = True
    Dim bPrint1 As Boolean
    Dim bPrint2 As Boolean
    Dim iYear As Integer
    Dim i2 As String
    Dim trd As New Thread(AddressOf myBackgroundThread)
    Dim tre As New Thread(AddressOf mywaitThread)

    Private Sub Production_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        'MessageBox.Show("Hello, " & URCSP3.GlobalVariables.iYear)
        iYear = URCSP3.GlobalVariables.iYear
        bPrint1 = URCSP3.GlobalVariables.iWriteLoop1
        bPrint2 = URCSP3.GlobalVariables.iWriteLoop2

        'Get the connection string from the app.config
        ConnString = New String(ConfigurationManager.AppSettings("ConnectionString"))
        'MessageBox.Show("connection string is  " & ConnString)

        Label1.Visible = False

    End Sub

    Private Sub btn_exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_exit.Click
        'exit the Form
        If trd.IsAlive Then
            trd.Abort()
            'MessageBox.Show("thread trd was alive")
        End If
        If tre.IsAlive Then
            tre.Abort()
            'MessageBox.Show("thread tre was alive")
        End If
        Dim _URCSP3 As New URCSP3
        _URCSP3.Show() 'redisplay main form residing in memory
        Me.Close()
    End Sub

    Private Sub btn_Cost_WayBill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Cost_WayBill.Click
        Label1.Text = "Working..."
        Label1.Visible = True
        btn_GenMakeWhole.Enabled = False
        btn_ApplyMakeWhole.Enabled = False
        btn_commit.Enabled = False

        Timer1.Interval = 1000
        Timer1.Start()
        Me.Refresh()
        'All we need to do to cost the sample is get an instance of the CostWayBill() class
        trd.IsBackground = True
        tre.IsBackground = True
        trd.Start()
        tre.Start()
    End Sub

    Private Sub btn_GenMakeWhole_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_GenMakeWhole.Click
        Label1.Text = "Working..."
        Label1.Visible = True

        Timer1.Interval = 1000
        Timer1.Start()
        Me.Refresh()
        'All we need to do to generate the make whole factors is get an instance of the GenerateMakeWholeFactors() class
        Dim clsGenMakeWhole As New GenerateMakeWholeFactors(ConnString, iYear, False)

        Timer1.Stop()
        Label1.Text = "Done"
        Me.Refresh()
    End Sub

    Private Sub btn_ApplyMakeWhole_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ApplyMakeWhole.Click
        Label1.Text = "Working..."
        Label1.Visible = True

        Timer1.Interval = 1000
        Timer1.Start()
        Me.Refresh()
        'All we need to do to generate the make whole factors is get an instance of the ApplyMakeWholeToCRPSEG() class
        Dim clsApplyMW As New ApplyMakeWholeToCRPSEG(ConnString, iYear, False)

        Timer1.Stop()
        Label1.Text = "Done"
        Me.Refresh()
    End Sub

    Private Sub btn_commit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_commit.Click
        Label1.Text = "Working..."
        Label1.Visible = True

        Timer1.Interval = 1000
        Timer1.Start()
        Me.Refresh()
        'All we need to do to write these out is get an instance of the CommitToSQL() class
        Dim clsCommit As New CommitToSQL(ConnString, iYear, False)

        Timer1.Stop()
        Label1.Text = "Done"
        Me.Refresh()
    End Sub

    Private Sub btn_errorLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_errorLog.Click
        'get the error table from SQL
        If Label1.Text = "" Then
            Label1.Text = "Working..."
            Label1.Enabled = True
        End If
        Label1.Text = "Working..."
        Label1.Visible = True
        Me.Refresh()

        Dim watch = Stopwatch.StartNew

        Do
            'do work here
        Loop Until watch.ElapsedMilliseconds >= 2000

        Dim _ProductionErrorLog As New ProductionErrorLog
        _ProductionErrorLog.Show()
        Me.Hide()
        'MessageBox.Show("elapsed time is " & watch.ElapsedMilliseconds)
        Label1.Visible = False
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Static labelText As String = Label1.Text
        i2 = Label1.Text
        showText = Not showText

        If showText = True Then
            Label1.Text = labelText
        Else
            Label1.Text = String.Empty
        End If
        'Do Loop stuff
        Timer1.Enabled = True

    End Sub
    Private Sub mywaitThread()
        'test of background thread
        'timeout is one hour
        Thread.Sleep(3600000)
        'for testing, next line
        'Thread.Sleep(30000)
        If Me.InvokeRequired Then 'are we running on a secondary thread
            Dim d As New updateLabel(AddressOf updateLabelHandler2)
            Me.Invoke(d, New Object() {"working"})
        Else
            updateLabelHandler2("")
        End If
        trd.Abort()
        'MessageBox.Show("stop all threads ")
        tre.Abort()

    End Sub

    Private Sub myBackgroundThread()
        If Me.InvokeRequired Then 'are we running on a secondary thread
            Dim d As New updateLabel(AddressOf updateLabelHandler)
            Me.Invoke(d, New Object() {"start-thread"})
            'call the vb waybill sample as legacy
            'All we need to do to cost the sample is get an instance of the CostWayBill() class
            Dim cWB As New CostWayBill(iYear, _
                                       False, _
                                       cb_FilterNonUS.Checked, _
                                       cb_RoundTotalMiles.Checked, _
                                       cb_RoundMiles.Checked, _
                                       bPrint1, _
                                       bPrint2, _
                                       cb_FilterSingleTotalDist.Checked, _
                                       cb_FilterRev0.Checked, _
                                       cb_ChangeTons.Checked, _
                                       cb_ChangeTotalDist.Checked, _
                                       cb_ChangeTSCAssgn.Checked, _
                                       cb_UseCarOwn.Checked, _
                                       cb_Adjust257.Checked, _
                                       cb_adjustRR769.Checked, _
                                       cb_Adjust307.Checked, _
                                       cb_updatedSTCC_Code.Checked)
        Else
            updateLabelHandler("")
        End If

        If trd.IsAlive Then
            trd.Abort()
        End If
        If tre.IsAlive Then
            tre.Abort()
        End If
        btn_GenMakeWhole.Enabled = True
        btn_ApplyMakeWhole.Enabled = True
        btn_commit.Enabled = True

        'MessageBox.Show("thread is over ")
        Timer1.Stop()
        Label1.Text = "Done"
        Me.Refresh()
    End Sub

    Private Delegate Sub updateLabel(ByVal newLabel As String)

    Private Sub updateLabelHandler(ByVal labelText As String)
        Label1.Text = "Working..."
        Label1.Visible = True
        Me.Refresh()
    End Sub

    Private Sub updateLabelHandler2(ByVal labelText As String)
        Timer1.Stop()
        Label1.Text = "Timed Out"
        Label1.Visible = True
        Me.Refresh()
        btn_GenMakeWhole.Enabled = True
        btn_ApplyMakeWhole.Enabled = True
        btn_commit.Enabled = True

        'Label1.Text = labelText
    End Sub

End Class