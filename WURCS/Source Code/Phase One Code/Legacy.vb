Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Threading

Public Class Legacy : Inherits Form

    Dim ConnString As String

    Private showText As Boolean = True
    Dim bPrint1 As Boolean
    Dim bPrint2 As Boolean
    Dim iYear As Integer

    Dim i2 As String
    Dim trd As New Thread(AddressOf myBackgroundThread)
    Dim tre As New Thread(AddressOf mywaitThread)
    Private Sub Legacy_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
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
        'EXIT the Form
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

    Private Sub btn_Cost_WayBillLegacy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Cost_WayBillLegacy.Click
        ' cost the waybill sample as legacy
        btn_GenMakeWhole.Enabled = False
        btn_ApplyMakeWhole.Enabled = False
        btn_commit.Enabled = False
        Label1.Text = "Working..."
        Label1.Visible = True

        Timer1.Interval = 1000
        Timer1.Start()
        Me.Refresh()
        'call the vb waybill sample as legacy
        'waybill costing on trd thread, timeout monitor on tre thread
        trd.IsBackground = True
        tre.IsBackground = True
        trd.Start()
        tre.Start()

    End Sub
    Private Sub mywaitThread()
        'test of background thread
        'timeout is one hour
        Thread.Sleep(3600000)
        'uncomment next line for testing timeout
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
                                       True, _
                                       False, _
                                       True, _
                                       True, _
                                       bPrint1, _
                                       bPrint2, _
                                       True, _
                                       True, _
                                       False, _
                                       False, _
                                       False, _
                                       False, _
                                       False, _
                                       True, _
                                       False, _
                                       False)
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
        Label1.Visible = True
        Me.Refresh()
    End Sub

    Private Delegate Sub updateLabel(ByVal newLabel As String)

    Private Sub updateLabelHandler(ByVal labelText As String)
        'MessageBox.Show("stop thread ")
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


    Private Sub btn_GenMakeWhole_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_GenMakeWhole.Click
        'generate make whole factors
        Label1.Text = "Working..."
        Label1.Visible = True

        Timer1.Interval = 1000
        Timer1.Start()
        Me.Refresh()
        'All we need to do to generate the make whole factors is get an instance of the GenerateMakeWholeFactors() class
        Dim clsGenMakeWhole As New GenerateMakeWholeFactors(ConnString, iYear, True)
        Timer1.Stop()
        Label1.Text = "Done"
        Me.Refresh()
    End Sub

    Private Sub btn_ApplyMakeWhole_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ApplyMakeWhole.Click
        'Apply Make Whole to CRPSEG
        Label1.Text = "Working..."
        Label1.Visible = True

        Timer1.Interval = 1000
        Timer1.Start()
        Me.Refresh()
        'All we need to do to generate the make whole factors is get an instance of the ApplyMakeWholeToCRPSEG() class
        Dim clsApplyMW As New ApplyMakeWholeToCRPSEG(ConnString, iYear, True)
        Timer1.Stop()
        Label1.Text = "Done"
        Me.Refresh()
    End Sub

    Private Sub btn_commit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_commit.Click
        'commit results to SQL
        Label1.Text = "Working..."
        Label1.Visible = True

        Timer1.Interval = 1000
        Timer1.Start()
        Me.Refresh()
        'All we need to do to write these out is get an instance of the CommitToSQL() class
        Dim clsCommit As New CommitToSQL(ConnString, iYear, True)
        Timer1.Stop()
        Label1.Text = "Done"
        Me.Refresh()
    End Sub


    Private Sub btn_errorLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_errorLog.Click
        'View the Error log
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

        Dim _legacyErrorLog As New LegacyErrorLog
        _legacyErrorLog.Show()
        Me.Hide()
        'MessageBox.Show("elapsed time is " & watch.ElapsedMilliseconds)
        Label1.Visible = False

    End Sub


End Class