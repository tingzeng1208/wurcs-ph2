
Public Class SplashScreen

    Private Sub SplashScreen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim mWorkstr As String

        Me.Show()

        'Set the form so it centers on the user's screen
        CenterToScreen()
        Application.DoEvents()
        Refresh()

        'Set the version number of the program into the Main Menu
        txt_VersionNo.Text = Global_Variables.Version_Number
        Refresh()
        Application.DoEvents()

        'Determine which data source we're connected to - LAN Sql server, Laptop SQL server or None.
        If Len(Global_Variables.gbl_SQLConn) = 0 Then
            ' Set the environment to not connected to a database
            ' Default: We're not connected to a SQL Server
            Global_Variables.gbl_SQLConn = ""
            Global_Variables.gbl_IsWriter = True

            
            mWorkstr = "Not connected to SQL database."
            Me.txt_ConnectedBox.ForeColor = Color.Red

            ' Now we'll try to see if we're on the STB network
            txt_ConnectedBox.Text = "Looking for database on STB Network..."
            Refresh()
            Application.DoEvents()
            If IO.Directory.Exists("\\STBHQSQLOE\E$") Then
                Global_Variables.gbl_SQLConn = "Server=STBHQSQLOE" & _
                        ";Database=URCS;Trusted_Connection = True;"
                mWorkstr = "Connected to Database on STBHQSQLOE"
                gbl_DatabaseLocation = mWorkstr
                txt_ConnectedBox.Text = gbl_DatabaseLocation

            Else

                ' Now we'll have to default to local database
                txt_ConnectedBox.Text = "Not on STB Network.  Checking local machine..."
                Refresh()
                Application.DoEvents()
                Global_Variables.gbl_SQLConn = "Server=" & My.Computer.Name & _
                    ";Database=URCS;Trusted_Connection = True;"
                Me.txt_ConnectedBox.ForeColor = Color.Green
                mWorkstr = "Defaulted to Local Database on " & My.Computer.Name
                gbl_DatabaseLocation = mWorkstr
                txt_ConnectedBox.Text = gbl_DatabaseLocation
                Refresh()
                Application.DoEvents()
            End If
        End If

        txt_Advisory.Text = "Initializing...  Please wait."
        Refresh()
        Application.DoEvents()

        'Now we load the MainMenu and close the SplashScreen
        ' Open the subform
        Dim frmNew As New frm_MainMenu()
        frmNew.Show()
        ' Close the SplashScreen
        Me.Close()
    End Sub


End Class
