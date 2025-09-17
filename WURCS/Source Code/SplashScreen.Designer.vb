<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SplashScreen
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txt_VersionNo = New System.Windows.Forms.TextBox()
        Me.pic_STB_Logo = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_ConnectedBox = New System.Windows.Forms.TextBox()
        Me.txt_Advisory = New System.Windows.Forms.TextBox()
        CType(Me.pic_STB_Logo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txt_VersionNo
        '
        Me.txt_VersionNo.BackColor = System.Drawing.SystemColors.Control
        Me.txt_VersionNo.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_VersionNo.Location = New System.Drawing.Point(145, 169)
        Me.txt_VersionNo.Name = "txt_VersionNo"
        Me.txt_VersionNo.Size = New System.Drawing.Size(102, 14)
        Me.txt_VersionNo.TabIndex = 0
        Me.txt_VersionNo.TabStop = False
        Me.txt_VersionNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'pic_STB_Logo
        '
        Me.pic_STB_Logo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pic_STB_Logo.Image = Global.URCS_And_Waybills.My.Resources.Resources.stb_seal_sm
        Me.pic_STB_Logo.Location = New System.Drawing.Point(145, 12)
        Me.pic_STB_Logo.Name = "pic_STB_Logo"
        Me.pic_STB_Logo.Size = New System.Drawing.Size(102, 105)
        Me.pic_STB_Logo.TabIndex = 8
        Me.pic_STB_Logo.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(92, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(212, 46)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Surface Transportation Board" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "URCS && Waybills"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txt_ConnectedBox
        '
        Me.txt_ConnectedBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_ConnectedBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_ConnectedBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_ConnectedBox.Location = New System.Drawing.Point(15, 191)
        Me.txt_ConnectedBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_ConnectedBox.Name = "txt_ConnectedBox"
        Me.txt_ConnectedBox.Size = New System.Drawing.Size(362, 14)
        Me.txt_ConnectedBox.TabIndex = 0
        Me.txt_ConnectedBox.TabStop = False
        Me.txt_ConnectedBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txt_Advisory
        '
        Me.txt_Advisory.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_Advisory.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_Advisory.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_Advisory.Location = New System.Drawing.Point(15, 222)
        Me.txt_Advisory.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_Advisory.Name = "txt_Advisory"
        Me.txt_Advisory.Size = New System.Drawing.Size(362, 14)
        Me.txt_Advisory.TabIndex = 11
        Me.txt_Advisory.TabStop = False
        Me.txt_Advisory.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'SplashScreen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(392, 250)
        Me.ControlBox = False
        Me.Controls.Add(Me.txt_Advisory)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.pic_STB_Logo)
        Me.Controls.Add(Me.txt_ConnectedBox)
        Me.Controls.Add(Me.txt_VersionNo)
        Me.Name = "SplashScreen"
        Me.Text = "Initializing URCS & Waybills"
        CType(Me.pic_STB_Logo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_VersionNo As System.Windows.Forms.TextBox
    Friend WithEvents pic_STB_Logo As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txt_ConnectedBox As System.Windows.Forms.TextBox
    Friend WithEvents txt_Advisory As System.Windows.Forms.TextBox

End Class
