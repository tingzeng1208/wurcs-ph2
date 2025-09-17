<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MsgForm
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pic_STB_Logo = New System.Windows.Forms.PictureBox()
        Me.MyMsgText = New System.Windows.Forms.TextBox()
        CType(Me.pic_STB_Logo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(106, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(212, 46)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Surface Transportation Board" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "URCS && Waybills"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'pic_STB_Logo
        '
        Me.pic_STB_Logo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pic_STB_Logo.Image = Global.WURCS.My.Resources.Resources.stb_seal_sm
        Me.pic_STB_Logo.Location = New System.Drawing.Point(161, 12)
        Me.pic_STB_Logo.Name = "pic_STB_Logo"
        Me.pic_STB_Logo.Size = New System.Drawing.Size(102, 105)
        Me.pic_STB_Logo.TabIndex = 11
        Me.pic_STB_Logo.TabStop = False
        '
        'MyMsgText
        '
        Me.MyMsgText.Location = New System.Drawing.Point(32, 184)
        Me.MyMsgText.Name = "MyMsgText"
        Me.MyMsgText.Size = New System.Drawing.Size(361, 21)
        Me.MyMsgText.TabIndex = 13
        Me.MyMsgText.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'MsgForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(427, 217)
        Me.ControlBox = False
        Me.Controls.Add(Me.MyMsgText)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.pic_STB_Logo)
        Me.Name = "MsgForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.pic_STB_Logo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pic_STB_Logo As System.Windows.Forms.PictureBox
    Friend WithEvents MyMsgText As System.Windows.Forms.TextBox
End Class
