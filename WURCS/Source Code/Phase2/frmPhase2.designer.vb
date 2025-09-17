<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPhase2
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
        Me.grpRun = New System.Windows.Forms.GroupBox()
        Me.btnLaunch = New System.Windows.Forms.Button()
        Me.rbAllSteps = New System.Windows.Forms.RadioButton()
        Me.btnEValues = New System.Windows.Forms.Button()
        Me.rbStepByStep = New System.Windows.Forms.RadioButton()
        Me.btnReport = New System.Windows.Forms.Button()
        Me.ssStatus = New System.Windows.Forms.StatusStrip()
        Me.tssLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.txtFolder = New System.Windows.Forms.TextBox()
        Me.grpParameters = New System.Windows.Forms.GroupBox()
        Me.cbLog = New System.Windows.Forms.CheckBox()
        Me.cbAll = New System.Windows.Forms.CheckBox()
        Me.clbRailroads = New System.Windows.Forms.CheckedListBox()
        Me.lblYear = New System.Windows.Forms.Label()
        Me.cmbYear = New System.Windows.Forms.ComboBox()
        Me.btnFolder = New System.Windows.Forms.Button()
        Me.grpRun.SuspendLayout()
        Me.ssStatus.SuspendLayout()
        Me.grpParameters.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpRun
        '
        Me.grpRun.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpRun.Controls.Add(Me.btnLaunch)
        Me.grpRun.Controls.Add(Me.rbAllSteps)
        Me.grpRun.Controls.Add(Me.btnEValues)
        Me.grpRun.Controls.Add(Me.rbStepByStep)
        Me.grpRun.Controls.Add(Me.btnReport)
        Me.grpRun.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.grpRun.Location = New System.Drawing.Point(12, 272)
        Me.grpRun.Name = "grpRun"
        Me.grpRun.Size = New System.Drawing.Size(525, 70)
        Me.grpRun.TabIndex = 1
        Me.grpRun.TabStop = False
        Me.grpRun.Text = "Create Output Report and Generate E-Values"
        '
        'btnLaunch
        '
        Me.btnLaunch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLaunch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLaunch.Location = New System.Drawing.Point(396, 41)
        Me.btnLaunch.Name = "btnLaunch"
        Me.btnLaunch.Size = New System.Drawing.Size(123, 23)
        Me.btnLaunch.TabIndex = 4
        Me.btnLaunch.Text = "Execute"
        Me.btnLaunch.UseVisualStyleBackColor = True
        '
        'rbAllSteps
        '
        Me.rbAllSteps.AutoSize = True
        Me.rbAllSteps.Checked = True
        Me.rbAllSteps.Location = New System.Drawing.Point(396, 20)
        Me.rbAllSteps.Name = "rbAllSteps"
        Me.rbAllSteps.Size = New System.Drawing.Size(122, 17)
        Me.rbAllSteps.TabIndex = 1
        Me.rbAllSteps.TabStop = True
        Me.rbAllSteps.Text = "All Steps in One Run"
        Me.rbAllSteps.UseVisualStyleBackColor = True
        '
        'btnEValues
        '
        Me.btnEValues.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnEValues.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnEValues.Location = New System.Drawing.Point(137, 41)
        Me.btnEValues.Name = "btnEValues"
        Me.btnEValues.Size = New System.Drawing.Size(123, 23)
        Me.btnEValues.TabIndex = 3
        Me.btnEValues.Text = "Create E Values"
        Me.btnEValues.UseVisualStyleBackColor = True
        '
        'rbStepByStep
        '
        Me.rbStepByStep.AutoSize = True
        Me.rbStepByStep.Location = New System.Drawing.Point(11, 20)
        Me.rbStepByStep.Name = "rbStepByStep"
        Me.rbStepByStep.Size = New System.Drawing.Size(109, 17)
        Me.rbStepByStep.TabIndex = 0
        Me.rbStepByStep.Text = "Step by Step Run"
        Me.rbStepByStep.UseVisualStyleBackColor = True
        '
        'btnReport
        '
        Me.btnReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnReport.Location = New System.Drawing.Point(8, 41)
        Me.btnReport.Name = "btnReport"
        Me.btnReport.Size = New System.Drawing.Size(123, 23)
        Me.btnReport.TabIndex = 2
        Me.btnReport.Text = "Create Output Report"
        Me.btnReport.UseVisualStyleBackColor = True
        '
        'ssStatus
        '
        Me.ssStatus.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ssStatus.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tssLabel})
        Me.ssStatus.Location = New System.Drawing.Point(0, 348)
        Me.ssStatus.Name = "ssStatus"
        Me.ssStatus.Size = New System.Drawing.Size(547, 22)
        Me.ssStatus.TabIndex = 2
        '
        'tssLabel
        '
        Me.tssLabel.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.tssLabel.Name = "tssLabel"
        Me.tssLabel.Size = New System.Drawing.Size(0, 17)
        '
        'txtFolder
        '
        Me.txtFolder.Location = New System.Drawing.Point(9, 48)
        Me.txtFolder.Name = "txtFolder"
        Me.txtFolder.ReadOnly = True
        Me.txtFolder.Size = New System.Drawing.Size(360, 21)
        Me.txtFolder.TabIndex = 3
        '
        'grpParameters
        '
        Me.grpParameters.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpParameters.Controls.Add(Me.cbLog)
        Me.grpParameters.Controls.Add(Me.cbAll)
        Me.grpParameters.Controls.Add(Me.clbRailroads)
        Me.grpParameters.Controls.Add(Me.lblYear)
        Me.grpParameters.Controls.Add(Me.txtFolder)
        Me.grpParameters.Controls.Add(Me.cmbYear)
        Me.grpParameters.Controls.Add(Me.btnFolder)
        Me.grpParameters.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.grpParameters.Location = New System.Drawing.Point(13, 12)
        Me.grpParameters.Name = "grpParameters"
        Me.grpParameters.Size = New System.Drawing.Size(525, 254)
        Me.grpParameters.TabIndex = 0
        Me.grpParameters.TabStop = False
        Me.grpParameters.Text = "Select Input Parameters"
        '
        'cbLog
        '
        Me.cbLog.AutoSize = True
        Me.cbLog.Location = New System.Drawing.Point(396, 22)
        Me.cbLog.Name = "cbLog"
        Me.cbLog.Size = New System.Drawing.Size(79, 17)
        Me.cbLog.TabIndex = 0
        Me.cbLog.Text = "Create Log"
        Me.cbLog.UseVisualStyleBackColor = True
        '
        'cbAll
        '
        Me.cbAll.AutoSize = True
        Me.cbAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAll.Location = New System.Drawing.Point(10, 75)
        Me.cbAll.Name = "cbAll"
        Me.cbAll.Size = New System.Drawing.Size(68, 17)
        Me.cbAll.TabIndex = 5
        Me.cbAll.Text = "Select all"
        Me.cbAll.UseVisualStyleBackColor = True
        '
        'clbRailroads
        '
        Me.clbRailroads.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.clbRailroads.CheckOnClick = True
        Me.clbRailroads.FormattingEnabled = True
        Me.clbRailroads.Location = New System.Drawing.Point(7, 95)
        Me.clbRailroads.Name = "clbRailroads"
        Me.clbRailroads.Size = New System.Drawing.Size(512, 148)
        Me.clbRailroads.TabIndex = 6
        '
        'lblYear
        '
        Me.lblYear.AutoSize = True
        Me.lblYear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblYear.Location = New System.Drawing.Point(10, 24)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(69, 13)
        Me.lblYear.TabIndex = 1
        Me.lblYear.Text = "Current year"
        '
        'cmbYear
        '
        Me.cmbYear.Location = New System.Drawing.Point(85, 20)
        Me.cmbYear.MaxLength = 4
        Me.cmbYear.Name = "cmbYear"
        Me.cmbYear.Size = New System.Drawing.Size(57, 21)
        Me.cmbYear.TabIndex = 2
        '
        'btnFolder
        '
        Me.btnFolder.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFolder.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnFolder.Location = New System.Drawing.Point(396, 46)
        Me.btnFolder.Name = "btnFolder"
        Me.btnFolder.Size = New System.Drawing.Size(123, 23)
        Me.btnFolder.TabIndex = 4
        Me.btnFolder.Text = "Select Output Folder"
        Me.btnFolder.UseVisualStyleBackColor = True
        '
        'frmPhase2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(547, 370)
        Me.Controls.Add(Me.grpParameters)
        Me.Controls.Add(Me.ssStatus)
        Me.Controls.Add(Me.grpRun)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmPhase2"
        Me.Text = "URCS Phase II Unit Costs"
        Me.grpRun.ResumeLayout(False)
        Me.grpRun.PerformLayout()
        Me.ssStatus.ResumeLayout(False)
        Me.ssStatus.PerformLayout()
        Me.grpParameters.ResumeLayout(False)
        Me.grpParameters.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grpRun As System.Windows.Forms.GroupBox
    Friend WithEvents btnReport As System.Windows.Forms.Button
    Friend WithEvents ssStatus As System.Windows.Forms.StatusStrip
    Friend WithEvents tssLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents txtFolder As System.Windows.Forms.TextBox
    Friend WithEvents btnEValues As System.Windows.Forms.Button
    Friend WithEvents grpParameters As System.Windows.Forms.GroupBox
    Friend WithEvents cbAll As System.Windows.Forms.CheckBox
    Friend WithEvents clbRailroads As System.Windows.Forms.CheckedListBox
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents cmbYear As System.Windows.Forms.ComboBox
    Friend WithEvents btnFolder As System.Windows.Forms.Button
    Friend WithEvents rbStepByStep As System.Windows.Forms.RadioButton
    Friend WithEvents btnLaunch As System.Windows.Forms.Button
    Friend WithEvents rbAllSteps As System.Windows.Forms.RadioButton
    Friend WithEvents cbLog As System.Windows.Forms.CheckBox

End Class
