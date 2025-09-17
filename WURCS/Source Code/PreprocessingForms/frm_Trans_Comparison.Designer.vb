<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Trans_Comparison
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
        Me.btn_Output_File = New System.Windows.Forms.Button()
        Me.txt_Output_FilePath = New System.Windows.Forms.TextBox()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_Return_To_Menu = New System.Windows.Forms.Button()
        Me.lblBaseYear = New System.Windows.Forms.Label()
        Me.cmb_PreviousYear = New System.Windows.Forms.ComboBox()
        Me.lblComparisonYear = New System.Windows.Forms.Label()
        Me.cmb_CurrentYear = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'btn_Output_File
        '
        Me.btn_Output_File.Location = New System.Drawing.Point(18, 117)
        Me.btn_Output_File.Name = "btn_Output_File"
        Me.btn_Output_File.Size = New System.Drawing.Size(104, 21)
        Me.btn_Output_File.TabIndex = 99
        Me.btn_Output_File.TabStop = False
        Me.btn_Output_File.Text = "Select Output File:"
        Me.btn_Output_File.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_File.UseVisualStyleBackColor = True
        '
        'txt_Output_FilePath
        '
        Me.txt_Output_FilePath.Location = New System.Drawing.Point(129, 117)
        Me.txt_Output_FilePath.Name = "txt_Output_FilePath"
        Me.txt_Output_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Output_FilePath.TabIndex = 98
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(64, 147)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 97
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(220, 170)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(104, 27)
        Me.btn_Execute.TabIndex = 96
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(159, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(231, 46)
        Me.Label1.TabIndex = 95
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Trans Data Comparison Module"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btn_Return_To_Menu
        '
        Me.btn_Return_To_Menu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_Menu.Location = New System.Drawing.Point(468, 170)
        Me.btn_Return_To_Menu.Name = "btn_Return_To_Menu"
        Me.btn_Return_To_Menu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_Menu.TabIndex = 94
        Me.btn_Return_To_Menu.UseVisualStyleBackColor = True
        '
        'lblBaseYear
        '
        Me.lblBaseYear.AutoSize = True
        Me.lblBaseYear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBaseYear.Location = New System.Drawing.Point(203, 61)
        Me.lblBaseYear.Name = "lblBaseYear"
        Me.lblBaseYear.Size = New System.Drawing.Size(77, 13)
        Me.lblBaseYear.TabIndex = 100
        Me.lblBaseYear.Text = "Previous Year:"
        '
        'cmb_PreviousYear
        '
        Me.cmb_PreviousYear.Location = New System.Drawing.Point(286, 58)
        Me.cmb_PreviousYear.MaxLength = 4
        Me.cmb_PreviousYear.Name = "cmb_PreviousYear"
        Me.cmb_PreviousYear.Size = New System.Drawing.Size(60, 21)
        Me.cmb_PreviousYear.TabIndex = 101
        '
        'lblComparisonYear
        '
        Me.lblComparisonYear.AutoSize = True
        Me.lblComparisonYear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblComparisonYear.Location = New System.Drawing.Point(207, 88)
        Me.lblComparisonYear.Name = "lblComparisonYear"
        Me.lblComparisonYear.Size = New System.Drawing.Size(73, 13)
        Me.lblComparisonYear.TabIndex = 102
        Me.lblComparisonYear.Text = "Current Year:"
        '
        'cmb_CurrentYear
        '
        Me.cmb_CurrentYear.Location = New System.Drawing.Point(286, 85)
        Me.cmb_CurrentYear.MaxLength = 4
        Me.cmb_CurrentYear.Name = "cmb_CurrentYear"
        Me.cmb_CurrentYear.Size = New System.Drawing.Size(60, 21)
        Me.cmb_CurrentYear.TabIndex = 103
        '
        'frm_Trans_Comparison
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(540, 239)
        Me.Controls.Add(Me.lblComparisonYear)
        Me.Controls.Add(Me.cmb_CurrentYear)
        Me.Controls.Add(Me.lblBaseYear)
        Me.Controls.Add(Me.cmb_PreviousYear)
        Me.Controls.Add(Me.btn_Output_File)
        Me.Controls.Add(Me.txt_Output_FilePath)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn_Return_To_Menu)
        Me.Name = "frm_Trans_Comparison"
        Me.Text = "Trans Comparison Module"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Output_File As System.Windows.Forms.Button
    Friend WithEvents txt_Output_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_Return_To_Menu As System.Windows.Forms.Button
    Friend WithEvents lblBaseYear As System.Windows.Forms.Label
    Friend WithEvents cmb_PreviousYear As System.Windows.Forms.ComboBox
    Friend WithEvents lblComparisonYear As System.Windows.Forms.Label
    Friend WithEvents cmb_CurrentYear As System.Windows.Forms.ComboBox
End Class
