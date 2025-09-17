<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Annual_WB_Generator
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
        Me.btn_Output_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Output_DirPath = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.cmb_WB_Year = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.opt_RecordLayoutSelector = New System.Windows.Forms.GroupBox()
        Me.RadioButton_570 = New System.Windows.Forms.RadioButton()
        Me.RadioButton_913 = New System.Windows.Forms.RadioButton()
        Me.UnmaskedDataCheck = New System.Windows.Forms.CheckBox()
        Me.UnmaskedSTCCCheck = New System.Windows.Forms.CheckBox()
        Me.MaskedRevenueAndLocationsCheck = New System.Windows.Forms.CheckBox()
        Me.btn_Return_To_WBGeneratorsMenu = New System.Windows.Forms.Button()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.chk_ZeroRevenues = New System.Windows.Forms.CheckBox()
        Me.opt_RecordLayoutSelector.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_Output_File_Entry
        '
        Me.btn_Output_File_Entry.Location = New System.Drawing.Point(25, 227)
        Me.btn_Output_File_Entry.Name = "btn_Output_File_Entry"
        Me.btn_Output_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Output_File_Entry.TabIndex = 55
        Me.btn_Output_File_Entry.TabStop = False
        Me.btn_Output_File_Entry.Text = "Select Output File:"
        Me.btn_Output_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Output_DirPath
        '
        Me.txt_Output_DirPath.Location = New System.Drawing.Point(136, 227)
        Me.txt_Output_DirPath.Name = "txt_Output_DirPath"
        Me.txt_Output_DirPath.Size = New System.Drawing.Size(401, 21)
        Me.txt_Output_DirPath.TabIndex = 54
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(227, 285)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(108, 27)
        Me.btn_Execute.TabIndex = 53
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(172, 89)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(65, 13)
        Me.lbl_Select_Year_Combobox.TabIndex = 52
        Me.lbl_Select_Year_Combobox.Text = "Select Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmb_WB_Year
        '
        Me.cmb_WB_Year.FormattingEnabled = True
        Me.cmb_WB_Year.Location = New System.Drawing.Point(243, 86)
        Me.cmb_WB_Year.Name = "cmb_WB_Year"
        Me.cmb_WB_Year.Size = New System.Drawing.Size(108, 21)
        Me.cmb_WB_Year.TabIndex = 51
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(170, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(234, 46)
        Me.Label1.TabIndex = 50
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Annual Waybill Export Generator"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'opt_RecordLayoutSelector
        '
        Me.opt_RecordLayoutSelector.Controls.Add(Me.RadioButton_570)
        Me.opt_RecordLayoutSelector.Controls.Add(Me.RadioButton_913)
        Me.opt_RecordLayoutSelector.Location = New System.Drawing.Point(73, 118)
        Me.opt_RecordLayoutSelector.Name = "opt_RecordLayoutSelector"
        Me.opt_RecordLayoutSelector.Size = New System.Drawing.Size(203, 92)
        Me.opt_RecordLayoutSelector.TabIndex = 56
        Me.opt_RecordLayoutSelector.TabStop = False
        Me.opt_RecordLayoutSelector.Text = "Record Format:"
        '
        'RadioButton_570
        '
        Me.RadioButton_570.AutoSize = True
        Me.RadioButton_570.ForeColor = System.Drawing.Color.Black
        Me.RadioButton_570.Location = New System.Drawing.Point(17, 43)
        Me.RadioButton_570.Name = "RadioButton_570"
        Me.RadioButton_570.Size = New System.Drawing.Size(152, 17)
        Me.RadioButton_570.TabIndex = 1
        Me.RadioButton_570.TabStop = True
        Me.RadioButton_570.Text = "570 Byte (Legacy Costing)"
        Me.RadioButton_570.UseVisualStyleBackColor = True
        '
        'RadioButton_913
        '
        Me.RadioButton_913.AutoSize = True
        Me.RadioButton_913.Location = New System.Drawing.Point(17, 20)
        Me.RadioButton_913.Name = "RadioButton_913"
        Me.RadioButton_913.Size = New System.Drawing.Size(68, 17)
        Me.RadioButton_913.TabIndex = 0
        Me.RadioButton_913.TabStop = True
        Me.RadioButton_913.Text = "913 Byte"
        Me.RadioButton_913.UseVisualStyleBackColor = True
        '
        'UnmaskedDataCheck
        '
        Me.UnmaskedDataCheck.AutoSize = True
        Me.UnmaskedDataCheck.Location = New System.Drawing.Point(293, 123)
        Me.UnmaskedDataCheck.Name = "UnmaskedDataCheck"
        Me.UnmaskedDataCheck.Size = New System.Drawing.Size(194, 17)
        Me.UnmaskedDataCheck.TabIndex = 57
        Me.UnmaskedDataCheck.Text = "Produce Unmasked Revenue Data?"
        Me.UnmaskedDataCheck.UseVisualStyleBackColor = True
        '
        'UnmaskedSTCCCheck
        '
        Me.UnmaskedSTCCCheck.AutoSize = True
        Me.UnmaskedSTCCCheck.Location = New System.Drawing.Point(293, 146)
        Me.UnmaskedSTCCCheck.Name = "UnmaskedSTCCCheck"
        Me.UnmaskedSTCCCheck.Size = New System.Drawing.Size(182, 17)
        Me.UnmaskedSTCCCheck.TabIndex = 58
        Me.UnmaskedSTCCCheck.Text = "Unmask Ordinance STCC Codes?"
        Me.UnmaskedSTCCCheck.UseVisualStyleBackColor = True
        '
        'MaskedRevenueAndLocationsCheck
        '
        Me.MaskedRevenueAndLocationsCheck.AutoSize = True
        Me.MaskedRevenueAndLocationsCheck.Location = New System.Drawing.Point(293, 169)
        Me.MaskedRevenueAndLocationsCheck.Name = "MaskedRevenueAndLocationsCheck"
        Me.MaskedRevenueAndLocationsCheck.Size = New System.Drawing.Size(170, 17)
        Me.MaskedRevenueAndLocationsCheck.TabIndex = 60
        Me.MaskedRevenueAndLocationsCheck.Text = "Mask Revenue and Locations?"
        Me.MaskedRevenueAndLocationsCheck.UseVisualStyleBackColor = True
        '
        'btn_Return_To_WBGeneratorsMenu
        '
        Me.btn_Return_To_WBGeneratorsMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_WBGeneratorsMenu.Location = New System.Drawing.Point(510, 307)
        Me.btn_Return_To_WBGeneratorsMenu.Name = "btn_Return_To_WBGeneratorsMenu"
        Me.btn_Return_To_WBGeneratorsMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_WBGeneratorsMenu.TabIndex = 61
        Me.btn_Return_To_WBGeneratorsMenu.UseVisualStyleBackColor = True
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(71, 262)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 67
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'chk_ZeroRevenues
        '
        Me.chk_ZeroRevenues.AutoSize = True
        Me.chk_ZeroRevenues.Location = New System.Drawing.Point(293, 192)
        Me.chk_ZeroRevenues.Name = "chk_ZeroRevenues"
        Me.chk_ZeroRevenues.Size = New System.Drawing.Size(94, 17)
        Me.chk_ZeroRevenues.TabIndex = 68
        Me.chk_ZeroRevenues.Text = "Zero Revenue"
        Me.chk_ZeroRevenues.UseVisualStyleBackColor = True
        '
        'Annual_WB_Generator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(573, 376)
        Me.ControlBox = False
        Me.Controls.Add(Me.chk_ZeroRevenues)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Return_To_WBGeneratorsMenu)
        Me.Controls.Add(Me.MaskedRevenueAndLocationsCheck)
        Me.Controls.Add(Me.UnmaskedSTCCCheck)
        Me.Controls.Add(Me.UnmaskedDataCheck)
        Me.Controls.Add(Me.opt_RecordLayoutSelector)
        Me.Controls.Add(Me.btn_Output_File_Entry)
        Me.Controls.Add(Me.txt_Output_DirPath)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.cmb_WB_Year)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Annual_WB_Generator"
        Me.Text = "Annual Waybills Export Generator"
        Me.opt_RecordLayoutSelector.ResumeLayout(False)
        Me.opt_RecordLayoutSelector.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Output_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Output_DirPath As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents lbl_Select_Year_Combobox As System.Windows.Forms.Label
    Friend WithEvents cmb_WB_Year As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents opt_RecordLayoutSelector As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton_570 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_913 As System.Windows.Forms.RadioButton
    Friend WithEvents UnmaskedDataCheck As System.Windows.Forms.CheckBox
    Friend WithEvents UnmaskedSTCCCheck As System.Windows.Forms.CheckBox
    Friend WithEvents MaskedRevenueAndLocationsCheck As System.Windows.Forms.CheckBox
    Friend WithEvents btn_Return_To_WBGeneratorsMenu As System.Windows.Forms.Button
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents chk_ZeroRevenues As CheckBox
End Class
