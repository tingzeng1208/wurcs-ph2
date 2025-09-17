<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class State_WB_Generator
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
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.opt_RecordLayoutSelector = New System.Windows.Forms.GroupBox()
        Me.RadioButton_CSV = New System.Windows.Forms.RadioButton()
        Me.RadioButton_913 = New System.Windows.Forms.RadioButton()
        Me.btn_Output_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Output_FilePath = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.cmb_WB_Year = New System.Windows.Forms.ComboBox()
        Me.btn_Return_To_WBGeneratorsMenu = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CheckBox_Unmasked = New System.Windows.Forms.CheckBox()
        Me.lbx_States = New System.Windows.Forms.ListBox()
        Me.opt_RecordLayoutSelector.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(208, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(281, 58)
        Me.Label1.TabIndex = 51
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "State Waybill Export Generator"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(85, 471)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(7, 7, 7, 7)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(504, 16)
        Me.txt_StatusBox.TabIndex = 74
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'opt_RecordLayoutSelector
        '
        Me.opt_RecordLayoutSelector.Controls.Add(Me.RadioButton_CSV)
        Me.opt_RecordLayoutSelector.Controls.Add(Me.RadioButton_913)
        Me.opt_RecordLayoutSelector.Location = New System.Drawing.Point(283, 315)
        Me.opt_RecordLayoutSelector.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.opt_RecordLayoutSelector.Name = "opt_RecordLayoutSelector"
        Me.opt_RecordLayoutSelector.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.opt_RecordLayoutSelector.Size = New System.Drawing.Size(125, 84)
        Me.opt_RecordLayoutSelector.TabIndex = 73
        Me.opt_RecordLayoutSelector.TabStop = False
        Me.opt_RecordLayoutSelector.Text = "Record Format:"
        '
        'RadioButton_CSV
        '
        Me.RadioButton_CSV.AutoSize = True
        Me.RadioButton_CSV.Location = New System.Drawing.Point(20, 53)
        Me.RadioButton_CSV.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.RadioButton_CSV.Name = "RadioButton_CSV"
        Me.RadioButton_CSV.Size = New System.Drawing.Size(54, 21)
        Me.RadioButton_CSV.TabIndex = 2
        Me.RadioButton_CSV.TabStop = True
        Me.RadioButton_CSV.Text = "CSV"
        Me.RadioButton_CSV.UseVisualStyleBackColor = True
        '
        'RadioButton_913
        '
        Me.RadioButton_913.AutoSize = True
        Me.RadioButton_913.Location = New System.Drawing.Point(20, 25)
        Me.RadioButton_913.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.RadioButton_913.Name = "RadioButton_913"
        Me.RadioButton_913.Size = New System.Drawing.Size(85, 21)
        Me.RadioButton_913.TabIndex = 0
        Me.RadioButton_913.TabStop = True
        Me.RadioButton_913.Text = "913 Byte"
        Me.RadioButton_913.UseVisualStyleBackColor = True
        '
        'btn_Output_File_Entry
        '
        Me.btn_Output_File_Entry.Location = New System.Drawing.Point(19, 434)
        Me.btn_Output_File_Entry.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btn_Output_File_Entry.Name = "btn_Output_File_Entry"
        Me.btn_Output_File_Entry.Size = New System.Drawing.Size(121, 26)
        Me.btn_Output_File_Entry.TabIndex = 72
        Me.btn_Output_File_Entry.TabStop = False
        Me.btn_Output_File_Entry.Text = "Select Output File:"
        Me.btn_Output_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Output_FilePath
        '
        Me.txt_Output_FilePath.Location = New System.Drawing.Point(148, 434)
        Me.txt_Output_FilePath.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_Output_FilePath.Name = "txt_Output_FilePath"
        Me.txt_Output_FilePath.Size = New System.Drawing.Size(486, 23)
        Me.txt_Output_FilePath.TabIndex = 71
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(283, 500)
        Me.btn_Execute.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(126, 33)
        Me.btn_Execute.TabIndex = 70
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(196, 86)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(80, 17)
        Me.lbl_Select_Year_Combobox.TabIndex = 69
        Me.lbl_Select_Year_Combobox.Text = "Select Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmb_WB_Year
        '
        Me.cmb_WB_Year.FormattingEnabled = True
        Me.cmb_WB_Year.Location = New System.Drawing.Point(279, 82)
        Me.cmb_WB_Year.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmb_WB_Year.Name = "cmb_WB_Year"
        Me.cmb_WB_Year.Size = New System.Drawing.Size(125, 24)
        Me.cmb_WB_Year.TabIndex = 68
        '
        'btn_Return_To_WBGeneratorsMenu
        '
        Me.btn_Return_To_WBGeneratorsMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_WBGeneratorsMenu.Location = New System.Drawing.Point(575, 548)
        Me.btn_Return_To_WBGeneratorsMenu.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btn_Return_To_WBGeneratorsMenu.Name = "btn_Return_To_WBGeneratorsMenu"
        Me.btn_Return_To_WBGeneratorsMenu.Size = New System.Drawing.Size(59, 64)
        Me.btn_Return_To_WBGeneratorsMenu.TabIndex = 75
        Me.btn_Return_To_WBGeneratorsMenu.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(177, 116)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(101, 17)
        Me.Label3.TabIndex = 81
        Me.Label3.Text = "Select State(s):"
        '
        'CheckBox_Unmasked
        '
        Me.CheckBox_Unmasked.AutoSize = True
        Me.CheckBox_Unmasked.Location = New System.Drawing.Point(258, 406)
        Me.CheckBox_Unmasked.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CheckBox_Unmasked.Name = "CheckBox_Unmasked"
        Me.CheckBox_Unmasked.Size = New System.Drawing.Size(185, 21)
        Me.CheckBox_Unmasked.TabIndex = 84
        Me.CheckBox_Unmasked.Text = "Unmasked Revenue Data"
        Me.CheckBox_Unmasked.UseVisualStyleBackColor = True
        '
        'lbx_States
        '
        Me.lbx_States.FormattingEnabled = True
        Me.lbx_States.ItemHeight = 16
        Me.lbx_States.Location = New System.Drawing.Point(279, 116)
        Me.lbx_States.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lbx_States.Name = "lbx_States"
        Me.lbx_States.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lbx_States.Size = New System.Drawing.Size(100, 196)
        Me.lbx_States.TabIndex = 86
        '
        'State_WB_Generator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(649, 634)
        Me.Controls.Add(Me.lbx_States)
        Me.Controls.Add(Me.CheckBox_Unmasked)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btn_Return_To_WBGeneratorsMenu)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.opt_RecordLayoutSelector)
        Me.Controls.Add(Me.btn_Output_File_Entry)
        Me.Controls.Add(Me.txt_Output_FilePath)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.cmb_WB_Year)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "State_WB_Generator"
        Me.Text = "Waybills By State Generator"
        Me.opt_RecordLayoutSelector.ResumeLayout(False)
        Me.opt_RecordLayoutSelector.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents opt_RecordLayoutSelector As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton_CSV As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_913 As System.Windows.Forms.RadioButton
    Friend WithEvents btn_Output_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Output_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents lbl_Select_Year_Combobox As System.Windows.Forms.Label
    Friend WithEvents cmb_WB_Year As System.Windows.Forms.ComboBox
    Friend WithEvents btn_Return_To_WBGeneratorsMenu As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CheckBox_Unmasked As System.Windows.Forms.CheckBox
    Friend WithEvents lbx_States As ListBox
End Class
