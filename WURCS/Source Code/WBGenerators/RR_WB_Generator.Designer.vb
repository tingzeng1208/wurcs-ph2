<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RR_WB_Generator
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
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.cmb_URCS_Year = New System.Windows.Forms.ComboBox()
        Me.cmb_Railroad = New System.Windows.Forms.ComboBox()
        Me.chk_Unmask_STCC = New System.Windows.Forms.CheckBox()
        Me.chk_Litigation = New System.Windows.Forms.CheckBox()
        Me.chk_Unmasked = New System.Windows.Forms.CheckBox()
        Me.btn_Return_To_WBGeneratorsMenu = New System.Windows.Forms.Button()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Output_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Output_FilePath = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.chk_RedactOthers = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(180, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(243, 46)
        Me.Label1.TabIndex = 52
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Railroad Waybill Export Generator"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(47, 102)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 13)
        Me.Label3.TabIndex = 86
        Me.Label3.Text = "Select RR:"
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(200, 70)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(65, 13)
        Me.lbl_Select_Year_Combobox.TabIndex = 85
        Me.lbl_Select_Year_Combobox.Text = "Select Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmb_URCS_Year
        '
        Me.cmb_URCS_Year.FormattingEnabled = True
        Me.cmb_URCS_Year.Location = New System.Drawing.Point(271, 67)
        Me.cmb_URCS_Year.Name = "cmb_URCS_Year"
        Me.cmb_URCS_Year.Size = New System.Drawing.Size(108, 21)
        Me.cmb_URCS_Year.TabIndex = 84
        '
        'cmb_Railroad
        '
        Me.cmb_Railroad.FormattingEnabled = True
        Me.cmb_Railroad.Location = New System.Drawing.Point(110, 99)
        Me.cmb_Railroad.Name = "cmb_Railroad"
        Me.cmb_Railroad.Size = New System.Drawing.Size(419, 21)
        Me.cmb_Railroad.TabIndex = 87
        '
        'chk_Unmask_STCC
        '
        Me.chk_Unmask_STCC.AutoSize = True
        Me.chk_Unmask_STCC.Location = New System.Drawing.Point(203, 195)
        Me.chk_Unmask_STCC.Name = "chk_Unmask_STCC"
        Me.chk_Unmask_STCC.Size = New System.Drawing.Size(175, 17)
        Me.chk_Unmask_STCC.TabIndex = 90
        Me.chk_Unmask_STCC.Text = "Unmask Series 19 STCC codes?"
        Me.chk_Unmask_STCC.UseVisualStyleBackColor = True
        '
        'chk_Litigation
        '
        Me.chk_Litigation.AutoSize = True
        Me.chk_Litigation.Location = New System.Drawing.Point(203, 172)
        Me.chk_Litigation.Name = "chk_Litigation"
        Me.chk_Litigation.Size = New System.Drawing.Size(220, 17)
        Me.chk_Litigation.TabIndex = 89
        Me.chk_Litigation.Text = "Unmask all data for all roads (Litigation)?"
        Me.chk_Litigation.UseVisualStyleBackColor = True
        '
        'chk_Unmasked
        '
        Me.chk_Unmasked.AutoSize = True
        Me.chk_Unmasked.Location = New System.Drawing.Point(203, 126)
        Me.chk_Unmasked.Name = "chk_Unmasked"
        Me.chk_Unmasked.Size = New System.Drawing.Size(202, 17)
        Me.chk_Unmasked.TabIndex = 88
        Me.chk_Unmasked.Text = "Produce Unmasked Data for this RR?"
        Me.chk_Unmasked.UseVisualStyleBackColor = True
        '
        'btn_Return_To_WBGeneratorsMenu
        '
        Me.btn_Return_To_WBGeneratorsMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_WBGeneratorsMenu.Location = New System.Drawing.Point(507, 299)
        Me.btn_Return_To_WBGeneratorsMenu.Name = "btn_Return_To_WBGeneratorsMenu"
        Me.btn_Return_To_WBGeneratorsMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_WBGeneratorsMenu.TabIndex = 95
        Me.btn_Return_To_WBGeneratorsMenu.UseVisualStyleBackColor = True
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(87, 248)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 94
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Output_File_Entry
        '
        Me.btn_Output_File_Entry.Location = New System.Drawing.Point(30, 218)
        Me.btn_Output_File_Entry.Name = "btn_Output_File_Entry"
        Me.btn_Output_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Output_File_Entry.TabIndex = 93
        Me.btn_Output_File_Entry.TabStop = False
        Me.btn_Output_File_Entry.Text = "Select Output File:"
        Me.btn_Output_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Output_FilePath
        '
        Me.txt_Output_FilePath.Location = New System.Drawing.Point(141, 218)
        Me.txt_Output_FilePath.Name = "txt_Output_FilePath"
        Me.txt_Output_FilePath.Size = New System.Drawing.Size(417, 21)
        Me.txt_Output_FilePath.TabIndex = 92
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(257, 271)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(108, 27)
        Me.btn_Execute.TabIndex = 91
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'chk_RedactOthers
        '
        Me.chk_RedactOthers.AutoSize = True
        Me.chk_RedactOthers.Location = New System.Drawing.Point(229, 149)
        Me.chk_RedactOthers.Name = "chk_RedactOthers"
        Me.chk_RedactOthers.Size = New System.Drawing.Size(150, 17)
        Me.chk_RedactOthers.TabIndex = 96
        Me.chk_RedactOthers.Text = "Redact All Other RR Info?"
        Me.chk_RedactOthers.UseVisualStyleBackColor = True
        '
        'RR_WB_Generator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(579, 387)
        Me.ControlBox = False
        Me.Controls.Add(Me.chk_RedactOthers)
        Me.Controls.Add(Me.btn_Return_To_WBGeneratorsMenu)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Output_File_Entry)
        Me.Controls.Add(Me.txt_Output_FilePath)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.chk_Unmask_STCC)
        Me.Controls.Add(Me.chk_Litigation)
        Me.Controls.Add(Me.chk_Unmasked)
        Me.Controls.Add(Me.cmb_Railroad)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.cmb_URCS_Year)
        Me.Controls.Add(Me.Label1)
        Me.Name = "RR_WB_Generator"
        Me.Text = "Railroad Waybill Export Generator"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lbl_Select_Year_Combobox As System.Windows.Forms.Label
    Friend WithEvents cmb_URCS_Year As System.Windows.Forms.ComboBox
    Friend WithEvents cmb_Railroad As System.Windows.Forms.ComboBox
    Friend WithEvents chk_Unmask_STCC As System.Windows.Forms.CheckBox
    Friend WithEvents chk_Litigation As System.Windows.Forms.CheckBox
    Friend WithEvents chk_Unmasked As System.Windows.Forms.CheckBox
    Friend WithEvents btn_Return_To_WBGeneratorsMenu As System.Windows.Forms.Button
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents btn_Output_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Output_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents chk_RedactOthers As CheckBox
End Class
