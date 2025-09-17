<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WB_By_STCC_Generator
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
        Me.txt_Output_FilePath = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.chk_Unmask_STCC = New System.Windows.Forms.CheckBox()
        Me.chk_Unmasked = New System.Windows.Forms.CheckBox()
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.cmb_WB_Years = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_Return_To_WBGeneratorsMenu = New System.Windows.Forms.Button()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txt_STCC_Code = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btn_Output_File_Entry
        '
        Me.btn_Output_File_Entry.Location = New System.Drawing.Point(22, 200)
        Me.btn_Output_File_Entry.Name = "btn_Output_File_Entry"
        Me.btn_Output_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Output_File_Entry.TabIndex = 103
        Me.btn_Output_File_Entry.TabStop = False
        Me.btn_Output_File_Entry.Text = "Select Output File:"
        Me.btn_Output_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Output_FilePath
        '
        Me.txt_Output_FilePath.Location = New System.Drawing.Point(133, 200)
        Me.txt_Output_FilePath.Name = "txt_Output_FilePath"
        Me.txt_Output_FilePath.Size = New System.Drawing.Size(417, 21)
        Me.txt_Output_FilePath.TabIndex = 102
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(249, 253)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(108, 27)
        Me.btn_Execute.TabIndex = 101
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'chk_Unmask_STCC
        '
        Me.chk_Unmask_STCC.AutoSize = True
        Me.chk_Unmask_STCC.Location = New System.Drawing.Point(195, 177)
        Me.chk_Unmask_STCC.Name = "chk_Unmask_STCC"
        Me.chk_Unmask_STCC.Size = New System.Drawing.Size(175, 17)
        Me.chk_Unmask_STCC.TabIndex = 100
        Me.chk_Unmask_STCC.Text = "Unmask Series 19 STCC codes?"
        Me.chk_Unmask_STCC.UseVisualStyleBackColor = True
        '
        'chk_Unmasked
        '
        Me.chk_Unmasked.AutoSize = True
        Me.chk_Unmasked.Location = New System.Drawing.Point(195, 154)
        Me.chk_Unmasked.Name = "chk_Unmasked"
        Me.chk_Unmasked.Size = New System.Drawing.Size(214, 17)
        Me.chk_Unmasked.TabIndex = 99
        Me.chk_Unmasked.Text = "Produce Unmasked Data for this STCC?"
        Me.chk_Unmasked.UseVisualStyleBackColor = True
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(191, 70)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(65, 13)
        Me.lbl_Select_Year_Combobox.TabIndex = 98
        Me.lbl_Select_Year_Combobox.Text = "Select Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmb_WB_Years
        '
        Me.cmb_WB_Years.FormattingEnabled = True
        Me.cmb_WB_Years.Location = New System.Drawing.Point(262, 67)
        Me.cmb_WB_Years.Name = "cmb_WB_Years"
        Me.cmb_WB_Years.Size = New System.Drawing.Size(108, 21)
        Me.cmb_WB_Years.TabIndex = 97
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(171, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(247, 46)
        Me.Label1.TabIndex = 96
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Waybill By STCC Export Generator"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btn_Return_To_WBGeneratorsMenu
        '
        Me.btn_Return_To_WBGeneratorsMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_WBGeneratorsMenu.Location = New System.Drawing.Point(499, 281)
        Me.btn_Return_To_WBGeneratorsMenu.Name = "btn_Return_To_WBGeneratorsMenu"
        Me.btn_Return_To_WBGeneratorsMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_WBGeneratorsMenu.TabIndex = 104
        Me.btn_Return_To_WBGeneratorsMenu.UseVisualStyleBackColor = True
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(77, 231)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 105
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(45, 127)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(494, 13)
        Me.Label2.TabIndex = 106
        Me.Label2.Text = "Note:  You may use the leading digits as wildcard.  i.e. ""01"" would include any v" &
    "alues starting with 01."
        '
        'txt_STCC_Code
        '
        Me.txt_STCC_Code.Location = New System.Drawing.Point(262, 94)
        Me.txt_STCC_Code.Name = "txt_STCC_Code"
        Me.txt_STCC_Code.Size = New System.Drawing.Size(93, 21)
        Me.txt_STCC_Code.TabIndex = 107
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(191, 97)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 13)
        Me.Label3.TabIndex = 108
        Me.Label3.Text = "STCC Code:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'WB_By_STCC_Generator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(571, 349)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txt_STCC_Code)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Return_To_WBGeneratorsMenu)
        Me.Controls.Add(Me.btn_Output_File_Entry)
        Me.Controls.Add(Me.txt_Output_FilePath)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.chk_Unmask_STCC)
        Me.Controls.Add(Me.chk_Unmasked)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.cmb_WB_Years)
        Me.Controls.Add(Me.Label1)
        Me.Name = "WB_By_STCC_Generator"
        Me.Text = "Waybills By STCC Code Generator"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Return_To_WBGeneratorsMenu As System.Windows.Forms.Button
    Friend WithEvents btn_Output_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Output_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents chk_Unmask_STCC As System.Windows.Forms.CheckBox
    Friend WithEvents chk_Unmasked As System.Windows.Forms.CheckBox
    Friend WithEvents lbl_Select_Year_Combobox As System.Windows.Forms.Label
    Friend WithEvents cmb_WB_Years As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txt_STCC_Code As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
