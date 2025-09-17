<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class QuarterlyMaskingUtil
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btn_Output_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Output_FilePath = New System.Windows.Forms.TextBox()
        Me.btn_Text_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Input_Text_FilePath = New System.Windows.Forms.TextBox()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_Return_To_Menu = New System.Windows.Forms.Button()
        Me.rdo_Masked = New System.Windows.Forms.RadioButton()
        Me.rdo_Unmasked = New System.Windows.Forms.RadioButton()
        Me.rdo_Unmask_Some_49s_For_EIA = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.SuspendLayout()
        '
        'btn_Output_File_Entry
        '
        Me.btn_Output_File_Entry.Location = New System.Drawing.Point(30, 211)
        Me.btn_Output_File_Entry.Name = "btn_Output_File_Entry"
        Me.btn_Output_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Output_File_Entry.TabIndex = 99
        Me.btn_Output_File_Entry.TabStop = False
        Me.btn_Output_File_Entry.Text = "Select Output File:"
        Me.btn_Output_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Output_FilePath
        '
        Me.txt_Output_FilePath.Location = New System.Drawing.Point(141, 211)
        Me.txt_Output_FilePath.Name = "txt_Output_FilePath"
        Me.txt_Output_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Output_FilePath.TabIndex = 98
        '
        'btn_Text_File_Entry
        '
        Me.btn_Text_File_Entry.Location = New System.Drawing.Point(30, 71)
        Me.btn_Text_File_Entry.Name = "btn_Text_File_Entry"
        Me.btn_Text_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Text_File_Entry.TabIndex = 97
        Me.btn_Text_File_Entry.TabStop = False
        Me.btn_Text_File_Entry.Text = "Select Text File:"
        Me.btn_Text_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Text_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Input_Text_FilePath
        '
        Me.txt_Input_Text_FilePath.Location = New System.Drawing.Point(141, 71)
        Me.txt_Input_Text_FilePath.Name = "txt_Input_Text_FilePath"
        Me.txt_Input_Text_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Input_Text_FilePath.TabIndex = 96
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(57, 247)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 95
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(221, 270)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(104, 27)
        Me.btn_Execute.TabIndex = 94
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(145, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(284, 46)
        Me.Label1.TabIndex = 93
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Quarterly/Monthly Data Processing Utility" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btn_Return_To_Menu
        '
        Me.btn_Return_To_Menu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_Menu.Location = New System.Drawing.Point(479, 270)
        Me.btn_Return_To_Menu.Name = "btn_Return_To_Menu"
        Me.btn_Return_To_Menu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_Menu.TabIndex = 92
        Me.btn_Return_To_Menu.UseVisualStyleBackColor = True
        '
        'rdo_Masked
        '
        Me.rdo_Masked.AutoSize = True
        Me.rdo_Masked.Location = New System.Drawing.Point(195, 127)
        Me.rdo_Masked.Name = "rdo_Masked"
        Me.rdo_Masked.Size = New System.Drawing.Size(87, 17)
        Me.rdo_Masked.TabIndex = 101
        Me.rdo_Masked.TabStop = True
        Me.rdo_Masked.Text = "Masked Data"
        Me.rdo_Masked.UseVisualStyleBackColor = True
        '
        'rdo_Unmasked
        '
        Me.rdo_Unmasked.AutoSize = True
        Me.rdo_Unmasked.Location = New System.Drawing.Point(195, 150)
        Me.rdo_Unmasked.Name = "rdo_Unmasked"
        Me.rdo_Unmasked.Size = New System.Drawing.Size(100, 17)
        Me.rdo_Unmasked.TabIndex = 102
        Me.rdo_Unmasked.TabStop = True
        Me.rdo_Unmasked.Text = "Unmasked Data"
        Me.rdo_Unmasked.UseVisualStyleBackColor = True
        '
        'rdo_Unmask_Some_49s_For_EIA
        '
        Me.rdo_Unmask_Some_49s_For_EIA.AutoSize = True
        Me.rdo_Unmask_Some_49s_For_EIA.Location = New System.Drawing.Point(195, 173)
        Me.rdo_Unmask_Some_49s_For_EIA.Name = "rdo_Unmask_Some_49s_For_EIA"
        Me.rdo_Unmask_Some_49s_For_EIA.Size = New System.Drawing.Size(165, 17)
        Me.rdo_Unmask_Some_49s_For_EIA.TabIndex = 103
        Me.rdo_Unmask_Some_49s_For_EIA.TabStop = True
        Me.rdo_Unmask_Some_49s_For_EIA.Text = "Unmask Select STCCs for EIA"
        Me.rdo_Unmask_Some_49s_For_EIA.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(141, 108)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(251, 97)
        Me.GroupBox1.TabIndex = 104
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select One:"
        '
        'QuarterlyMaskingUtil
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(555, 341)
        Me.ControlBox = False
        Me.Controls.Add(Me.rdo_Unmask_Some_49s_For_EIA)
        Me.Controls.Add(Me.rdo_Unmasked)
        Me.Controls.Add(Me.rdo_Masked)
        Me.Controls.Add(Me.btn_Output_File_Entry)
        Me.Controls.Add(Me.txt_Output_FilePath)
        Me.Controls.Add(Me.btn_Text_File_Entry)
        Me.Controls.Add(Me.txt_Input_Text_FilePath)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn_Return_To_Menu)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "QuarterlyMaskingUtil"
        Me.Text = "Quarterly/Monthly Masking Util"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Output_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Output_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents btn_Text_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Input_Text_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_Return_To_Menu As System.Windows.Forms.Button
    Friend WithEvents rdo_Masked As RadioButton
    Friend WithEvents rdo_Unmasked As RadioButton
    Friend WithEvents rdo_Unmask_Some_49s_For_EIA As RadioButton
    Friend WithEvents GroupBox1 As GroupBox
End Class
