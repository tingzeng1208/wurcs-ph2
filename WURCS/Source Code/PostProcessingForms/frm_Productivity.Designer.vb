<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Productivity
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
        Me.btn_Return_To_MainMenu = New System.Windows.Forms.Button()
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.cmb_URCSYear = New System.Windows.Forms.ComboBox()
        Me.btn_Output_Dir_Entry = New System.Windows.Forms.Button()
        Me.txt_Output_FilePath = New System.Windows.Forms.TextBox()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.txt_Trailer_Usage_East = New System.Windows.Forms.TextBox()
        Me.lbl_Trailer_Usage_East = New System.Windows.Forms.Label()
        Me.lbl_Trailer_Usage_West = New System.Windows.Forms.Label()
        Me.txt_Trailer_Usage_West = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(192, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(166, 46)
        Me.Label1.TabIndex = 38
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Post Processing Menu"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btn_Return_To_MainMenu
        '
        Me.btn_Return_To_MainMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_MainMenu.Location = New System.Drawing.Point(492, 256)
        Me.btn_Return_To_MainMenu.Name = "btn_Return_To_MainMenu"
        Me.btn_Return_To_MainMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_MainMenu.TabIndex = 37
        Me.btn_Return_To_MainMenu.UseVisualStyleBackColor = True
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(185, 86)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(65, 13)
        Me.lbl_Select_Year_Combobox.TabIndex = 89
        Me.lbl_Select_Year_Combobox.Text = "Select Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmb_URCSYear
        '
        Me.cmb_URCSYear.FormattingEnabled = True
        Me.cmb_URCSYear.Location = New System.Drawing.Point(256, 83)
        Me.cmb_URCSYear.Name = "cmb_URCSYear"
        Me.cmb_URCSYear.Size = New System.Drawing.Size(108, 21)
        Me.cmb_URCSYear.TabIndex = 88
        '
        'btn_Output_Dir_Entry
        '
        Me.btn_Output_Dir_Entry.Location = New System.Drawing.Point(26, 164)
        Me.btn_Output_Dir_Entry.Name = "btn_Output_Dir_Entry"
        Me.btn_Output_Dir_Entry.Size = New System.Drawing.Size(141, 21)
        Me.btn_Output_Dir_Entry.TabIndex = 87
        Me.btn_Output_Dir_Entry.TabStop = False
        Me.btn_Output_Dir_Entry.Text = "Select Output Directory:"
        Me.btn_Output_Dir_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_Dir_Entry.UseVisualStyleBackColor = True
        '
        'txt_Output_FilePath
        '
        Me.txt_Output_FilePath.Location = New System.Drawing.Point(173, 164)
        Me.txt_Output_FilePath.Name = "txt_Output_FilePath"
        Me.txt_Output_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Output_FilePath.TabIndex = 86
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(76, 197)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 85
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(232, 220)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(104, 27)
        Me.btn_Execute.TabIndex = 84
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'txt_Trailer_Usage_East
        '
        Me.txt_Trailer_Usage_East.Enabled = False
        Me.txt_Trailer_Usage_East.Location = New System.Drawing.Point(322, 110)
        Me.txt_Trailer_Usage_East.Name = "txt_Trailer_Usage_East"
        Me.txt_Trailer_Usage_East.Size = New System.Drawing.Size(69, 21)
        Me.txt_Trailer_Usage_East.TabIndex = 90
        '
        'lbl_Trailer_Usage_East
        '
        Me.lbl_Trailer_Usage_East.AutoSize = True
        Me.lbl_Trailer_Usage_East.Location = New System.Drawing.Point(176, 113)
        Me.lbl_Trailer_Usage_East.Name = "lbl_Trailer_Usage_East"
        Me.lbl_Trailer_Usage_East.Size = New System.Drawing.Size(137, 13)
        Me.lbl_Trailer_Usage_East.TabIndex = 91
        Me.lbl_Trailer_Usage_East.Text = "TOFC Usage Factor (East):"
        '
        'lbl_Trailer_Usage_West
        '
        Me.lbl_Trailer_Usage_West.AutoSize = True
        Me.lbl_Trailer_Usage_West.Location = New System.Drawing.Point(176, 140)
        Me.lbl_Trailer_Usage_West.Name = "lbl_Trailer_Usage_West"
        Me.lbl_Trailer_Usage_West.Size = New System.Drawing.Size(141, 13)
        Me.lbl_Trailer_Usage_West.TabIndex = 93
        Me.lbl_Trailer_Usage_West.Text = "TOFC Usage Factor (West):"
        '
        'txt_Trailer_Usage_West
        '
        Me.txt_Trailer_Usage_West.Enabled = False
        Me.txt_Trailer_Usage_West.Location = New System.Drawing.Point(322, 137)
        Me.txt_Trailer_Usage_West.Name = "txt_Trailer_Usage_West"
        Me.txt_Trailer_Usage_West.Size = New System.Drawing.Size(69, 21)
        Me.txt_Trailer_Usage_West.TabIndex = 92
        '
        'frm_Productivity
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(566, 325)
        Me.Controls.Add(Me.lbl_Trailer_Usage_West)
        Me.Controls.Add(Me.txt_Trailer_Usage_West)
        Me.Controls.Add(Me.lbl_Trailer_Usage_East)
        Me.Controls.Add(Me.txt_Trailer_Usage_East)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.cmb_URCSYear)
        Me.Controls.Add(Me.btn_Output_Dir_Entry)
        Me.Controls.Add(Me.txt_Output_FilePath)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn_Return_To_MainMenu)
        Me.Name = "frm_Productivity"
        Me.Text = "Productivity"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_Return_To_MainMenu As System.Windows.Forms.Button
    Friend WithEvents lbl_Select_Year_Combobox As System.Windows.Forms.Label
    Friend WithEvents cmb_URCSYear As System.Windows.Forms.ComboBox
    Friend WithEvents btn_Output_Dir_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Output_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents txt_Trailer_Usage_East As System.Windows.Forms.TextBox
    Friend WithEvents lbl_Trailer_Usage_East As System.Windows.Forms.Label
    Friend WithEvents lbl_Trailer_Usage_West As System.Windows.Forms.Label
    Friend WithEvents txt_Trailer_Usage_West As System.Windows.Forms.TextBox
End Class
