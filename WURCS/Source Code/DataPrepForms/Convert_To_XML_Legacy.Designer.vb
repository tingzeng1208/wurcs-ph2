<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Convert_To_XML_Legacy
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
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.cmb_URCS_Year = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Return_To_DataPrepMenu = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btn_Output_File_Entry
        '
        Me.btn_Output_File_Entry.Location = New System.Drawing.Point(11, 107)
        Me.btn_Output_File_Entry.Name = "btn_Output_File_Entry"
        Me.btn_Output_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Output_File_Entry.TabIndex = 55
        Me.btn_Output_File_Entry.TabStop = False
        Me.btn_Output_File_Entry.Text = "Select Output File:"
        Me.btn_Output_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Output_FilePath
        '
        Me.txt_Output_FilePath.Location = New System.Drawing.Point(122, 107)
        Me.txt_Output_FilePath.Name = "txt_Output_FilePath"
        Me.txt_Output_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Output_FilePath.TabIndex = 54
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(196, 173)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(108, 27)
        Me.btn_Execute.TabIndex = 53
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(160, 71)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(65, 13)
        Me.lbl_Select_Year_Combobox.TabIndex = 52
        Me.lbl_Select_Year_Combobox.Text = "Select Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmb_URCS_Year
        '
        Me.cmb_URCS_Year.FormattingEnabled = True
        Me.cmb_URCS_Year.Location = New System.Drawing.Point(231, 68)
        Me.cmb_URCS_Year.Name = "cmb_URCS_Year"
        Me.cmb_URCS_Year.Size = New System.Drawing.Size(108, 21)
        Me.cmb_URCS_Year.TabIndex = 51
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(155, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(195, 46)
        Me.Label1.TabIndex = 50
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "URCS Tare Weight Loader"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.Enabled = False
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(42, 150)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(418, 14)
        Me.txt_StatusBox.TabIndex = 57
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Return_To_DataPrepMenu
        '
        Me.btn_Return_To_DataPrepMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_DataPrepMenu.Location = New System.Drawing.Point(436, 198)
        Me.btn_Return_To_DataPrepMenu.Name = "btn_Return_To_DataPrepMenu"
        Me.btn_Return_To_DataPrepMenu.Size = New System.Drawing.Size(55, 52)
        Me.btn_Return_To_DataPrepMenu.TabIndex = 56
        Me.btn_Return_To_DataPrepMenu.UseVisualStyleBackColor = True
        '
        'Convert_To_XML_Legacy
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(504, 267)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Return_To_DataPrepMenu)
        Me.Controls.Add(Me.btn_Output_File_Entry)
        Me.Controls.Add(Me.txt_Output_FilePath)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.cmb_URCS_Year)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Convert_To_XML_Legacy"
        Me.Text = "Convert_To_XML_Legacy"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Output_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Output_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents lbl_Select_Year_Combobox As System.Windows.Forms.Label
    Friend WithEvents cmb_URCS_Year As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents btn_Return_To_DataPrepMenu As System.Windows.Forms.Button
End Class
