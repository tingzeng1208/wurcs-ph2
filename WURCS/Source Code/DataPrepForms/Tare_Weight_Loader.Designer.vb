<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Tare_Weight_Loader
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
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.btn_Return_To_DataPrepMenu = New System.Windows.Forms.Button()
        Me.txt_Input_FilePath = New System.Windows.Forms.TextBox()
        Me.txt_Report_FilePath = New System.Windows.Forms.TextBox()
        Me.btn_Report_File_Entry = New System.Windows.Forms.Button()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Input_File_Entry = New System.Windows.Forms.Button()
        Me.OpenInputFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.OpenReportFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.cmb_URCSYear = New System.Windows.Forms.ComboBox()
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(186, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(195, 46)
        Me.Label1.TabIndex = 33
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "URCS Tare Weight Loader"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(232, 178)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(104, 27)
        Me.btn_Execute.TabIndex = 34
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'btn_Return_To_DataPrepMenu
        '
        Me.btn_Return_To_DataPrepMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_DataPrepMenu.Location = New System.Drawing.Point(481, 215)
        Me.btn_Return_To_DataPrepMenu.Name = "btn_Return_To_DataPrepMenu"
        Me.btn_Return_To_DataPrepMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_DataPrepMenu.TabIndex = 35
        Me.btn_Return_To_DataPrepMenu.UseVisualStyleBackColor = True
        '
        'txt_Input_FilePath
        '
        Me.txt_Input_FilePath.Location = New System.Drawing.Point(150, 98)
        Me.txt_Input_FilePath.Name = "txt_Input_FilePath"
        Me.txt_Input_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Input_FilePath.TabIndex = 36
        '
        'txt_Report_FilePath
        '
        Me.txt_Report_FilePath.Location = New System.Drawing.Point(150, 125)
        Me.txt_Report_FilePath.Name = "txt_Report_FilePath"
        Me.txt_Report_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Report_FilePath.TabIndex = 37
        '
        'btn_Report_File_Entry
        '
        Me.btn_Report_File_Entry.Location = New System.Drawing.Point(39, 125)
        Me.btn_Report_File_Entry.Name = "btn_Report_File_Entry"
        Me.btn_Report_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Report_File_Entry.TabIndex = 39
        Me.btn_Report_File_Entry.TabStop = False
        Me.btn_Report_File_Entry.Text = "Select Report File:"
        Me.btn_Report_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Report_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(64, 155)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 40
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Input_File_Entry
        '
        Me.btn_Input_File_Entry.Location = New System.Drawing.Point(39, 98)
        Me.btn_Input_File_Entry.Name = "btn_Input_File_Entry"
        Me.btn_Input_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Input_File_Entry.TabIndex = 38
        Me.btn_Input_File_Entry.TabStop = False
        Me.btn_Input_File_Entry.Text = "Select Input File:"
        Me.btn_Input_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Input_File_Entry.UseVisualStyleBackColor = True
        '
        'OpenReportFileDialog
        '
        Me.OpenReportFileDialog.CheckFileExists = False
        Me.OpenReportFileDialog.DefaultExt = "rtf"
        '
        'cmb_URCSYear
        '
        Me.cmb_URCSYear.FormattingEnabled = True
        Me.cmb_URCSYear.Location = New System.Drawing.Point(232, 71)
        Me.cmb_URCSYear.Name = "cmb_URCSYear"
        Me.cmb_URCSYear.Size = New System.Drawing.Size(108, 21)
        Me.cmb_URCSYear.TabIndex = 41
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(161, 74)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(65, 13)
        Me.lbl_Select_Year_Combobox.TabIndex = 42
        Me.lbl_Select_Year_Combobox.Text = "Select Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frm_Tare_Weight_Loader
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(544, 279)
        Me.ControlBox = False
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.cmb_URCSYear)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Report_File_Entry)
        Me.Controls.Add(Me.btn_Input_File_Entry)
        Me.Controls.Add(Me.txt_Report_FilePath)
        Me.Controls.Add(Me.txt_Input_FilePath)
        Me.Controls.Add(Me.btn_Return_To_DataPrepMenu)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frm_Tare_Weight_Loader"
        Me.Text = "URCS Tare Weight Loader"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents btn_Return_To_DataPrepMenu As System.Windows.Forms.Button
    Friend WithEvents txt_Input_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents txt_Report_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents btn_Report_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents btn_Input_File_Entry As System.Windows.Forms.Button
    Friend WithEvents OpenInputFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents OpenReportFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmb_URCSYear As System.Windows.Forms.ComboBox
    Friend WithEvents lbl_Select_Year_Combobox As System.Windows.Forms.Label
End Class
