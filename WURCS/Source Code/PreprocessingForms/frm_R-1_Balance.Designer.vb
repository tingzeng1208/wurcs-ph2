<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_R_1_Balance
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
        Me.lblCS54_Data_Load = New System.Windows.Forms.Label()
        Me.txt_Output_FilePath = New System.Windows.Forms.TextBox()
        Me.cmb_URCS_Year = New System.Windows.Forms.ComboBox()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.btn_Return_To_DataPrepMenu = New System.Windows.Forms.Button()
        Me.btn_Output_File_Entry = New System.Windows.Forms.Button()
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblCS54_Data_Load
        '
        Me.lblCS54_Data_Load.AutoSize = True
        Me.lblCS54_Data_Load.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.lblCS54_Data_Load.Location = New System.Drawing.Point(186, 9)
        Me.lblCS54_Data_Load.Name = "lblCS54_Data_Load"
        Me.lblCS54_Data_Load.Size = New System.Drawing.Size(209, 46)
        Me.lblCS54_Data_Load.TabIndex = 71
        Me.lblCS54_Data_Load.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "R-1 Balance Check Program"
        Me.lblCS54_Data_Load.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txt_Output_FilePath
        '
        Me.txt_Output_FilePath.Location = New System.Drawing.Point(154, 89)
        Me.txt_Output_FilePath.Name = "txt_Output_FilePath"
        Me.txt_Output_FilePath.Size = New System.Drawing.Size(383, 21)
        Me.txt_Output_FilePath.TabIndex = 81
        '
        'cmb_URCS_Year
        '
        Me.cmb_URCS_Year.FormattingEnabled = True
        Me.cmb_URCS_Year.Location = New System.Drawing.Point(245, 62)
        Me.cmb_URCS_Year.Name = "cmb_URCS_Year"
        Me.cmb_URCS_Year.Size = New System.Drawing.Size(102, 21)
        Me.cmb_URCS_Year.TabIndex = 80
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(71, 113)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 79
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(243, 135)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(104, 27)
        Me.btn_Execute.TabIndex = 77
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'btn_Return_To_DataPrepMenu
        '
        Me.btn_Return_To_DataPrepMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_DataPrepMenu.Location = New System.Drawing.Point(515, 189)
        Me.btn_Return_To_DataPrepMenu.Name = "btn_Return_To_DataPrepMenu"
        Me.btn_Return_To_DataPrepMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_DataPrepMenu.TabIndex = 78
        Me.btn_Return_To_DataPrepMenu.UseVisualStyleBackColor = True
        '
        'btn_Output_File_Entry
        '
        Me.btn_Output_File_Entry.Location = New System.Drawing.Point(38, 89)
        Me.btn_Output_File_Entry.Name = "btn_Output_File_Entry"
        Me.btn_Output_File_Entry.Size = New System.Drawing.Size(110, 21)
        Me.btn_Output_File_Entry.TabIndex = 83
        Me.btn_Output_File_Entry.TabStop = False
        Me.btn_Output_File_Entry.Text = "Select Output File:"
        Me.btn_Output_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_File_Entry.UseVisualStyleBackColor = True
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(179, 63)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(65, 13)
        Me.lbl_Select_Year_Combobox.TabIndex = 82
        Me.lbl_Select_Year_Combobox.Text = "Select Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frm_R_1_Balance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(578, 262)
        Me.Controls.Add(Me.btn_Output_File_Entry)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.txt_Output_FilePath)
        Me.Controls.Add(Me.cmb_URCS_Year)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Return_To_DataPrepMenu)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.lblCS54_Data_Load)
        Me.Name = "frm_R_1_Balance"
        Me.Text = "frm_R_1_Balance"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblCS54_Data_Load As System.Windows.Forms.Label
    Friend WithEvents txt_Output_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents cmb_URCS_Year As System.Windows.Forms.ComboBox
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents btn_Return_To_DataPrepMenu As System.Windows.Forms.Button
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents btn_Output_File_Entry As System.Windows.Forms.Button
    Friend WithEvents lbl_Select_Year_Combobox As System.Windows.Forms.Label
End Class
