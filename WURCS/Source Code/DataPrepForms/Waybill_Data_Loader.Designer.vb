<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Waybill_Data_Loader
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
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.cmb_URCSYear = New System.Windows.Forms.ComboBox()
        Me.btn_Input_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Input_FilePath = New System.Windows.Forms.TextBox()
        Me.btn_Return_To_DataPrepMenu = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(67, 141)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 66
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(222, 165)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(104, 27)
        Me.btn_Execute.TabIndex = 65
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(203, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(149, 46)
        Me.Label1.TabIndex = 64
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Waybill Data Loader"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(189, 87)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(65, 13)
        Me.lbl_Select_Year_Combobox.TabIndex = 63
        Me.lbl_Select_Year_Combobox.Text = "Select Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmb_URCSYear
        '
        Me.cmb_URCSYear.FormattingEnabled = True
        Me.cmb_URCSYear.Location = New System.Drawing.Point(260, 84)
        Me.cmb_URCSYear.Name = "cmb_URCSYear"
        Me.cmb_URCSYear.Size = New System.Drawing.Size(108, 21)
        Me.cmb_URCSYear.TabIndex = 62
        '
        'btn_Input_File_Entry
        '
        Me.btn_Input_File_Entry.Location = New System.Drawing.Point(34, 111)
        Me.btn_Input_File_Entry.Name = "btn_Input_File_Entry"
        Me.btn_Input_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Input_File_Entry.TabIndex = 61
        Me.btn_Input_File_Entry.TabStop = False
        Me.btn_Input_File_Entry.Text = "Select Input File:"
        Me.btn_Input_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Input_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Input_FilePath
        '
        Me.txt_Input_FilePath.Location = New System.Drawing.Point(145, 111)
        Me.txt_Input_FilePath.Name = "txt_Input_FilePath"
        Me.txt_Input_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Input_FilePath.TabIndex = 60
        '
        'btn_Return_To_DataPrepMenu
        '
        Me.btn_Return_To_DataPrepMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_DataPrepMenu.Location = New System.Drawing.Point(489, 170)
        Me.btn_Return_To_DataPrepMenu.Name = "btn_Return_To_DataPrepMenu"
        Me.btn_Return_To_DataPrepMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_DataPrepMenu.TabIndex = 59
        Me.btn_Return_To_DataPrepMenu.UseVisualStyleBackColor = True
        '
        'Waybill_Data_Loader
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(557, 238)
        Me.ControlBox = False
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.cmb_URCSYear)
        Me.Controls.Add(Me.btn_Input_File_Entry)
        Me.Controls.Add(Me.txt_Input_FilePath)
        Me.Controls.Add(Me.btn_Return_To_DataPrepMenu)
        Me.Name = "Waybill_Data_Loader"
        Me.Text = "Waybill_Data_Loader"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lbl_Select_Year_Combobox As System.Windows.Forms.Label
    Friend WithEvents cmb_URCSYear As System.Windows.Forms.ComboBox
    Friend WithEvents btn_Input_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Input_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents btn_Return_To_DataPrepMenu As System.Windows.Forms.Button
End Class
