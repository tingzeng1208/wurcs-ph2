<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AAR_Index_Loader
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
        Me.btn_Return_To_DataPrepMenu = New System.Windows.Forms.Button()
        Me.lblCS54_Data_Load = New System.Windows.Forms.Label()
        Me.URCS_Year_Combobox = New System.Windows.Forms.ComboBox()
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.Label_Fuel = New System.Windows.Forms.Label()
        Me.Label_Materials_Supplies = New System.Windows.Forms.Label()
        Me.Label_Purchased_Services = New System.Windows.Forms.Label()
        Me.Label_Wages_Supplements = New System.Windows.Forms.Label()
        Me.Label_Material_Prices_Wage_Rates = New System.Windows.Forms.Label()
        Me.TextBox_Fuel_US = New System.Windows.Forms.TextBox()
        Me.TextBox_MS_US = New System.Windows.Forms.TextBox()
        Me.TextBox_PS_US = New System.Windows.Forms.TextBox()
        Me.TextBox_Wage_US = New System.Windows.Forms.TextBox()
        Me.TextBox_MP_US = New System.Windows.Forms.TextBox()
        Me.Label_US_Column = New System.Windows.Forms.Label()
        Me.TextBox_Fuel_East = New System.Windows.Forms.TextBox()
        Me.TextBox_MS_East = New System.Windows.Forms.TextBox()
        Me.TextBox_PS_East = New System.Windows.Forms.TextBox()
        Me.TextBox_Wage_East = New System.Windows.Forms.TextBox()
        Me.TextBox_MP_East = New System.Windows.Forms.TextBox()
        Me.Label_East_Column = New System.Windows.Forms.Label()
        Me.TextBox_MP_West = New System.Windows.Forms.TextBox()
        Me.TextBox_Wage_West = New System.Windows.Forms.TextBox()
        Me.TextBox_PS_West = New System.Windows.Forms.TextBox()
        Me.TextBox_MS_West = New System.Windows.Forms.TextBox()
        Me.TextBox_Fuel_West = New System.Windows.Forms.TextBox()
        Me.Label_West_Column = New System.Windows.Forms.Label()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.btn_Report_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Report_FilePath = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btn_Return_To_DataPrepMenu
        '
        Me.btn_Return_To_DataPrepMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_DataPrepMenu.Location = New System.Drawing.Point(528, 316)
        Me.btn_Return_To_DataPrepMenu.Name = "btn_Return_To_DataPrepMenu"
        Me.btn_Return_To_DataPrepMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_DataPrepMenu.TabIndex = 43
        Me.btn_Return_To_DataPrepMenu.UseVisualStyleBackColor = True
        '
        'lblCS54_Data_Load
        '
        Me.lblCS54_Data_Load.AutoSize = True
        Me.lblCS54_Data_Load.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.lblCS54_Data_Load.Location = New System.Drawing.Point(191, 9)
        Me.lblCS54_Data_Load.Name = "lblCS54_Data_Load"
        Me.lblCS54_Data_Load.Size = New System.Drawing.Size(223, 46)
        Me.lblCS54_Data_Load.TabIndex = 44
        Me.lblCS54_Data_Load.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "AAR Index Data Load Program"
        Me.lblCS54_Data_Load.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'URCS_Year_Combobox
        '
        Me.URCS_Year_Combobox.FormattingEnabled = True
        Me.URCS_Year_Combobox.Location = New System.Drawing.Point(287, 64)
        Me.URCS_Year_Combobox.Name = "URCS_Year_Combobox"
        Me.URCS_Year_Combobox.Size = New System.Drawing.Size(102, 21)
        Me.URCS_Year_Combobox.TabIndex = 57
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(210, 65)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(62, 13)
        Me.lbl_Select_Year_Combobox.TabIndex = 56
        Me.lbl_Select_Year_Combobox.Text = "Enter Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label_Fuel
        '
        Me.Label_Fuel.AutoSize = True
        Me.Label_Fuel.Location = New System.Drawing.Point(226, 119)
        Me.Label_Fuel.Name = "Label_Fuel"
        Me.Label_Fuel.Size = New System.Drawing.Size(31, 13)
        Me.Label_Fuel.TabIndex = 58
        Me.Label_Fuel.Text = "Fuel:"
        '
        'Label_Materials_Supplies
        '
        Me.Label_Materials_Supplies.AutoSize = True
        Me.Label_Materials_Supplies.Location = New System.Drawing.Point(139, 142)
        Me.Label_Materials_Supplies.Name = "Label_Materials_Supplies"
        Me.Label_Materials_Supplies.Size = New System.Drawing.Size(118, 13)
        Me.Label_Materials_Supplies.TabIndex = 59
        Me.Label_Materials_Supplies.Text = "Materials And Supplies:"
        '
        'Label_Purchased_Services
        '
        Me.Label_Purchased_Services.AutoSize = True
        Me.Label_Purchased_Services.Location = New System.Drawing.Point(153, 167)
        Me.Label_Purchased_Services.Name = "Label_Purchased_Services"
        Me.Label_Purchased_Services.Size = New System.Drawing.Size(104, 13)
        Me.Label_Purchased_Services.TabIndex = 60
        Me.Label_Purchased_Services.Text = "Purchased Services:"
        '
        'Label_Wages_Supplements
        '
        Me.Label_Wages_Supplements.AutoSize = True
        Me.Label_Wages_Supplements.Location = New System.Drawing.Point(102, 192)
        Me.Label_Wages_Supplements.Name = "Label_Wages_Supplements"
        Me.Label_Wages_Supplements.Size = New System.Drawing.Size(155, 13)
        Me.Label_Wages_Supplements.TabIndex = 61
        Me.Label_Wages_Supplements.Text = "Wage Rates and Supplements:"
        '
        'Label_Material_Prices_Wage_Rates
        '
        Me.Label_Material_Prices_Wage_Rates.AutoSize = True
        Me.Label_Material_Prices_Wage_Rates.Location = New System.Drawing.Point(94, 216)
        Me.Label_Material_Prices_Wage_Rates.Name = "Label_Material_Prices_Wage_Rates"
        Me.Label_Material_Prices_Wage_Rates.Size = New System.Drawing.Size(163, 13)
        Me.Label_Material_Prices_Wage_Rates.TabIndex = 62
        Me.Label_Material_Prices_Wage_Rates.Text = "Material Prices and Wage Rates:"
        '
        'TextBox_Fuel_US
        '
        Me.TextBox_Fuel_US.Location = New System.Drawing.Point(263, 115)
        Me.TextBox_Fuel_US.Name = "TextBox_Fuel_US"
        Me.TextBox_Fuel_US.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_Fuel_US.TabIndex = 63
        Me.TextBox_Fuel_US.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_MS_US
        '
        Me.TextBox_MS_US.Location = New System.Drawing.Point(263, 139)
        Me.TextBox_MS_US.Name = "TextBox_MS_US"
        Me.TextBox_MS_US.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_MS_US.TabIndex = 64
        Me.TextBox_MS_US.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_PS_US
        '
        Me.TextBox_PS_US.Location = New System.Drawing.Point(263, 163)
        Me.TextBox_PS_US.Name = "TextBox_PS_US"
        Me.TextBox_PS_US.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_PS_US.TabIndex = 65
        Me.TextBox_PS_US.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_Wage_US
        '
        Me.TextBox_Wage_US.Location = New System.Drawing.Point(263, 187)
        Me.TextBox_Wage_US.Name = "TextBox_Wage_US"
        Me.TextBox_Wage_US.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_Wage_US.TabIndex = 66
        Me.TextBox_Wage_US.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_MP_US
        '
        Me.TextBox_MP_US.Location = New System.Drawing.Point(263, 211)
        Me.TextBox_MP_US.Name = "TextBox_MP_US"
        Me.TextBox_MP_US.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_MP_US.TabIndex = 67
        Me.TextBox_MP_US.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label_US_Column
        '
        Me.Label_US_Column.AutoSize = True
        Me.Label_US_Column.Font = New System.Drawing.Font("Tahoma", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_US_Column.Location = New System.Drawing.Point(283, 99)
        Me.Label_US_Column.Name = "Label_US_Column"
        Me.Label_US_Column.Size = New System.Drawing.Size(22, 13)
        Me.Label_US_Column.TabIndex = 68
        Me.Label_US_Column.Text = "US"
        '
        'TextBox_Fuel_East
        '
        Me.TextBox_Fuel_East.Location = New System.Drawing.Point(329, 116)
        Me.TextBox_Fuel_East.Name = "TextBox_Fuel_East"
        Me.TextBox_Fuel_East.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_Fuel_East.TabIndex = 69
        Me.TextBox_Fuel_East.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_MS_East
        '
        Me.TextBox_MS_East.Location = New System.Drawing.Point(329, 140)
        Me.TextBox_MS_East.Name = "TextBox_MS_East"
        Me.TextBox_MS_East.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_MS_East.TabIndex = 70
        Me.TextBox_MS_East.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_PS_East
        '
        Me.TextBox_PS_East.Location = New System.Drawing.Point(329, 164)
        Me.TextBox_PS_East.Name = "TextBox_PS_East"
        Me.TextBox_PS_East.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_PS_East.TabIndex = 71
        Me.TextBox_PS_East.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_Wage_East
        '
        Me.TextBox_Wage_East.Location = New System.Drawing.Point(329, 188)
        Me.TextBox_Wage_East.Name = "TextBox_Wage_East"
        Me.TextBox_Wage_East.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_Wage_East.TabIndex = 72
        Me.TextBox_Wage_East.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_MP_East
        '
        Me.TextBox_MP_East.Location = New System.Drawing.Point(329, 212)
        Me.TextBox_MP_East.Name = "TextBox_MP_East"
        Me.TextBox_MP_East.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_MP_East.TabIndex = 73
        Me.TextBox_MP_East.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label_East_Column
        '
        Me.Label_East_Column.AutoSize = True
        Me.Label_East_Column.Font = New System.Drawing.Font("Tahoma", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_East_Column.Location = New System.Drawing.Point(341, 99)
        Me.Label_East_Column.Name = "Label_East_Column"
        Me.Label_East_Column.Size = New System.Drawing.Size(31, 13)
        Me.Label_East_Column.TabIndex = 74
        Me.Label_East_Column.Text = "East"
        '
        'TextBox_MP_West
        '
        Me.TextBox_MP_West.Location = New System.Drawing.Point(395, 213)
        Me.TextBox_MP_West.Name = "TextBox_MP_West"
        Me.TextBox_MP_West.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_MP_West.TabIndex = 79
        Me.TextBox_MP_West.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_Wage_West
        '
        Me.TextBox_Wage_West.Location = New System.Drawing.Point(395, 189)
        Me.TextBox_Wage_West.Name = "TextBox_Wage_West"
        Me.TextBox_Wage_West.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_Wage_West.TabIndex = 78
        Me.TextBox_Wage_West.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_PS_West
        '
        Me.TextBox_PS_West.Location = New System.Drawing.Point(395, 165)
        Me.TextBox_PS_West.Name = "TextBox_PS_West"
        Me.TextBox_PS_West.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_PS_West.TabIndex = 77
        Me.TextBox_PS_West.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_MS_West
        '
        Me.TextBox_MS_West.Location = New System.Drawing.Point(395, 141)
        Me.TextBox_MS_West.Name = "TextBox_MS_West"
        Me.TextBox_MS_West.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_MS_West.TabIndex = 76
        Me.TextBox_MS_West.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_Fuel_West
        '
        Me.TextBox_Fuel_West.Location = New System.Drawing.Point(395, 117)
        Me.TextBox_Fuel_West.Name = "TextBox_Fuel_West"
        Me.TextBox_Fuel_West.Size = New System.Drawing.Size(60, 21)
        Me.TextBox_Fuel_West.TabIndex = 75
        Me.TextBox_Fuel_West.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label_West_Column
        '
        Me.Label_West_Column.AutoSize = True
        Me.Label_West_Column.Font = New System.Drawing.Font("Tahoma", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_West_Column.Location = New System.Drawing.Point(411, 99)
        Me.Label_West_Column.Name = "Label_West_Column"
        Me.Label_West_Column.Size = New System.Drawing.Size(36, 13)
        Me.Label_West_Column.TabIndex = 80
        Me.Label_West_Column.Text = "West"
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(82, 269)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 83
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(247, 292)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(104, 27)
        Me.btn_Execute.TabIndex = 84
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'btn_Report_File_Entry
        '
        Me.btn_Report_File_Entry.Location = New System.Drawing.Point(55, 240)
        Me.btn_Report_File_Entry.Name = "btn_Report_File_Entry"
        Me.btn_Report_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Report_File_Entry.TabIndex = 86
        Me.btn_Report_File_Entry.TabStop = False
        Me.btn_Report_File_Entry.Text = "Select Report File:"
        Me.btn_Report_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Report_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Report_FilePath
        '
        Me.txt_Report_FilePath.Location = New System.Drawing.Point(166, 240)
        Me.txt_Report_FilePath.Name = "txt_Report_FilePath"
        Me.txt_Report_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Report_FilePath.TabIndex = 85
        '
        'AAR_Index_Loader
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(591, 382)
        Me.ControlBox = False
        Me.Controls.Add(Me.btn_Report_File_Entry)
        Me.Controls.Add(Me.txt_Report_FilePath)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.Label_West_Column)
        Me.Controls.Add(Me.TextBox_MP_West)
        Me.Controls.Add(Me.TextBox_Wage_West)
        Me.Controls.Add(Me.TextBox_PS_West)
        Me.Controls.Add(Me.TextBox_MS_West)
        Me.Controls.Add(Me.TextBox_Fuel_West)
        Me.Controls.Add(Me.Label_East_Column)
        Me.Controls.Add(Me.TextBox_MP_East)
        Me.Controls.Add(Me.TextBox_Wage_East)
        Me.Controls.Add(Me.TextBox_PS_East)
        Me.Controls.Add(Me.TextBox_MS_East)
        Me.Controls.Add(Me.TextBox_Fuel_East)
        Me.Controls.Add(Me.Label_US_Column)
        Me.Controls.Add(Me.TextBox_MP_US)
        Me.Controls.Add(Me.TextBox_Wage_US)
        Me.Controls.Add(Me.TextBox_PS_US)
        Me.Controls.Add(Me.TextBox_MS_US)
        Me.Controls.Add(Me.TextBox_Fuel_US)
        Me.Controls.Add(Me.Label_Material_Prices_Wage_Rates)
        Me.Controls.Add(Me.Label_Wages_Supplements)
        Me.Controls.Add(Me.Label_Purchased_Services)
        Me.Controls.Add(Me.Label_Materials_Supplies)
        Me.Controls.Add(Me.Label_Fuel)
        Me.Controls.Add(Me.URCS_Year_Combobox)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.lblCS54_Data_Load)
        Me.Controls.Add(Me.btn_Return_To_DataPrepMenu)
        Me.Name = "AAR_Index_Loader"
        Me.Text = "AAR_Index"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Return_To_DataPrepMenu As System.Windows.Forms.Button
    Friend WithEvents lblCS54_Data_Load As System.Windows.Forms.Label
    Friend WithEvents URCS_Year_Combobox As System.Windows.Forms.ComboBox
    Friend WithEvents lbl_Select_Year_Combobox As System.Windows.Forms.Label
    Friend WithEvents Label_Fuel As System.Windows.Forms.Label
    Friend WithEvents Label_Materials_Supplies As System.Windows.Forms.Label
    Friend WithEvents Label_Purchased_Services As System.Windows.Forms.Label
    Friend WithEvents Label_Wages_Supplements As System.Windows.Forms.Label
    Friend WithEvents Label_Material_Prices_Wage_Rates As System.Windows.Forms.Label
    Friend WithEvents TextBox_Fuel_US As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_MS_US As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_PS_US As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Wage_US As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_MP_US As System.Windows.Forms.TextBox
    Friend WithEvents Label_US_Column As System.Windows.Forms.Label
    Friend WithEvents TextBox_Fuel_East As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_MS_East As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_PS_East As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Wage_East As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_MP_East As System.Windows.Forms.TextBox
    Friend WithEvents Label_East_Column As System.Windows.Forms.Label
    Friend WithEvents TextBox_MP_West As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Wage_West As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_PS_West As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_MS_West As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Fuel_West As System.Windows.Forms.TextBox
    Friend WithEvents Label_West_Column As System.Windows.Forms.Label
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents btn_Report_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Report_FilePath As System.Windows.Forms.TextBox
End Class
