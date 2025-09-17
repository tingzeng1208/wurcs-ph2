<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Interim_Data_Loader
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
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.lbl_Form_Title = New System.Windows.Forms.Label()
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.btn_Input_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Input_FilePath = New System.Windows.Forms.TextBox()
        Me.rdo_Annual = New System.Windows.Forms.RadioButton()
        Me.rdo_Quarterly = New System.Windows.Forms.RadioButton()
        Me.rdo_Monthly = New System.Windows.Forms.RadioButton()
        Me.gbx_Data_Type = New System.Windows.Forms.GroupBox()
        Me.txt_Year = New System.Windows.Forms.TextBox()
        Me.rdo_1st_Quarter = New System.Windows.Forms.RadioButton()
        Me.rdo_4th_Quarter = New System.Windows.Forms.RadioButton()
        Me.rdo_3rd_Quarter = New System.Windows.Forms.RadioButton()
        Me.rdo_2nd_Quarter = New System.Windows.Forms.RadioButton()
        Me.gbx_Quarter = New System.Windows.Forms.GroupBox()
        Me.cmb_Month = New System.Windows.Forms.ComboBox()
        Me.lbl_Month_Combobox = New System.Windows.Forms.Label()
        Me.btn_Return_To_DataPrepMenu = New System.Windows.Forms.Button()
        Me.gbx_Monthly = New System.Windows.Forms.GroupBox()
        Me.btn_Select_Work_Directory = New System.Windows.Forms.Button()
        Me.txt_Work_Directory = New System.Windows.Forms.TextBox()
        Me.chk_Skip_Masked_Data_Load = New System.Windows.Forms.CheckBox()
        Me.chk_Skip_BatchPro_Processing = New System.Windows.Forms.CheckBox()
        Me.chk_Skip_Segments_Data_Load = New System.Windows.Forms.CheckBox()
        Me.gbx_Data_Type.SuspendLayout()
        Me.gbx_Quarter.SuspendLayout()
        Me.gbx_Monthly.SuspendLayout()
        Me.SuspendLayout()
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(90, 567)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(9, 8, 9, 8)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(648, 20)
        Me.txt_StatusBox.TabIndex = 8
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(336, 619)
        Me.btn_Execute.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(156, 39)
        Me.btn_Execute.TabIndex = 9
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'lbl_Form_Title
        '
        Me.lbl_Form_Title.AutoSize = True
        Me.lbl_Form_Title.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.lbl_Form_Title.Location = New System.Drawing.Point(246, 11)
        Me.lbl_Form_Title.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbl_Form_Title.Name = "lbl_Form_Title"
        Me.lbl_Form_Title.Size = New System.Drawing.Size(304, 66)
        Me.lbl_Form_Title.TabIndex = 0
        Me.lbl_Form_Title.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Interim Waybill Data Loader"
        Me.lbl_Form_Title.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(280, 99)
        Me.lbl_Select_Year_Combobox.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(138, 19)
        Me.lbl_Select_Year_Combobox.TabIndex = 1
        Me.lbl_Select_Year_Combobox.Text = "Enter Year (yyyy):"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_Input_File_Entry
        '
        Me.btn_Input_File_Entry.Location = New System.Drawing.Point(21, 489)
        Me.btn_Input_File_Entry.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btn_Input_File_Entry.Name = "btn_Input_File_Entry"
        Me.btn_Input_File_Entry.Size = New System.Drawing.Size(187, 31)
        Me.btn_Input_File_Entry.TabIndex = 6
        Me.btn_Input_File_Entry.TabStop = False
        Me.btn_Input_File_Entry.Text = "Select Input File:"
        Me.btn_Input_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Input_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Input_FilePath
        '
        Me.txt_Input_FilePath.Location = New System.Drawing.Point(215, 493)
        Me.txt_Input_FilePath.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_Input_FilePath.Name = "txt_Input_FilePath"
        Me.txt_Input_FilePath.Size = New System.Drawing.Size(548, 27)
        Me.txt_Input_FilePath.TabIndex = 7
        '
        'rdo_Annual
        '
        Me.rdo_Annual.AutoSize = True
        Me.rdo_Annual.Location = New System.Drawing.Point(33, 90)
        Me.rdo_Annual.Margin = New System.Windows.Forms.Padding(4)
        Me.rdo_Annual.Name = "rdo_Annual"
        Me.rdo_Annual.Size = New System.Drawing.Size(84, 23)
        Me.rdo_Annual.TabIndex = 2
        Me.rdo_Annual.TabStop = True
        Me.rdo_Annual.Text = "Annual"
        Me.rdo_Annual.UseVisualStyleBackColor = True
        '
        'rdo_Quarterly
        '
        Me.rdo_Quarterly.AutoSize = True
        Me.rdo_Quarterly.Location = New System.Drawing.Point(33, 58)
        Me.rdo_Quarterly.Margin = New System.Windows.Forms.Padding(4)
        Me.rdo_Quarterly.Name = "rdo_Quarterly"
        Me.rdo_Quarterly.Size = New System.Drawing.Size(100, 23)
        Me.rdo_Quarterly.TabIndex = 1
        Me.rdo_Quarterly.TabStop = True
        Me.rdo_Quarterly.Text = "Quarterly"
        Me.rdo_Quarterly.UseVisualStyleBackColor = True
        '
        'rdo_Monthly
        '
        Me.rdo_Monthly.AutoSize = True
        Me.rdo_Monthly.Location = New System.Drawing.Point(33, 26)
        Me.rdo_Monthly.Margin = New System.Windows.Forms.Padding(4)
        Me.rdo_Monthly.Name = "rdo_Monthly"
        Me.rdo_Monthly.Size = New System.Drawing.Size(90, 23)
        Me.rdo_Monthly.TabIndex = 0
        Me.rdo_Monthly.TabStop = True
        Me.rdo_Monthly.Text = "Monthly"
        Me.rdo_Monthly.UseVisualStyleBackColor = True
        '
        'gbx_Data_Type
        '
        Me.gbx_Data_Type.Controls.Add(Me.rdo_Monthly)
        Me.gbx_Data_Type.Controls.Add(Me.rdo_Quarterly)
        Me.gbx_Data_Type.Controls.Add(Me.rdo_Annual)
        Me.gbx_Data_Type.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbx_Data_Type.Location = New System.Drawing.Point(201, 156)
        Me.gbx_Data_Type.Margin = New System.Windows.Forms.Padding(4)
        Me.gbx_Data_Type.Name = "gbx_Data_Type"
        Me.gbx_Data_Type.Padding = New System.Windows.Forms.Padding(4)
        Me.gbx_Data_Type.Size = New System.Drawing.Size(193, 140)
        Me.gbx_Data_Type.TabIndex = 3
        Me.gbx_Data_Type.TabStop = False
        Me.gbx_Data_Type.Text = "Select Data File Type"
        '
        'txt_Year
        '
        Me.txt_Year.Location = New System.Drawing.Point(446, 95)
        Me.txt_Year.Margin = New System.Windows.Forms.Padding(4)
        Me.txt_Year.Name = "txt_Year"
        Me.txt_Year.Size = New System.Drawing.Size(86, 27)
        Me.txt_Year.TabIndex = 2
        '
        'rdo_1st_Quarter
        '
        Me.rdo_1st_Quarter.AutoSize = True
        Me.rdo_1st_Quarter.Location = New System.Drawing.Point(26, 26)
        Me.rdo_1st_Quarter.Margin = New System.Windows.Forms.Padding(4)
        Me.rdo_1st_Quarter.Name = "rdo_1st_Quarter"
        Me.rdo_1st_Quarter.Size = New System.Drawing.Size(114, 23)
        Me.rdo_1st_Quarter.TabIndex = 0
        Me.rdo_1st_Quarter.TabStop = True
        Me.rdo_1st_Quarter.Text = "1st Quarter"
        Me.rdo_1st_Quarter.UseVisualStyleBackColor = True
        Me.rdo_1st_Quarter.Visible = False
        '
        'rdo_4th_Quarter
        '
        Me.rdo_4th_Quarter.AutoSize = True
        Me.rdo_4th_Quarter.Location = New System.Drawing.Point(26, 122)
        Me.rdo_4th_Quarter.Margin = New System.Windows.Forms.Padding(4)
        Me.rdo_4th_Quarter.Name = "rdo_4th_Quarter"
        Me.rdo_4th_Quarter.Size = New System.Drawing.Size(116, 23)
        Me.rdo_4th_Quarter.TabIndex = 3
        Me.rdo_4th_Quarter.TabStop = True
        Me.rdo_4th_Quarter.Text = "4th Quarter"
        Me.rdo_4th_Quarter.UseVisualStyleBackColor = True
        Me.rdo_4th_Quarter.Visible = False
        '
        'rdo_3rd_Quarter
        '
        Me.rdo_3rd_Quarter.AutoSize = True
        Me.rdo_3rd_Quarter.Location = New System.Drawing.Point(26, 91)
        Me.rdo_3rd_Quarter.Margin = New System.Windows.Forms.Padding(4)
        Me.rdo_3rd_Quarter.Name = "rdo_3rd_Quarter"
        Me.rdo_3rd_Quarter.Size = New System.Drawing.Size(117, 23)
        Me.rdo_3rd_Quarter.TabIndex = 2
        Me.rdo_3rd_Quarter.TabStop = True
        Me.rdo_3rd_Quarter.Text = "3rd Quarter"
        Me.rdo_3rd_Quarter.UseVisualStyleBackColor = True
        Me.rdo_3rd_Quarter.Visible = False
        '
        'rdo_2nd_Quarter
        '
        Me.rdo_2nd_Quarter.AutoSize = True
        Me.rdo_2nd_Quarter.Location = New System.Drawing.Point(26, 58)
        Me.rdo_2nd_Quarter.Margin = New System.Windows.Forms.Padding(4)
        Me.rdo_2nd_Quarter.Name = "rdo_2nd_Quarter"
        Me.rdo_2nd_Quarter.Size = New System.Drawing.Size(120, 23)
        Me.rdo_2nd_Quarter.TabIndex = 1
        Me.rdo_2nd_Quarter.TabStop = True
        Me.rdo_2nd_Quarter.Text = "2nd Quarter"
        Me.rdo_2nd_Quarter.UseVisualStyleBackColor = True
        Me.rdo_2nd_Quarter.Visible = False
        '
        'gbx_Quarter
        '
        Me.gbx_Quarter.Controls.Add(Me.rdo_2nd_Quarter)
        Me.gbx_Quarter.Controls.Add(Me.rdo_3rd_Quarter)
        Me.gbx_Quarter.Controls.Add(Me.rdo_4th_Quarter)
        Me.gbx_Quarter.Controls.Add(Me.rdo_1st_Quarter)
        Me.gbx_Quarter.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbx_Quarter.Location = New System.Drawing.Point(413, 214)
        Me.gbx_Quarter.Margin = New System.Windows.Forms.Padding(4)
        Me.gbx_Quarter.Name = "gbx_Quarter"
        Me.gbx_Quarter.Padding = New System.Windows.Forms.Padding(4)
        Me.gbx_Quarter.Size = New System.Drawing.Size(159, 167)
        Me.gbx_Quarter.TabIndex = 4
        Me.gbx_Quarter.TabStop = False
        Me.gbx_Quarter.Text = "Select Quarter"
        Me.gbx_Quarter.Visible = False
        '
        'cmb_Month
        '
        Me.cmb_Month.FormattingEnabled = True
        Me.cmb_Month.Items.AddRange(New Object() {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"})
        Me.cmb_Month.Location = New System.Drawing.Point(143, 10)
        Me.cmb_Month.Margin = New System.Windows.Forms.Padding(4)
        Me.cmb_Month.Name = "cmb_Month"
        Me.cmb_Month.Size = New System.Drawing.Size(54, 27)
        Me.cmb_Month.TabIndex = 1
        Me.cmb_Month.Visible = False
        '
        'lbl_Month_Combobox
        '
        Me.lbl_Month_Combobox.AutoSize = True
        Me.lbl_Month_Combobox.Location = New System.Drawing.Point(14, 13)
        Me.lbl_Month_Combobox.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbl_Month_Combobox.Name = "lbl_Month_Combobox"
        Me.lbl_Month_Combobox.Size = New System.Drawing.Size(105, 19)
        Me.lbl_Month_Combobox.TabIndex = 0
        Me.lbl_Month_Combobox.Text = "Select Month:"
        Me.lbl_Month_Combobox.Visible = False
        '
        'btn_Return_To_DataPrepMenu
        '
        Me.btn_Return_To_DataPrepMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_DataPrepMenu.Location = New System.Drawing.Point(723, 600)
        Me.btn_Return_To_DataPrepMenu.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btn_Return_To_DataPrepMenu.Name = "btn_Return_To_DataPrepMenu"
        Me.btn_Return_To_DataPrepMenu.Size = New System.Drawing.Size(76, 76)
        Me.btn_Return_To_DataPrepMenu.TabIndex = 10
        Me.btn_Return_To_DataPrepMenu.UseVisualStyleBackColor = True
        '
        'gbx_Monthly
        '
        Me.gbx_Monthly.Controls.Add(Me.lbl_Month_Combobox)
        Me.gbx_Monthly.Controls.Add(Me.cmb_Month)
        Me.gbx_Monthly.Location = New System.Drawing.Point(413, 147)
        Me.gbx_Monthly.Margin = New System.Windows.Forms.Padding(4)
        Me.gbx_Monthly.Name = "gbx_Monthly"
        Me.gbx_Monthly.Padding = New System.Windows.Forms.Padding(4)
        Me.gbx_Monthly.Size = New System.Drawing.Size(199, 59)
        Me.gbx_Monthly.TabIndex = 5
        Me.gbx_Monthly.TabStop = False
        '
        'btn_Select_Work_Directory
        '
        Me.btn_Select_Work_Directory.Location = New System.Drawing.Point(21, 526)
        Me.btn_Select_Work_Directory.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btn_Select_Work_Directory.Name = "btn_Select_Work_Directory"
        Me.btn_Select_Work_Directory.Size = New System.Drawing.Size(187, 31)
        Me.btn_Select_Work_Directory.TabIndex = 17
        Me.btn_Select_Work_Directory.TabStop = False
        Me.btn_Select_Work_Directory.Text = "Select Work Directory:"
        Me.btn_Select_Work_Directory.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Select_Work_Directory.UseVisualStyleBackColor = True
        '
        'txt_Work_Directory
        '
        Me.txt_Work_Directory.Location = New System.Drawing.Point(216, 530)
        Me.txt_Work_Directory.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_Work_Directory.Name = "txt_Work_Directory"
        Me.txt_Work_Directory.Size = New System.Drawing.Size(548, 27)
        Me.txt_Work_Directory.TabIndex = 18
        '
        'chk_Skip_Masked_Data_Load
        '
        Me.chk_Skip_Masked_Data_Load.AutoSize = True
        Me.chk_Skip_Masked_Data_Load.Location = New System.Drawing.Point(234, 388)
        Me.chk_Skip_Masked_Data_Load.Name = "chk_Skip_Masked_Data_Load"
        Me.chk_Skip_Masked_Data_Load.Size = New System.Drawing.Size(198, 23)
        Me.chk_Skip_Masked_Data_Load.TabIndex = 19
        Me.chk_Skip_Masked_Data_Load.Text = "Skip Masked Data Load"
        Me.chk_Skip_Masked_Data_Load.UseVisualStyleBackColor = True
        '
        'chk_Skip_BatchPro_Processing
        '
        Me.chk_Skip_BatchPro_Processing.AutoSize = True
        Me.chk_Skip_BatchPro_Processing.Location = New System.Drawing.Point(234, 417)
        Me.chk_Skip_BatchPro_Processing.Name = "chk_Skip_BatchPro_Processing"
        Me.chk_Skip_BatchPro_Processing.Size = New System.Drawing.Size(212, 23)
        Me.chk_Skip_BatchPro_Processing.TabIndex = 20
        Me.chk_Skip_BatchPro_Processing.Text = "Skip BatchPro Processing"
        Me.chk_Skip_BatchPro_Processing.UseVisualStyleBackColor = True
        '
        'chk_Skip_Segments_Data_Load
        '
        Me.chk_Skip_Segments_Data_Load.AutoSize = True
        Me.chk_Skip_Segments_Data_Load.Location = New System.Drawing.Point(234, 446)
        Me.chk_Skip_Segments_Data_Load.Name = "chk_Skip_Segments_Data_Load"
        Me.chk_Skip_Segments_Data_Load.Size = New System.Drawing.Size(325, 23)
        Me.chk_Skip_Segments_Data_Load.TabIndex = 21
        Me.chk_Skip_Segments_Data_Load.Text = "Skip Segments Data Load from OUT files"
        Me.chk_Skip_Segments_Data_Load.UseVisualStyleBackColor = True
        '
        'Interim_Data_Loader
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(813, 719)
        Me.Controls.Add(Me.chk_Skip_Segments_Data_Load)
        Me.Controls.Add(Me.chk_Skip_BatchPro_Processing)
        Me.Controls.Add(Me.chk_Skip_Masked_Data_Load)
        Me.Controls.Add(Me.btn_Select_Work_Directory)
        Me.Controls.Add(Me.txt_Work_Directory)
        Me.Controls.Add(Me.gbx_Monthly)
        Me.Controls.Add(Me.gbx_Quarter)
        Me.Controls.Add(Me.txt_Year)
        Me.Controls.Add(Me.gbx_Data_Type)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.lbl_Form_Title)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.btn_Input_File_Entry)
        Me.Controls.Add(Me.txt_Input_FilePath)
        Me.Controls.Add(Me.btn_Return_To_DataPrepMenu)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Interim_Data_Loader"
        Me.Text = "Interim_Data_Loader"
        Me.gbx_Data_Type.ResumeLayout(False)
        Me.gbx_Data_Type.PerformLayout()
        Me.gbx_Quarter.ResumeLayout(False)
        Me.gbx_Quarter.PerformLayout()
        Me.gbx_Monthly.ResumeLayout(False)
        Me.gbx_Monthly.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txt_StatusBox As TextBox
    Friend WithEvents btn_Execute As Button
    Friend WithEvents lbl_Form_Title As Label
    Friend WithEvents lbl_Select_Year_Combobox As Label
    Friend WithEvents btn_Input_File_Entry As Button
    Friend WithEvents txt_Input_FilePath As TextBox
    Friend WithEvents btn_Return_To_DataPrepMenu As Button
    Friend WithEvents rdo_Annual As RadioButton
    Friend WithEvents rdo_Quarterly As RadioButton
    Friend WithEvents rdo_Monthly As RadioButton
    Friend WithEvents gbx_Data_Type As GroupBox
    Friend WithEvents txt_Year As TextBox
    Friend WithEvents rdo_1st_Quarter As RadioButton
    Friend WithEvents rdo_4th_Quarter As RadioButton
    Friend WithEvents rdo_3rd_Quarter As RadioButton
    Friend WithEvents rdo_2nd_Quarter As RadioButton
    Friend WithEvents gbx_Quarter As GroupBox
    Friend WithEvents cmb_Month As ComboBox
    Friend WithEvents lbl_Month_Combobox As Label
    Friend WithEvents gbx_Monthly As GroupBox
    Friend WithEvents btn_Select_Work_Directory As Button
    Friend WithEvents txt_Work_Directory As TextBox
    Friend WithEvents chk_Skip_Masked_Data_Load As CheckBox
    Friend WithEvents chk_Skip_BatchPro_Processing As CheckBox
    Friend WithEvents chk_Skip_Segments_Data_Load As CheckBox
End Class
