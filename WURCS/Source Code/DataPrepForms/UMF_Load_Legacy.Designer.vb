<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UMF_Load_Legacy
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
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.cmb_URCS_Year = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Output_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Output_FilePath = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txt_Processing_Year = New System.Windows.Forms.TextBox()
        Me.chk_Created_Blank_Owner_Records = New System.Windows.Forms.CheckBox()
        Me.chk_Created_System_Records = New System.Windows.Forms.CheckBox()
        Me.chk_Trans_Table_Update = New System.Windows.Forms.CheckBox()
        Me.btn_Return_To_DataPrepMenu = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(229, 71)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(65, 13)
        Me.lbl_Select_Year_Combobox.TabIndex = 45
        Me.lbl_Select_Year_Combobox.Text = "Select Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmb_URCS_Year
        '
        Me.cmb_URCS_Year.FormattingEnabled = True
        Me.cmb_URCS_Year.Location = New System.Drawing.Point(300, 68)
        Me.cmb_URCS_Year.Name = "cmb_URCS_Year"
        Me.cmb_URCS_Year.Size = New System.Drawing.Size(108, 21)
        Me.cmb_URCS_Year.TabIndex = 44
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(252, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(140, 46)
        Me.Label1.TabIndex = 43
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "UMF Load Module"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.Enabled = False
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(103, 296)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(418, 14)
        Me.txt_StatusBox.TabIndex = 50
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Output_File_Entry
        '
        Me.btn_Output_File_Entry.Location = New System.Drawing.Point(43, 107)
        Me.btn_Output_File_Entry.Name = "btn_Output_File_Entry"
        Me.btn_Output_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Output_File_Entry.TabIndex = 49
        Me.btn_Output_File_Entry.TabStop = False
        Me.btn_Output_File_Entry.Text = "Select Output File:"
        Me.btn_Output_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Output_FilePath
        '
        Me.txt_Output_FilePath.Location = New System.Drawing.Point(154, 107)
        Me.txt_Output_FilePath.Name = "txt_Output_FilePath"
        Me.txt_Output_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Output_FilePath.TabIndex = 48
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(271, 134)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(108, 27)
        Me.btn_Execute.TabIndex = 47
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txt_Processing_Year)
        Me.GroupBox1.Controls.Add(Me.chk_Created_Blank_Owner_Records)
        Me.GroupBox1.Controls.Add(Me.chk_Created_System_Records)
        Me.GroupBox1.Controls.Add(Me.chk_Trans_Table_Update)
        Me.GroupBox1.Location = New System.Drawing.Point(103, 167)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(383, 126)
        Me.GroupBox1.TabIndex = 52
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Processing Progress:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(99, 91)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(188, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Processing Railroad Records for Year:"
        '
        'txt_Processing_Year
        '
        Me.txt_Processing_Year.Enabled = False
        Me.txt_Processing_Year.Location = New System.Drawing.Point(293, 88)
        Me.txt_Processing_Year.Name = "txt_Processing_Year"
        Me.txt_Processing_Year.Size = New System.Drawing.Size(64, 21)
        Me.txt_Processing_Year.TabIndex = 3
        '
        'chk_Created_Blank_Owner_Records
        '
        Me.chk_Created_Blank_Owner_Records.AutoSize = True
        Me.chk_Created_Blank_Owner_Records.Location = New System.Drawing.Point(111, 66)
        Me.chk_Created_Blank_Owner_Records.Name = "chk_Created_Blank_Owner_Records"
        Me.chk_Created_Blank_Owner_Records.Size = New System.Drawing.Size(170, 17)
        Me.chk_Created_Blank_Owner_Records.TabIndex = 2
        Me.chk_Created_Blank_Owner_Records.Text = "Created Blank Owner Records"
        Me.chk_Created_Blank_Owner_Records.UseVisualStyleBackColor = True
        '
        'chk_Created_System_Records
        '
        Me.chk_Created_System_Records.AutoSize = True
        Me.chk_Created_System_Records.Location = New System.Drawing.Point(111, 43)
        Me.chk_Created_System_Records.Name = "chk_Created_System_Records"
        Me.chk_Created_System_Records.Size = New System.Drawing.Size(145, 17)
        Me.chk_Created_System_Records.TabIndex = 1
        Me.chk_Created_System_Records.Text = "Created System Records"
        Me.chk_Created_System_Records.UseVisualStyleBackColor = True
        '
        'chk_Trans_Table_Update
        '
        Me.chk_Trans_Table_Update.AutoSize = True
        Me.chk_Trans_Table_Update.Location = New System.Drawing.Point(111, 20)
        Me.chk_Trans_Table_Update.Name = "chk_Trans_Table_Update"
        Me.chk_Trans_Table_Update.Size = New System.Drawing.Size(178, 17)
        Me.chk_Trans_Table_Update.TabIndex = 0
        Me.chk_Trans_Table_Update.Text = "Performed Trans Table Updates"
        Me.chk_Trans_Table_Update.UseVisualStyleBackColor = True
        '
        'btn_Return_To_DataPrepMenu
        '
        Me.btn_Return_To_DataPrepMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_DataPrepMenu.Location = New System.Drawing.Point(523, 319)
        Me.btn_Return_To_DataPrepMenu.Name = "btn_Return_To_DataPrepMenu"
        Me.btn_Return_To_DataPrepMenu.Size = New System.Drawing.Size(55, 52)
        Me.btn_Return_To_DataPrepMenu.TabIndex = 46
        Me.btn_Return_To_DataPrepMenu.UseVisualStyleBackColor = True
        '
        'UMF_Load_Legacy
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(590, 383)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Output_File_Entry)
        Me.Controls.Add(Me.txt_Output_FilePath)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.btn_Return_To_DataPrepMenu)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.cmb_URCS_Year)
        Me.Controls.Add(Me.Label1)
        Me.Name = "UMF_Load_Legacy"
        Me.Text = "UMF_Load_Legacy"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbl_Select_Year_Combobox As System.Windows.Forms.Label
    Friend WithEvents cmb_URCS_Year As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_Return_To_DataPrepMenu As System.Windows.Forms.Button
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents btn_Output_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Output_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txt_Processing_Year As System.Windows.Forms.TextBox
    Friend WithEvents chk_Created_Blank_Owner_Records As System.Windows.Forms.CheckBox
    Friend WithEvents chk_Created_System_Records As System.Windows.Forms.CheckBox
    Friend WithEvents chk_Trans_Table_Update As System.Windows.Forms.CheckBox
End Class
