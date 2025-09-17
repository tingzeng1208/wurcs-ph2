<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPhase3Main
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
        Me.btn_Return_To_MainMenu = New System.Windows.Forms.Button()
        Me.lblYear = New System.Windows.Forms.Label()
        Me.btnSelectOutputFolder = New System.Windows.Forms.Button()
        Me.txtFolder = New System.Windows.Forms.TextBox()
        Me.cmb_URCS_Year = New System.Windows.Forms.ComboBox()
        Me.txtP3FilePath = New System.Windows.Forms.TextBox()
        Me.btnSelectP3File = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chk100Series = New System.Windows.Forms.CheckBox()
        Me.chk200Series = New System.Windows.Forms.CheckBox()
        Me.chk300Series = New System.Windows.Forms.CheckBox()
        Me.chk400Series = New System.Windows.Forms.CheckBox()
        Me.chk500Series = New System.Windows.Forms.CheckBox()
        Me.btnExecute = New System.Windows.Forms.Button()
        Me.txtStatus = New System.Windows.Forms.TextBox()
        Me.chk600Series = New System.Windows.Forms.CheckBox()
        Me.rdo_Legacy = New System.Windows.Forms.RadioButton()
        Me.rdo_Current = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chk_Cost_All_Segments = New System.Windows.Forms.CheckBox()
        Me.btnSelectMWFFile = New System.Windows.Forms.Button()
        Me.txtMakeWholeFilePath = New System.Windows.Forms.TextBox()
        Me.chk_Save_CRPRESRecords = New System.Windows.Forms.CheckBox()
        Me.chkSaveResults = New System.Windows.Forms.CheckBox()
        Me.chk_UseDifferentYear = New System.Windows.Forms.CheckBox()
        Me.cmb_Different_Year = New System.Windows.Forms.ComboBox()
        Me.chk_SaveToSQL = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Chk_SkipCostSegments = New System.Windows.Forms.CheckBox()
        Me.Chk_Skip_MakeWhole = New System.Windows.Forms.CheckBox()
        Me.Chk_UpdateWaybills = New System.Windows.Forms.CheckBox()
        Me.txt_Target_Server_Name = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_Return_To_MainMenu
        '
        Me.btn_Return_To_MainMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_MainMenu.Location = New System.Drawing.Point(904, 710)
        Me.btn_Return_To_MainMenu.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_Return_To_MainMenu.Name = "btn_Return_To_MainMenu"
        Me.btn_Return_To_MainMenu.Size = New System.Drawing.Size(76, 76)
        Me.btn_Return_To_MainMenu.TabIndex = 14
        Me.btn_Return_To_MainMenu.UseVisualStyleBackColor = True
        '
        'lblYear
        '
        Me.lblYear.AutoSize = True
        Me.lblYear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblYear.Location = New System.Drawing.Point(52, 101)
        Me.lblYear.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(123, 19)
        Me.lblYear.TabIndex = 0
        Me.lblYear.Text = "Select Cost Year"
        '
        'btnSelectOutputFolder
        '
        Me.btnSelectOutputFolder.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSelectOutputFolder.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSelectOutputFolder.Location = New System.Drawing.Point(50, 227)
        Me.btnSelectOutputFolder.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSelectOutputFolder.Name = "btnSelectOutputFolder"
        Me.btnSelectOutputFolder.Size = New System.Drawing.Size(297, 34)
        Me.btnSelectOutputFolder.TabIndex = 4
        Me.btnSelectOutputFolder.Text = "Select Output Folder:"
        Me.btnSelectOutputFolder.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnSelectOutputFolder.UseVisualStyleBackColor = True
        '
        'txtFolder
        '
        Me.txtFolder.Location = New System.Drawing.Point(363, 229)
        Me.txtFolder.Margin = New System.Windows.Forms.Padding(4)
        Me.txtFolder.Name = "txtFolder"
        Me.txtFolder.ReadOnly = True
        Me.txtFolder.Size = New System.Drawing.Size(589, 27)
        Me.txtFolder.TabIndex = 5
        Me.txtFolder.TabStop = False
        '
        'cmb_URCS_Year
        '
        Me.cmb_URCS_Year.Location = New System.Drawing.Point(186, 96)
        Me.cmb_URCS_Year.Margin = New System.Windows.Forms.Padding(4)
        Me.cmb_URCS_Year.MaxLength = 4
        Me.cmb_URCS_Year.Name = "cmb_URCS_Year"
        Me.cmb_URCS_Year.Size = New System.Drawing.Size(88, 27)
        Me.cmb_URCS_Year.TabIndex = 1
        '
        'txtP3FilePath
        '
        Me.txtP3FilePath.Location = New System.Drawing.Point(363, 145)
        Me.txtP3FilePath.Margin = New System.Windows.Forms.Padding(4)
        Me.txtP3FilePath.Name = "txtP3FilePath"
        Me.txtP3FilePath.ReadOnly = True
        Me.txtP3FilePath.Size = New System.Drawing.Size(602, 27)
        Me.txtP3FilePath.TabIndex = 3
        Me.txtP3FilePath.TabStop = False
        '
        'btnSelectP3File
        '
        Me.btnSelectP3File.Location = New System.Drawing.Point(50, 142)
        Me.btnSelectP3File.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSelectP3File.Name = "btnSelectP3File"
        Me.btnSelectP3File.Size = New System.Drawing.Size(297, 34)
        Me.btnSelectP3File.TabIndex = 2
        Me.btnSelectP3File.Text = "Select Phase III Spreadsheet File:"
        Me.btnSelectP3File.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnSelectP3File.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(368, 13)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(279, 66)
        Me.Label1.TabIndex = 50
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Phase III Costing Module"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'chk100Series
        '
        Me.chk100Series.AutoSize = True
        Me.chk100Series.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk100Series.Location = New System.Drawing.Point(9, 29)
        Me.chk100Series.Margin = New System.Windows.Forms.Padding(4)
        Me.chk100Series.Name = "chk100Series"
        Me.chk100Series.Size = New System.Drawing.Size(195, 25)
        Me.chk100Series.TabIndex = 6
        Me.chk100Series.Text = "Save L100 Series CSV"
        Me.chk100Series.UseVisualStyleBackColor = True
        '
        'chk200Series
        '
        Me.chk200Series.AutoSize = True
        Me.chk200Series.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk200Series.Location = New System.Drawing.Point(9, 63)
        Me.chk200Series.Margin = New System.Windows.Forms.Padding(4)
        Me.chk200Series.Name = "chk200Series"
        Me.chk200Series.Size = New System.Drawing.Size(195, 25)
        Me.chk200Series.TabIndex = 7
        Me.chk200Series.Text = "Save L200 Series CSV"
        Me.chk200Series.UseVisualStyleBackColor = True
        '
        'chk300Series
        '
        Me.chk300Series.AutoSize = True
        Me.chk300Series.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk300Series.Location = New System.Drawing.Point(9, 96)
        Me.chk300Series.Margin = New System.Windows.Forms.Padding(4)
        Me.chk300Series.Name = "chk300Series"
        Me.chk300Series.Size = New System.Drawing.Size(195, 25)
        Me.chk300Series.TabIndex = 8
        Me.chk300Series.Text = "Save L300 Series CSV"
        Me.chk300Series.UseVisualStyleBackColor = True
        '
        'chk400Series
        '
        Me.chk400Series.AutoSize = True
        Me.chk400Series.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk400Series.Location = New System.Drawing.Point(9, 130)
        Me.chk400Series.Margin = New System.Windows.Forms.Padding(4)
        Me.chk400Series.Name = "chk400Series"
        Me.chk400Series.Size = New System.Drawing.Size(195, 25)
        Me.chk400Series.TabIndex = 9
        Me.chk400Series.Text = "Save L400 Series CSV"
        Me.chk400Series.UseVisualStyleBackColor = True
        '
        'chk500Series
        '
        Me.chk500Series.AutoSize = True
        Me.chk500Series.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk500Series.Location = New System.Drawing.Point(9, 165)
        Me.chk500Series.Margin = New System.Windows.Forms.Padding(4)
        Me.chk500Series.Name = "chk500Series"
        Me.chk500Series.Size = New System.Drawing.Size(195, 25)
        Me.chk500Series.TabIndex = 10
        Me.chk500Series.Text = "Save L500 Series CSV"
        Me.chk500Series.UseVisualStyleBackColor = True
        '
        'btnExecute
        '
        Me.btnExecute.Location = New System.Drawing.Point(368, 694)
        Me.btnExecute.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExecute.Name = "btnExecute"
        Me.btnExecute.Size = New System.Drawing.Size(176, 42)
        Me.btnExecute.TabIndex = 13
        Me.btnExecute.Text = "Execute"
        Me.btnExecute.UseVisualStyleBackColor = True
        '
        'txtStatus
        '
        Me.txtStatus.BackColor = System.Drawing.SystemColors.Control
        Me.txtStatus.Location = New System.Drawing.Point(63, 655)
        Me.txtStatus.Margin = New System.Windows.Forms.Padding(4)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(788, 27)
        Me.txtStatus.TabIndex = 52
        Me.txtStatus.TabStop = False
        '
        'chk600Series
        '
        Me.chk600Series.AutoSize = True
        Me.chk600Series.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk600Series.Location = New System.Drawing.Point(9, 199)
        Me.chk600Series.Margin = New System.Windows.Forms.Padding(4)
        Me.chk600Series.Name = "chk600Series"
        Me.chk600Series.Size = New System.Drawing.Size(195, 25)
        Me.chk600Series.TabIndex = 11
        Me.chk600Series.Text = "Save L600 Series CSV"
        Me.chk600Series.UseVisualStyleBackColor = True
        '
        'rdo_Legacy
        '
        Me.rdo_Legacy.AutoSize = True
        Me.rdo_Legacy.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdo_Legacy.Location = New System.Drawing.Point(9, 29)
        Me.rdo_Legacy.Margin = New System.Windows.Forms.Padding(4)
        Me.rdo_Legacy.Name = "rdo_Legacy"
        Me.rdo_Legacy.Size = New System.Drawing.Size(86, 25)
        Me.rdo_Legacy.TabIndex = 0
        Me.rdo_Legacy.TabStop = True
        Me.rdo_Legacy.Text = "Legacy"
        Me.rdo_Legacy.UseVisualStyleBackColor = True
        '
        'rdo_Current
        '
        Me.rdo_Current.AutoSize = True
        Me.rdo_Current.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdo_Current.Location = New System.Drawing.Point(9, 56)
        Me.rdo_Current.Margin = New System.Windows.Forms.Padding(4)
        Me.rdo_Current.Name = "rdo_Current"
        Me.rdo_Current.Size = New System.Drawing.Size(90, 25)
        Me.rdo_Current.TabIndex = 1
        Me.rdo_Current.TabStop = True
        Me.rdo_Current.Text = "Current"
        Me.rdo_Current.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdo_Legacy)
        Me.GroupBox1.Controls.Add(Me.rdo_Current)
        Me.GroupBox1.Controls.Add(Me.chk_Cost_All_Segments)
        Me.GroupBox1.Font = New System.Drawing.Font("Tahoma", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(57, 336)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(318, 126)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Cost Using Waybill Input Logic"
        '
        'chk_Cost_All_Segments
        '
        Me.chk_Cost_All_Segments.AutoSize = True
        Me.chk_Cost_All_Segments.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_Cost_All_Segments.Location = New System.Drawing.Point(36, 89)
        Me.chk_Cost_All_Segments.Margin = New System.Windows.Forms.Padding(4)
        Me.chk_Cost_All_Segments.Name = "chk_Cost_All_Segments"
        Me.chk_Cost_All_Segments.Size = New System.Drawing.Size(282, 25)
        Me.chk_Cost_All_Segments.TabIndex = 59
        Me.chk_Cost_All_Segments.Text = "Cost All Segments (US, CA, MX)?"
        Me.chk_Cost_All_Segments.UseVisualStyleBackColor = True
        '
        'btnSelectMWFFile
        '
        Me.btnSelectMWFFile.Location = New System.Drawing.Point(50, 184)
        Me.btnSelectMWFFile.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSelectMWFFile.Name = "btnSelectMWFFile"
        Me.btnSelectMWFFile.Size = New System.Drawing.Size(297, 34)
        Me.btnSelectMWFFile.TabIndex = 53
        Me.btnSelectMWFFile.Text = "Select Make Whole Spreadsheet File:"
        Me.btnSelectMWFFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnSelectMWFFile.UseVisualStyleBackColor = True
        '
        'txtMakeWholeFilePath
        '
        Me.txtMakeWholeFilePath.Location = New System.Drawing.Point(363, 184)
        Me.txtMakeWholeFilePath.Margin = New System.Windows.Forms.Padding(4)
        Me.txtMakeWholeFilePath.Name = "txtMakeWholeFilePath"
        Me.txtMakeWholeFilePath.ReadOnly = True
        Me.txtMakeWholeFilePath.Size = New System.Drawing.Size(602, 27)
        Me.txtMakeWholeFilePath.TabIndex = 54
        Me.txtMakeWholeFilePath.TabStop = False
        '
        'chk_Save_CRPRESRecords
        '
        Me.chk_Save_CRPRESRecords.AutoSize = True
        Me.chk_Save_CRPRESRecords.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_Save_CRPRESRecords.Location = New System.Drawing.Point(8, 266)
        Me.chk_Save_CRPRESRecords.Margin = New System.Windows.Forms.Padding(4)
        Me.chk_Save_CRPRESRecords.Name = "chk_Save_CRPRESRecords"
        Me.chk_Save_CRPRESRecords.Size = New System.Drawing.Size(235, 25)
        Me.chk_Save_CRPRESRecords.TabIndex = 55
        Me.chk_Save_CRPRESRecords.Text = "Save CRPRES Records CSV"
        Me.chk_Save_CRPRESRecords.UseVisualStyleBackColor = True
        '
        'chkSaveResults
        '
        Me.chkSaveResults.AutoSize = True
        Me.chkSaveResults.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSaveResults.Location = New System.Drawing.Point(8, 232)
        Me.chkSaveResults.Margin = New System.Windows.Forms.Padding(4)
        Me.chkSaveResults.Name = "chkSaveResults"
        Me.chkSaveResults.Size = New System.Drawing.Size(165, 25)
        Me.chkSaveResults.TabIndex = 56
        Me.chkSaveResults.Text = "Save Results CSV"
        Me.chkSaveResults.UseVisualStyleBackColor = True
        '
        'chk_UseDifferentYear
        '
        Me.chk_UseDifferentYear.AutoSize = True
        Me.chk_UseDifferentYear.Location = New System.Drawing.Point(57, 282)
        Me.chk_UseDifferentYear.Margin = New System.Windows.Forms.Padding(4)
        Me.chk_UseDifferentYear.Name = "chk_UseDifferentYear"
        Me.chk_UseDifferentYear.Size = New System.Drawing.Size(265, 23)
        Me.chk_UseDifferentYear.TabIndex = 60
        Me.chk_UseDifferentYear.Text = "Use Different Waybill Year Data?"
        Me.chk_UseDifferentYear.UseVisualStyleBackColor = True
        '
        'cmb_Different_Year
        '
        Me.cmb_Different_Year.FormattingEnabled = True
        Me.cmb_Different_Year.Location = New System.Drawing.Point(363, 276)
        Me.cmb_Different_Year.Margin = New System.Windows.Forms.Padding(4)
        Me.cmb_Different_Year.Name = "cmb_Different_Year"
        Me.cmb_Different_Year.Size = New System.Drawing.Size(109, 27)
        Me.cmb_Different_Year.TabIndex = 61
        Me.cmb_Different_Year.Visible = False
        '
        'chk_SaveToSQL
        '
        Me.chk_SaveToSQL.AutoSize = True
        Me.chk_SaveToSQL.Location = New System.Drawing.Point(86, 507)
        Me.chk_SaveToSQL.Margin = New System.Windows.Forms.Padding(4)
        Me.chk_SaveToSQL.Name = "chk_SaveToSQL"
        Me.chk_SaveToSQL.Size = New System.Drawing.Size(121, 23)
        Me.chk_SaveToSQL.TabIndex = 62
        Me.chk_SaveToSQL.Text = "Save to SQL"
        Me.chk_SaveToSQL.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chk100Series)
        Me.GroupBox2.Controls.Add(Me.chk200Series)
        Me.GroupBox2.Controls.Add(Me.chk300Series)
        Me.GroupBox2.Controls.Add(Me.chk400Series)
        Me.GroupBox2.Controls.Add(Me.chkSaveResults)
        Me.GroupBox2.Controls.Add(Me.chk500Series)
        Me.GroupBox2.Controls.Add(Me.chk_Save_CRPRESRecords)
        Me.GroupBox2.Controls.Add(Me.chk600Series)
        Me.GroupBox2.Font = New System.Drawing.Font("Tahoma", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(438, 336)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Size = New System.Drawing.Size(314, 300)
        Me.GroupBox2.TabIndex = 63
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Intermediate Output File Options"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(62, 483)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(152, 21)
        Me.Label2.TabIndex = 64
        Me.Label2.Text = "SQL Data Option"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(62, 536)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(116, 21)
        Me.Label3.TabIndex = 65
        Me.Label3.Text = "Run Options"
        '
        'Chk_SkipCostSegments
        '
        Me.Chk_SkipCostSegments.AutoSize = True
        Me.Chk_SkipCostSegments.Location = New System.Drawing.Point(86, 561)
        Me.Chk_SkipCostSegments.Margin = New System.Windows.Forms.Padding(4)
        Me.Chk_SkipCostSegments.Name = "Chk_SkipCostSegments"
        Me.Chk_SkipCostSegments.Size = New System.Drawing.Size(249, 23)
        Me.Chk_SkipCostSegments.TabIndex = 66
        Me.Chk_SkipCostSegments.Text = "Skip 1st Pass (Cost Segments)"
        Me.Chk_SkipCostSegments.UseVisualStyleBackColor = True
        '
        'Chk_Skip_MakeWhole
        '
        Me.Chk_Skip_MakeWhole.AutoSize = True
        Me.Chk_Skip_MakeWhole.Location = New System.Drawing.Point(86, 592)
        Me.Chk_Skip_MakeWhole.Margin = New System.Windows.Forms.Padding(4)
        Me.Chk_Skip_MakeWhole.Name = "Chk_Skip_MakeWhole"
        Me.Chk_Skip_MakeWhole.Size = New System.Drawing.Size(346, 23)
        Me.Chk_Skip_MakeWhole.TabIndex = 67
        Me.Chk_Skip_MakeWhole.Text = "Skip 2nd Pass (Update Make-Whole Factors)"
        Me.Chk_Skip_MakeWhole.UseVisualStyleBackColor = True
        '
        'Chk_UpdateWaybills
        '
        Me.Chk_UpdateWaybills.AutoSize = True
        Me.Chk_UpdateWaybills.Location = New System.Drawing.Point(86, 623)
        Me.Chk_UpdateWaybills.Margin = New System.Windows.Forms.Padding(4)
        Me.Chk_UpdateWaybills.Name = "Chk_UpdateWaybills"
        Me.Chk_UpdateWaybills.Size = New System.Drawing.Size(317, 23)
        Me.Chk_UpdateWaybills.TabIndex = 68
        Me.Chk_UpdateWaybills.Text = "Skip 3rd Pass (Update Masked Waybills)"
        Me.Chk_UpdateWaybills.UseVisualStyleBackColor = True
        '
        'txt_Target_Server_Name
        '
        Me.txt_Target_Server_Name.Location = New System.Drawing.Point(528, 98)
        Me.txt_Target_Server_Name.Margin = New System.Windows.Forms.Padding(4)
        Me.txt_Target_Server_Name.Name = "txt_Target_Server_Name"
        Me.txt_Target_Server_Name.ReadOnly = True
        Me.txt_Target_Server_Name.Size = New System.Drawing.Size(256, 27)
        Me.txt_Target_Server_Name.TabIndex = 69
        Me.txt_Target_Server_Name.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(359, 101)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(161, 19)
        Me.Label4.TabIndex = 70
        Me.Label4.Text = "Run will be updating:"
        '
        'frmPhase3Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(999, 805)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txt_Target_Server_Name)
        Me.Controls.Add(Me.Chk_UpdateWaybills)
        Me.Controls.Add(Me.Chk_Skip_MakeWhole)
        Me.Controls.Add(Me.Chk_SkipCostSegments)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.chk_SaveToSQL)
        Me.Controls.Add(Me.cmb_Different_Year)
        Me.Controls.Add(Me.chk_UseDifferentYear)
        Me.Controls.Add(Me.btnSelectMWFFile)
        Me.Controls.Add(Me.txtMakeWholeFilePath)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.btnExecute)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnSelectP3File)
        Me.Controls.Add(Me.txtP3FilePath)
        Me.Controls.Add(Me.txtFolder)
        Me.Controls.Add(Me.cmb_URCS_Year)
        Me.Controls.Add(Me.btnSelectOutputFolder)
        Me.Controls.Add(Me.btn_Return_To_MainMenu)
        Me.Controls.Add(Me.lblYear)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "frmPhase3Main"
        Me.Text = "URCS Phase III - Waybill Costing"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Return_To_MainMenu As System.Windows.Forms.Button
    Friend WithEvents lblYear As Label
    Friend WithEvents btnSelectOutputFolder As Button
    Friend WithEvents txtFolder As TextBox
    Friend WithEvents cmb_URCS_Year As ComboBox
    Friend WithEvents txtP3FilePath As TextBox
    Friend WithEvents btnSelectP3File As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents chk100Series As CheckBox
    Friend WithEvents chk200Series As CheckBox
    Friend WithEvents chk300Series As CheckBox
    Friend WithEvents chk400Series As CheckBox
    Friend WithEvents chk500Series As CheckBox
    Friend WithEvents btnExecute As Button
    Friend WithEvents txtStatus As TextBox
    Friend WithEvents chk600Series As System.Windows.Forms.CheckBox
    Friend WithEvents rdo_Legacy As RadioButton
    Friend WithEvents rdo_Current As RadioButton
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents btnSelectMWFFile As Button
    Friend WithEvents txtMakeWholeFilePath As TextBox
    Friend WithEvents chk_Save_CRPRESRecords As CheckBox
    Friend WithEvents chkSaveResults As CheckBox
    Friend WithEvents chk_Cost_All_Segments As CheckBox
    Friend WithEvents chk_UseDifferentYear As CheckBox
    Friend WithEvents cmb_Different_Year As ComboBox
    Friend WithEvents chk_SaveToSQL As CheckBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Chk_SkipCostSegments As CheckBox
    Friend WithEvents Chk_Skip_MakeWhole As CheckBox
    Friend WithEvents Chk_UpdateWaybills As CheckBox
    Friend WithEvents txt_Target_Server_Name As TextBox
    Friend WithEvents Label4 As Label
End Class
