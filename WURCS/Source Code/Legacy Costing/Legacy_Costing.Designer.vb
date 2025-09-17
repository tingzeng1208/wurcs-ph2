<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Legacy_Costing
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
        Me.cmb_URCS_Year_Combobox = New System.Windows.Forms.ComboBox()
        Me.lbl_Select_Year_Combobox = New System.Windows.Forms.Label()
        Me.lblCS54_Data_Load = New System.Windows.Forms.Label()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Return_To_MainMenu = New System.Windows.Forms.Button()
        Me.btn_Input_File_Entry = New System.Windows.Forms.Button()
        Me.txt_InputXML_FilePath = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'cmb_URCS_Year_Combobox
        '
        Me.cmb_URCS_Year_Combobox.FormattingEnabled = True
        Me.cmb_URCS_Year_Combobox.Location = New System.Drawing.Point(242, 64)
        Me.cmb_URCS_Year_Combobox.Name = "cmb_URCS_Year_Combobox"
        Me.cmb_URCS_Year_Combobox.Size = New System.Drawing.Size(102, 21)
        Me.cmb_URCS_Year_Combobox.TabIndex = 60
        '
        'lbl_Select_Year_Combobox
        '
        Me.lbl_Select_Year_Combobox.AutoSize = True
        Me.lbl_Select_Year_Combobox.Location = New System.Drawing.Point(165, 65)
        Me.lbl_Select_Year_Combobox.Name = "lbl_Select_Year_Combobox"
        Me.lbl_Select_Year_Combobox.Size = New System.Drawing.Size(62, 13)
        Me.lbl_Select_Year_Combobox.TabIndex = 59
        Me.lbl_Select_Year_Combobox.Text = "Enter Year:"
        Me.lbl_Select_Year_Combobox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCS54_Data_Load
        '
        Me.lblCS54_Data_Load.AutoSize = True
        Me.lblCS54_Data_Load.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.lblCS54_Data_Load.Location = New System.Drawing.Point(164, 9)
        Me.lblCS54_Data_Load.Name = "lblCS54_Data_Load"
        Me.lblCS54_Data_Load.Size = New System.Drawing.Size(181, 46)
        Me.lblCS54_Data_Load.TabIndex = 58
        Me.lblCS54_Data_Load.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Legacy Costing Program"
        Me.lblCS54_Data_Load.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(202, 147)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(104, 27)
        Me.btn_Execute.TabIndex = 87
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(38, 124)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 86
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Return_To_MainMenu
        '
        Me.btn_Return_To_MainMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_MainMenu.Location = New System.Drawing.Point(442, 204)
        Me.btn_Return_To_MainMenu.Name = "btn_Return_To_MainMenu"
        Me.btn_Return_To_MainMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_MainMenu.TabIndex = 85
        Me.btn_Return_To_MainMenu.UseVisualStyleBackColor = True
        '
        'btn_Input_File_Entry
        '
        Me.btn_Input_File_Entry.Location = New System.Drawing.Point(16, 91)
        Me.btn_Input_File_Entry.Name = "btn_Input_File_Entry"
        Me.btn_Input_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Input_File_Entry.TabIndex = 89
        Me.btn_Input_File_Entry.TabStop = False
        Me.btn_Input_File_Entry.Text = "Select Input File:"
        Me.btn_Input_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Input_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_InputXML_FilePath
        '
        Me.txt_InputXML_FilePath.Location = New System.Drawing.Point(126, 91)
        Me.txt_InputXML_FilePath.Name = "txt_InputXML_FilePath"
        Me.txt_InputXML_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_InputXML_FilePath.TabIndex = 88
        '
        'Legacy_Costing
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(508, 268)
        Me.Controls.Add(Me.btn_Input_File_Entry)
        Me.Controls.Add(Me.txt_InputXML_FilePath)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Return_To_MainMenu)
        Me.Controls.Add(Me.cmb_URCS_Year_Combobox)
        Me.Controls.Add(Me.lbl_Select_Year_Combobox)
        Me.Controls.Add(Me.lblCS54_Data_Load)
        Me.Name = "Legacy_Costing"
        Me.Text = "Legacy Costing Module"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmb_URCS_Year_Combobox As ComboBox
    Friend WithEvents lbl_Select_Year_Combobox As Label
    Friend WithEvents lblCS54_Data_Load As Label
    Friend WithEvents btn_Execute As Button
    Friend WithEvents txt_StatusBox As TextBox
    Friend WithEvents btn_Return_To_MainMenu As Button
    Friend WithEvents btn_Input_File_Entry As Button
    Friend WithEvents txt_InputXML_FilePath As TextBox
End Class
