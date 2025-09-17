<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class STCC_Code_Table_Loader
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
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.txt_Input_FilePath = New System.Windows.Forms.TextBox()
        Me.lblR1_Data_Load = New System.Windows.Forms.Label()
        Me.btn_Input_File_Entry = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btn_Return_To_DataPrepMenu
        '
        Me.btn_Return_To_DataPrepMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_DataPrepMenu.Location = New System.Drawing.Point(465, 198)
        Me.btn_Return_To_DataPrepMenu.Name = "btn_Return_To_DataPrepMenu"
        Me.btn_Return_To_DataPrepMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_DataPrepMenu.TabIndex = 49
        Me.btn_Return_To_DataPrepMenu.UseVisualStyleBackColor = True
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(212, 132)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(104, 27)
        Me.btn_Execute.TabIndex = 48
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(48, 109)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 47
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txt_Input_FilePath
        '
        Me.txt_Input_FilePath.Location = New System.Drawing.Point(136, 79)
        Me.txt_Input_FilePath.Name = "txt_Input_FilePath"
        Me.txt_Input_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Input_FilePath.TabIndex = 45
        '
        'lblR1_Data_Load
        '
        Me.lblR1_Data_Load.AutoSize = True
        Me.lblR1_Data_Load.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.lblR1_Data_Load.Location = New System.Drawing.Point(166, 20)
        Me.lblR1_Data_Load.Name = "lblR1_Data_Load"
        Me.lblR1_Data_Load.Size = New System.Drawing.Size(197, 46)
        Me.lblR1_Data_Load.TabIndex = 44
        Me.lblR1_Data_Load.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "STCC Code Load Program"
        Me.lblR1_Data_Load.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btn_Input_File_Entry
        '
        Me.btn_Input_File_Entry.Location = New System.Drawing.Point(26, 79)
        Me.btn_Input_File_Entry.Name = "btn_Input_File_Entry"
        Me.btn_Input_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Input_File_Entry.TabIndex = 50
        Me.btn_Input_File_Entry.TabStop = False
        Me.btn_Input_File_Entry.Text = "Select Input File:"
        Me.btn_Input_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Input_File_Entry.UseVisualStyleBackColor = True
        '
        'STCC_Code_Table_Loader
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(528, 262)
        Me.Controls.Add(Me.btn_Input_File_Entry)
        Me.Controls.Add(Me.btn_Return_To_DataPrepMenu)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.txt_Input_FilePath)
        Me.Controls.Add(Me.lblR1_Data_Load)
        Me.Name = "STCC_Code_Table_Loader"
        Me.Text = "STCC Code Table Loader"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Return_To_DataPrepMenu As System.Windows.Forms.Button
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents txt_Input_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents lblR1_Data_Load As System.Windows.Forms.Label
    Friend WithEvents btn_Input_File_Entry As System.Windows.Forms.Button
End Class
