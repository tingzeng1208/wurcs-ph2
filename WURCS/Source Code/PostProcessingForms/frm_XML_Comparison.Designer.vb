<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_XML_Comparison
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
        Me.btn_Test_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Test_FilePath = New System.Windows.Forms.TextBox()
        Me.txt_StatusBox = New System.Windows.Forms.TextBox()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_Return_To_Menu = New System.Windows.Forms.Button()
        Me.btn_Base_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Base_FilePath = New System.Windows.Forms.TextBox()
        Me.btn_Output_File = New System.Windows.Forms.Button()
        Me.txt_Output_FilePath = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btn_Test_File_Entry
        '
        Me.btn_Test_File_Entry.Location = New System.Drawing.Point(26, 80)
        Me.btn_Test_File_Entry.Name = "btn_Test_File_Entry"
        Me.btn_Test_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Test_File_Entry.TabIndex = 89
        Me.btn_Test_File_Entry.TabStop = False
        Me.btn_Test_File_Entry.Text = "Select Test File:"
        Me.btn_Test_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Test_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Test_FilePath
        '
        Me.txt_Test_FilePath.Location = New System.Drawing.Point(137, 80)
        Me.txt_Test_FilePath.Name = "txt_Test_FilePath"
        Me.txt_Test_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Test_FilePath.TabIndex = 88
        '
        'txt_StatusBox
        '
        Me.txt_StatusBox.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_StatusBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_StatusBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_StatusBox.Location = New System.Drawing.Point(72, 164)
        Me.txt_StatusBox.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_StatusBox.Name = "txt_StatusBox"
        Me.txt_StatusBox.Size = New System.Drawing.Size(432, 14)
        Me.txt_StatusBox.TabIndex = 87
        Me.txt_StatusBox.TabStop = False
        Me.txt_StatusBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(228, 187)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(104, 27)
        Me.btn_Execute.TabIndex = 86
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(166, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(223, 46)
        Me.Label1.TabIndex = 85
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "XML Data Comparison Module"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btn_Return_To_Menu
        '
        Me.btn_Return_To_Menu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_Menu.Location = New System.Drawing.Point(476, 187)
        Me.btn_Return_To_Menu.Name = "btn_Return_To_Menu"
        Me.btn_Return_To_Menu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_Menu.TabIndex = 84
        Me.btn_Return_To_Menu.UseVisualStyleBackColor = True
        '
        'btn_Base_File_Entry
        '
        Me.btn_Base_File_Entry.Location = New System.Drawing.Point(26, 107)
        Me.btn_Base_File_Entry.Name = "btn_Base_File_Entry"
        Me.btn_Base_File_Entry.Size = New System.Drawing.Size(104, 21)
        Me.btn_Base_File_Entry.TabIndex = 91
        Me.btn_Base_File_Entry.TabStop = False
        Me.btn_Base_File_Entry.Text = "Select Base File:"
        Me.btn_Base_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Base_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Base_FilePath
        '
        Me.txt_Base_FilePath.Location = New System.Drawing.Point(137, 107)
        Me.txt_Base_FilePath.Name = "txt_Base_FilePath"
        Me.txt_Base_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Base_FilePath.TabIndex = 90
        '
        'btn_Output_File
        '
        Me.btn_Output_File.Location = New System.Drawing.Point(26, 134)
        Me.btn_Output_File.Name = "btn_Output_File"
        Me.btn_Output_File.Size = New System.Drawing.Size(104, 21)
        Me.btn_Output_File.TabIndex = 93
        Me.btn_Output_File.TabStop = False
        Me.btn_Output_File.Text = "Select Output File:"
        Me.btn_Output_File.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_File.UseVisualStyleBackColor = True
        '
        'txt_Output_FilePath
        '
        Me.txt_Output_FilePath.Location = New System.Drawing.Point(137, 134)
        Me.txt_Output_FilePath.Name = "txt_Output_FilePath"
        Me.txt_Output_FilePath.Size = New System.Drawing.Size(367, 21)
        Me.txt_Output_FilePath.TabIndex = 92
        '
        'frm_XML_Comparison
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(539, 252)
        Me.ControlBox = False
        Me.Controls.Add(Me.btn_Output_File)
        Me.Controls.Add(Me.txt_Output_FilePath)
        Me.Controls.Add(Me.btn_Base_File_Entry)
        Me.Controls.Add(Me.txt_Base_FilePath)
        Me.Controls.Add(Me.btn_Test_File_Entry)
        Me.Controls.Add(Me.txt_Test_FilePath)
        Me.Controls.Add(Me.txt_StatusBox)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn_Return_To_Menu)
        Me.Name = "frm_XML_Comparison"
        Me.Text = "frm_XML_Comparison"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Test_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Test_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents txt_StatusBox As System.Windows.Forms.TextBox
    Friend WithEvents btn_Execute As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_Return_To_Menu As System.Windows.Forms.Button
    Friend WithEvents btn_Base_File_Entry As System.Windows.Forms.Button
    Friend WithEvents txt_Base_FilePath As System.Windows.Forms.TextBox
    Friend WithEvents btn_Output_File As System.Windows.Forms.Button
    Friend WithEvents txt_Output_FilePath As System.Windows.Forms.TextBox
End Class
