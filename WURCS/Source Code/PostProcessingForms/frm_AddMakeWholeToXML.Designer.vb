<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_AddMakeWholeToXML
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
        Me.btn_Input_File_Entry = New System.Windows.Forms.Button()
        Me.btn_Return_To_PostProcessingMenu = New System.Windows.Forms.Button()
        Me.btn_Execute = New System.Windows.Forms.Button()
        Me.btn_Output_File_Entry = New System.Windows.Forms.Button()
        Me.txt_Status = New System.Windows.Forms.TextBox()
        Me.txt_Residual_FilePath = New System.Windows.Forms.TextBox()
        Me.txt_XML_FilePath = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(145, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(301, 46)
        Me.Label1.TabIndex = 39
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Add Make-Whole Factors To XML Module"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btn_Input_File_Entry
        '
        Me.btn_Input_File_Entry.Location = New System.Drawing.Point(39, 97)
        Me.btn_Input_File_Entry.Name = "btn_Input_File_Entry"
        Me.btn_Input_File_Entry.Size = New System.Drawing.Size(110, 21)
        Me.btn_Input_File_Entry.TabIndex = 90
        Me.btn_Input_File_Entry.TabStop = False
        Me.btn_Input_File_Entry.Text = "Select Residual File:"
        Me.btn_Input_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Input_File_Entry.UseVisualStyleBackColor = True
        '
        'btn_Return_To_PostProcessingMenu
        '
        Me.btn_Return_To_PostProcessingMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_PostProcessingMenu.Location = New System.Drawing.Point(516, 197)
        Me.btn_Return_To_PostProcessingMenu.Name = "btn_Return_To_PostProcessingMenu"
        Me.btn_Return_To_PostProcessingMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_PostProcessingMenu.TabIndex = 85
        Me.btn_Return_To_PostProcessingMenu.UseVisualStyleBackColor = True
        '
        'btn_Execute
        '
        Me.btn_Execute.Location = New System.Drawing.Point(240, 169)
        Me.btn_Execute.Name = "btn_Execute"
        Me.btn_Execute.Size = New System.Drawing.Size(104, 27)
        Me.btn_Execute.TabIndex = 84
        Me.btn_Execute.Text = "Execute"
        Me.btn_Execute.UseVisualStyleBackColor = True
        '
        'btn_Output_File_Entry
        '
        Me.btn_Output_File_Entry.Location = New System.Drawing.Point(39, 124)
        Me.btn_Output_File_Entry.Name = "btn_Output_File_Entry"
        Me.btn_Output_File_Entry.Size = New System.Drawing.Size(110, 21)
        Me.btn_Output_File_Entry.TabIndex = 93
        Me.btn_Output_File_Entry.TabStop = False
        Me.btn_Output_File_Entry.Text = "Select XML File:"
        Me.btn_Output_File_Entry.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.btn_Output_File_Entry.UseVisualStyleBackColor = True
        '
        'txt_Status
        '
        Me.txt_Status.BackColor = System.Drawing.SystemColors.Menu
        Me.txt_Status.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_Status.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_Status.Location = New System.Drawing.Point(72, 148)
        Me.txt_Status.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_Status.Name = "txt_Status"
        Me.txt_Status.Size = New System.Drawing.Size(432, 14)
        Me.txt_Status.TabIndex = 91
        Me.txt_Status.TabStop = False
        Me.txt_Status.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txt_Residual_FilePath
        '
        Me.txt_Residual_FilePath.Location = New System.Drawing.Point(155, 97)
        Me.txt_Residual_FilePath.Name = "txt_Residual_FilePath"
        Me.txt_Residual_FilePath.Size = New System.Drawing.Size(412, 21)
        Me.txt_Residual_FilePath.TabIndex = 94
        '
        'txt_XML_FilePath
        '
        Me.txt_XML_FilePath.Location = New System.Drawing.Point(155, 125)
        Me.txt_XML_FilePath.Name = "txt_XML_FilePath"
        Me.txt_XML_FilePath.Size = New System.Drawing.Size(412, 21)
        Me.txt_XML_FilePath.TabIndex = 95
        '
        'frm_AddMakeWholeToXML
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(590, 262)
        Me.Controls.Add(Me.txt_XML_FilePath)
        Me.Controls.Add(Me.txt_Residual_FilePath)
        Me.Controls.Add(Me.btn_Output_File_Entry)
        Me.Controls.Add(Me.txt_Status)
        Me.Controls.Add(Me.btn_Input_File_Entry)
        Me.Controls.Add(Me.btn_Return_To_PostProcessingMenu)
        Me.Controls.Add(Me.btn_Execute)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frm_AddMakeWholeToXML"
        Me.Text = "frm_AddMakeWholeToXML"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents btn_Input_File_Entry As Button
    Friend WithEvents txt_XML_File_Path As TextBox
    Friend WithEvents btn_Return_To_PostProcessingMenu As Button
    Friend WithEvents btn_Execute As Button
    Friend WithEvents btn_Output_File_Entry As Button
    Friend WithEvents txt_Status As TextBox
    Friend WithEvents txt_Residual_FilePath As TextBox
    Friend WithEvents txt_XML_FilePath As TextBox
End Class
