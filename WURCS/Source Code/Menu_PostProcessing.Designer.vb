<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Post_Processing_Menu
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
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.btn_UMF_Data_Load = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_XML_Loader = New System.Windows.Forms.Button()
        Me.btn_Return_To_MainMenu = New System.Windows.Forms.Button()
        Me.btn_FileToFile_Compare = New System.Windows.Forms.Button()
        Me.btn_Productivity = New System.Windows.Forms.Button()
        Me.btn_Add_Make_Whole = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox2
        '
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(73, 233)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(236, 13)
        Me.TextBox2.TabIndex = 33
        Me.TextBox2.Text = "URCS Legacy Support Programs"
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btn_UMF_Data_Load
        '
        Me.btn_UMF_Data_Load.Enabled = False
        Me.btn_UMF_Data_Load.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_UMF_Data_Load.Location = New System.Drawing.Point(73, 252)
        Me.btn_UMF_Data_Load.Name = "btn_UMF_Data_Load"
        Me.btn_UMF_Data_Load.Size = New System.Drawing.Size(236, 31)
        Me.btn_UMF_Data_Load.TabIndex = 32
        Me.btn_UMF_Data_Load.Text = "Update Costs From 570 File"
        Me.btn_UMF_Data_Load.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(108, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(166, 46)
        Me.Label1.TabIndex = 36
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Post Processing Menu"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btn_XML_Loader
        '
        Me.btn_XML_Loader.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_XML_Loader.Location = New System.Drawing.Point(73, 85)
        Me.btn_XML_Loader.Name = "btn_XML_Loader"
        Me.btn_XML_Loader.Size = New System.Drawing.Size(236, 30)
        Me.btn_XML_Loader.TabIndex = 43
        Me.btn_XML_Loader.Text = "XML Data Loader"
        Me.btn_XML_Loader.UseVisualStyleBackColor = True
        '
        'btn_Return_To_MainMenu
        '
        Me.btn_Return_To_MainMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_MainMenu.Location = New System.Drawing.Point(325, 289)
        Me.btn_Return_To_MainMenu.Name = "btn_Return_To_MainMenu"
        Me.btn_Return_To_MainMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_MainMenu.TabIndex = 20
        Me.btn_Return_To_MainMenu.UseVisualStyleBackColor = True
        '
        'btn_FileToFile_Compare
        '
        Me.btn_FileToFile_Compare.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_FileToFile_Compare.Location = New System.Drawing.Point(73, 121)
        Me.btn_FileToFile_Compare.Name = "btn_FileToFile_Compare"
        Me.btn_FileToFile_Compare.Size = New System.Drawing.Size(236, 30)
        Me.btn_FileToFile_Compare.TabIndex = 44
        Me.btn_FileToFile_Compare.Text = "XML File-To-File Compare"
        Me.btn_FileToFile_Compare.UseVisualStyleBackColor = True
        '
        'btn_Productivity
        '
        Me.btn_Productivity.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Productivity.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_Productivity.Location = New System.Drawing.Point(73, 193)
        Me.btn_Productivity.Name = "btn_Productivity"
        Me.btn_Productivity.Size = New System.Drawing.Size(236, 30)
        Me.btn_Productivity.TabIndex = 45
        Me.btn_Productivity.Text = "Productivity"
        Me.btn_Productivity.UseVisualStyleBackColor = True
        '
        'btn_Add_Make_Whole
        '
        Me.btn_Add_Make_Whole.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Add_Make_Whole.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_Add_Make_Whole.Location = New System.Drawing.Point(73, 157)
        Me.btn_Add_Make_Whole.Name = "btn_Add_Make_Whole"
        Me.btn_Add_Make_Whole.Size = New System.Drawing.Size(236, 30)
        Me.btn_Add_Make_Whole.TabIndex = 46
        Me.btn_Add_Make_Whole.Text = "Add Make-Whole Factors to XML"
        Me.btn_Add_Make_Whole.UseVisualStyleBackColor = True
        '
        'frm_Post_Processing_Menu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(388, 394)
        Me.ControlBox = False
        Me.Controls.Add(Me.btn_Add_Make_Whole)
        Me.Controls.Add(Me.btn_Productivity)
        Me.Controls.Add(Me.btn_FileToFile_Compare)
        Me.Controls.Add(Me.btn_XML_Loader)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.btn_UMF_Data_Load)
        Me.Controls.Add(Me.btn_Return_To_MainMenu)
        Me.Name = "frm_Post_Processing_Menu"
        Me.Text = "URCS Post Processing Menu"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Return_To_MainMenu As System.Windows.Forms.Button
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents btn_UMF_Data_Load As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_XML_Loader As System.Windows.Forms.Button
    Friend WithEvents btn_FileToFile_Compare As System.Windows.Forms.Button
    Friend WithEvents btn_Productivity As System.Windows.Forms.Button
    Friend WithEvents btn_Add_Make_Whole As Button
End Class
