<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_PreProcessing_Checks
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
        Me.btn_Return_To_MainMenu = New System.Windows.Forms.Button()
        Me.btn_Trans_Data_Status_Check = New System.Windows.Forms.Button()
        Me.btn_R1_Variance_Report = New System.Windows.Forms.Button()
        Me.btn_R1_Balance_Check = New System.Windows.Forms.Button()
        Me.btn_Trans_Data_Comparison = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 14.25!)
        Me.Label1.Location = New System.Drawing.Point(112, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(155, 46)
        Me.Label1.TabIndex = 37
        Me.Label1.Text = "URCS && Waybills" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Preprocessing Menu"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btn_Return_To_MainMenu
        '
        Me.btn_Return_To_MainMenu.Image = Global.WURCS.My.Resources.Resources.ExitDoor
        Me.btn_Return_To_MainMenu.Location = New System.Drawing.Point(308, 230)
        Me.btn_Return_To_MainMenu.Name = "btn_Return_To_MainMenu"
        Me.btn_Return_To_MainMenu.Size = New System.Drawing.Size(51, 52)
        Me.btn_Return_To_MainMenu.TabIndex = 38
        Me.btn_Return_To_MainMenu.UseVisualStyleBackColor = True
        '
        'btn_Trans_Data_Status_Check
        '
        Me.btn_Trans_Data_Status_Check.Enabled = False
        Me.btn_Trans_Data_Status_Check.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Trans_Data_Status_Check.Location = New System.Drawing.Point(73, 180)
        Me.btn_Trans_Data_Status_Check.Name = "btn_Trans_Data_Status_Check"
        Me.btn_Trans_Data_Status_Check.Size = New System.Drawing.Size(236, 30)
        Me.btn_Trans_Data_Status_Check.TabIndex = 41
        Me.btn_Trans_Data_Status_Check.Text = "Trans Data Status Check"
        Me.btn_Trans_Data_Status_Check.UseVisualStyleBackColor = True
        '
        'btn_R1_Variance_Report
        '
        Me.btn_R1_Variance_Report.Enabled = False
        Me.btn_R1_Variance_Report.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_R1_Variance_Report.Location = New System.Drawing.Point(73, 108)
        Me.btn_R1_Variance_Report.Name = "btn_R1_Variance_Report"
        Me.btn_R1_Variance_Report.Size = New System.Drawing.Size(236, 30)
        Me.btn_R1_Variance_Report.TabIndex = 40
        Me.btn_R1_Variance_Report.Text = "R-1 Variance Report"
        Me.btn_R1_Variance_Report.UseVisualStyleBackColor = True
        '
        'btn_R1_Balance_Check
        '
        Me.btn_R1_Balance_Check.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_R1_Balance_Check.ForeColor = System.Drawing.Color.Black
        Me.btn_R1_Balance_Check.Location = New System.Drawing.Point(73, 72)
        Me.btn_R1_Balance_Check.Name = "btn_R1_Balance_Check"
        Me.btn_R1_Balance_Check.Size = New System.Drawing.Size(236, 30)
        Me.btn_R1_Balance_Check.TabIndex = 39
        Me.btn_R1_Balance_Check.Text = "R-1 Balance Check"
        Me.btn_R1_Balance_Check.UseVisualStyleBackColor = True
        '
        'btn_Trans_Data_Comparison
        '
        Me.btn_Trans_Data_Comparison.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Trans_Data_Comparison.Location = New System.Drawing.Point(73, 144)
        Me.btn_Trans_Data_Comparison.Name = "btn_Trans_Data_Comparison"
        Me.btn_Trans_Data_Comparison.Size = New System.Drawing.Size(236, 30)
        Me.btn_Trans_Data_Comparison.TabIndex = 42
        Me.btn_Trans_Data_Comparison.Text = "Trans Data Comparison"
        Me.btn_Trans_Data_Comparison.UseVisualStyleBackColor = True
        '
        'frm_PreProcessing_Checks
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(371, 299)
        Me.Controls.Add(Me.btn_Trans_Data_Comparison)
        Me.Controls.Add(Me.btn_Trans_Data_Status_Check)
        Me.Controls.Add(Me.btn_R1_Variance_Report)
        Me.Controls.Add(Me.btn_R1_Balance_Check)
        Me.Controls.Add(Me.btn_Return_To_MainMenu)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frm_PreProcessing_Checks"
        Me.Text = "PreProcessing_Checks"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_Return_To_MainMenu As System.Windows.Forms.Button
    Friend WithEvents btn_Trans_Data_Status_Check As System.Windows.Forms.Button
    Friend WithEvents btn_R1_Variance_Report As System.Windows.Forms.Button
    Friend WithEvents btn_R1_Balance_Check As System.Windows.Forms.Button
    Friend WithEvents btn_Trans_Data_Comparison As System.Windows.Forms.Button
End Class
