<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Supervisor
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
        Me.Top_label = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.btAgain = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Top_label
        '
        Me.Top_label.AutoSize = True
        Me.Top_label.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Top_label.ForeColor = System.Drawing.Color.White
        Me.Top_label.Location = New System.Drawing.Point(85, 27)
        Me.Top_label.Name = "Top_label"
        Me.Top_label.Size = New System.Drawing.Size(208, 25)
        Me.Top_label.TabIndex = 0
        Me.Top_label.Text = "Supervisor Access"
        '
        'txtPassword
        '
        Me.txtPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassword.Location = New System.Drawing.Point(90, 55)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(196, 24)
        Me.txtPassword.TabIndex = 2
        Me.txtPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btAgain
        '
        Me.btAgain.BackColor = System.Drawing.Color.LimeGreen
        Me.btAgain.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btAgain.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btAgain.Location = New System.Drawing.Point(144, 95)
        Me.btAgain.Name = "btAgain"
        Me.btAgain.Size = New System.Drawing.Size(83, 23)
        Me.btAgain.TabIndex = 3
        Me.btAgain.Text = "Enter"
        Me.btAgain.UseVisualStyleBackColor = False
        '
        'Supervisor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(367, 181)
        Me.Controls.Add(Me.btAgain)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.Top_label)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Supervisor"
        Me.Text = "Supervisor password"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Top_label As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents btAgain As System.Windows.Forms.Button
End Class
