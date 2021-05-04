<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PortConfigView
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PortConfigView))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.cmbExecute = New System.Windows.Forms.Button()
        Me.cmbSelect = New System.Windows.Forms.ComboBox()
        Me.cmbPARM1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbPARM2 = New System.Windows.Forms.ComboBox()
        Me.E1 = New System.Windows.Forms.Label()
        Me.txtPARM1 = New System.Windows.Forms.TextBox()
        Me.A1 = New System.Windows.Forms.Label()
        Me.txtPARM2 = New System.Windows.Forms.TextBox()
        Me.E2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtPARM3 = New System.Windows.Forms.TextBox()
        Me.E3 = New System.Windows.Forms.Label()
        Me.A2 = New System.Windows.Forms.Label()
        Me.cmbPARM3 = New System.Windows.Forms.ComboBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(1, 78)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1310, 497)
        Me.DataGridView1.TabIndex = 0
        '
        'cmbExecute
        '
        Me.cmbExecute.BackColor = System.Drawing.SystemColors.ControlDark
        Me.cmbExecute.Location = New System.Drawing.Point(1200, 38)
        Me.cmbExecute.Name = "cmbExecute"
        Me.cmbExecute.Size = New System.Drawing.Size(81, 32)
        Me.cmbExecute.TabIndex = 1
        Me.cmbExecute.Text = "EXECUTE"
        Me.cmbExecute.UseVisualStyleBackColor = False
        '
        'cmbSelect
        '
        Me.cmbSelect.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSelect.FormattingEnabled = True
        Me.cmbSelect.Items.AddRange(New Object() {"Select", "Delete"})
        Me.cmbSelect.Location = New System.Drawing.Point(12, 43)
        Me.cmbSelect.Name = "cmbSelect"
        Me.cmbSelect.Size = New System.Drawing.Size(105, 24)
        Me.cmbSelect.TabIndex = 2
        '
        'cmbPARM1
        '
        Me.cmbPARM1.AutoCompleteCustomSource.AddRange(New String() {"", "DevType", "DevName", "Manfacturer"})
        Me.cmbPARM1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPARM1.FormattingEnabled = True
        Me.cmbPARM1.Items.AddRange(New Object() {"", "ID", "JobNumber", "PartNumber"})
        Me.cmbPARM1.Location = New System.Drawing.Point(332, 48)
        Me.cmbPARM1.Name = "cmbPARM1"
        Me.cmbPARM1.Size = New System.Drawing.Size(118, 24)
        Me.cmbPARM1.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Snow
        Me.Label1.Location = New System.Drawing.Point(141, 49)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(185, 20)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "* from PortConfig   where"
        '
        'cmbPARM2
        '
        Me.cmbPARM2.AutoCompleteCustomSource.AddRange(New String() {"", "DevType", "DevName", "Manfacturer"})
        Me.cmbPARM2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPARM2.FormattingEnabled = True
        Me.cmbPARM2.Items.AddRange(New Object() {"", "ID", "JobNumber", "PartNumber"})
        Me.cmbPARM2.Location = New System.Drawing.Point(629, 46)
        Me.cmbPARM2.Name = "cmbPARM2"
        Me.cmbPARM2.Size = New System.Drawing.Size(119, 24)
        Me.cmbPARM2.TabIndex = 5
        Me.cmbPARM2.Visible = False
        '
        'E1
        '
        Me.E1.AutoSize = True
        Me.E1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.E1.ForeColor = System.Drawing.Color.Snow
        Me.E1.Location = New System.Drawing.Point(461, 48)
        Me.E1.Name = "E1"
        Me.E1.Size = New System.Drawing.Size(18, 20)
        Me.E1.TabIndex = 6
        Me.E1.Text = "="
        Me.E1.Visible = False
        '
        'txtPARM1
        '
        Me.txtPARM1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPARM1.Location = New System.Drawing.Point(485, 45)
        Me.txtPARM1.Name = "txtPARM1"
        Me.txtPARM1.Size = New System.Drawing.Size(88, 26)
        Me.txtPARM1.TabIndex = 7
        Me.txtPARM1.Visible = False
        '
        'A1
        '
        Me.A1.AutoSize = True
        Me.A1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.A1.ForeColor = System.Drawing.Color.Snow
        Me.A1.Location = New System.Drawing.Point(579, 51)
        Me.A1.Name = "A1"
        Me.A1.Size = New System.Drawing.Size(36, 20)
        Me.A1.TabIndex = 8
        Me.A1.Text = "and"
        Me.A1.Visible = False
        '
        'txtPARM2
        '
        Me.txtPARM2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPARM2.Location = New System.Drawing.Point(778, 46)
        Me.txtPARM2.Name = "txtPARM2"
        Me.txtPARM2.Size = New System.Drawing.Size(93, 26)
        Me.txtPARM2.TabIndex = 11
        Me.txtPARM2.Visible = False
        '
        'E2
        '
        Me.E2.AutoSize = True
        Me.E2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.E2.ForeColor = System.Drawing.Color.Snow
        Me.E2.Location = New System.Drawing.Point(754, 49)
        Me.E2.Name = "E2"
        Me.E2.Size = New System.Drawing.Size(18, 20)
        Me.E2.TabIndex = 10
        Me.E2.Text = "="
        Me.E2.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Bradley Hand ITC", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Snow
        Me.Label5.Location = New System.Drawing.Point(537, 2)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(482, 40)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "PORT CONFIGURATION  VIEW"
        '
        'txtPARM3
        '
        Me.txtPARM3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPARM3.Location = New System.Drawing.Point(1068, 44)
        Me.txtPARM3.Name = "txtPARM3"
        Me.txtPARM3.Size = New System.Drawing.Size(93, 26)
        Me.txtPARM3.TabIndex = 52
        Me.txtPARM3.Visible = False
        '
        'E3
        '
        Me.E3.AutoSize = True
        Me.E3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.E3.ForeColor = System.Drawing.Color.Snow
        Me.E3.Location = New System.Drawing.Point(1044, 47)
        Me.E3.Name = "E3"
        Me.E3.Size = New System.Drawing.Size(18, 20)
        Me.E3.TabIndex = 51
        Me.E3.Text = "="
        Me.E3.Visible = False
        '
        'A2
        '
        Me.A2.AutoSize = True
        Me.A2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.A2.ForeColor = System.Drawing.Color.Snow
        Me.A2.Location = New System.Drawing.Point(877, 51)
        Me.A2.Name = "A2"
        Me.A2.Size = New System.Drawing.Size(36, 20)
        Me.A2.TabIndex = 50
        Me.A2.Text = "and"
        Me.A2.Visible = False
        '
        'cmbPARM3
        '
        Me.cmbPARM3.AutoCompleteCustomSource.AddRange(New String() {"", "DevType", "DevName", "Manfacturer"})
        Me.cmbPARM3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPARM3.FormattingEnabled = True
        Me.cmbPARM3.Items.AddRange(New Object() {"", "ID", "JobNumber", "PartNumber"})
        Me.cmbPARM3.Location = New System.Drawing.Point(919, 46)
        Me.cmbPARM3.Name = "cmbPARM3"
        Me.cmbPARM3.Size = New System.Drawing.Size(119, 24)
        Me.cmbPARM3.TabIndex = 49
        Me.cmbPARM3.Visible = False
        '
        'PortConfigView
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(1323, 587)
        Me.Controls.Add(Me.txtPARM3)
        Me.Controls.Add(Me.E3)
        Me.Controls.Add(Me.A2)
        Me.Controls.Add(Me.cmbPARM3)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtPARM2)
        Me.Controls.Add(Me.E2)
        Me.Controls.Add(Me.A1)
        Me.Controls.Add(Me.txtPARM1)
        Me.Controls.Add(Me.E1)
        Me.Controls.Add(Me.cmbPARM2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbPARM1)
        Me.Controls.Add(Me.cmbSelect)
        Me.Controls.Add(Me.cmbExecute)
        Me.Controls.Add(Me.DataGridView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.Name = "PortConfigView"
        Me.Text = "Port Configurtions View"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents cmbExecute As System.Windows.Forms.Button
    Friend WithEvents cmbSelect As System.Windows.Forms.ComboBox
    Friend WithEvents cmbPARM1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbPARM2 As System.Windows.Forms.ComboBox
    Friend WithEvents E1 As System.Windows.Forms.Label
    Friend WithEvents txtPARM1 As System.Windows.Forms.TextBox
    Friend WithEvents A1 As System.Windows.Forms.Label
    Friend WithEvents txtPARM2 As System.Windows.Forms.TextBox
    Friend WithEvents E2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPARM3 As System.Windows.Forms.TextBox
    Friend WithEvents E3 As System.Windows.Forms.Label
    Friend WithEvents A2 As System.Windows.Forms.Label
    Friend WithEvents cmbPARM3 As System.Windows.Forms.ComboBox
End Class
